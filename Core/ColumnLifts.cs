using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TeklaResultsInterrogator.Utils;
using TSD.API.Remoting.Structure;

public class ColumnLifts
{
    public IMember ParentMember { get; }
    public List<NamedList<IMemberSpan>> Lifts { get; } = new();
    public bool HasSplice { get; private set; }

    // Store splice info per span index for reporting or output
    public Dictionary<int, (bool HasSplice, double SpliceOffset)> SpanSpliceInfo { get; } = new();

    public ColumnLifts(IMember parentMember)
    {
        ParentMember = parentMember;
    }

    public async Task OrganizeBySpliceAsync()
    {
        var spans = (await ParentMember.GetSpanAsync())
            .OrderBy(s => s.Index)
            .ToList();

        HasSplice = false;
        SpanSpliceInfo.Clear();

        // Collect splice info and track which spans have splices
        var spliceIndices = new HashSet<int>();

        foreach (var span in spans)
        {
            bool spanHasSplice = false;
            double spliceOffset = 0;

            if (span.Data?.Value is ISteelColumnStackData stackData)
            {
                if (stackData.HasSplice.IsApplicable)
                    spanHasSplice = stackData.HasSplice.Value;

                if (stackData.SpliceOffset.IsApplicable)
                    spliceOffset = stackData.SpliceOffset.Value;
            }

            SpanSpliceInfo[span.Index] = (spanHasSplice, spliceOffset);
            if (spanHasSplice)
            {
                HasSplice = true;
                spliceIndices.Add(span.Index);
            }
        }

        Lifts.Clear();

        if (!HasSplice)
        {
            // No splices → single lift with all spans
            var singleLift = new NamedList<IMemberSpan>("L1");
            foreach (var span in spans)
                singleLift.Add(span);
            Lifts.Add(singleLift);
            return;
        }

        // Sort splice indices to process in order
        var sortedSplices = spliceIndices.OrderBy(i => i).ToList();

        // We will create lifts starting at:
        // - first span index
        // - every splice index (except the first if it is the very first span)
        // Each lift ends *just before* the next splice start span

        var liftStartIndices = new List<int> { spans.First().Index };
        // Add splice starts after the first span if not already first
        foreach (var spliceIdx in sortedSplices)
        {
            if (spliceIdx != spans.First().Index)
                liftStartIndices.Add(spliceIdx);
        }
        liftStartIndices = liftStartIndices.OrderBy(i => i).ToList();

        for (int i = 0; i < liftStartIndices.Count; i++)
        {
            int startIdx = liftStartIndices[i];
            int endIdxExclusive = (i + 1 < liftStartIndices.Count) ? liftStartIndices[i + 1] : spans.Last().Index + 1;

            var lift = new NamedList<IMemberSpan>($"L{i + 1}");

            // Add spans with index >= startIdx and < endIdxExclusive
            foreach (var span in spans)
            {
                if (span.Index >= startIdx && span.Index < endIdxExclusive)
                    lift.Add(span);
            }

            Lifts.Add(lift);
        }
    }

}
