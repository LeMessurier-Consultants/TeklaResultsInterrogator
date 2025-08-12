using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TSD.API.Remoting.Structure;
using TeklaResultsInterrogator.Utils;

public class ColumnSpansSteel
{
    public IMember ParentMember { get; }
    public List<IMemberSpan> Spans { get; private set; } = new();
    public bool HasSplice { get; private set; }

    // Store splice info per span index for reporting or output
    public Dictionary<int, (bool HasSplice, double SpliceOffset)> SpanSpliceInfo { get; } = new();

    public ColumnSpansSteel(IMember parentMember)
    {
        ParentMember = parentMember;
    }

    public async Task OrganizeSpansAsync()
    {
        var spans = (await ParentMember.GetSpanAsync())
            .OrderBy(s => s.Index)
            .ToList();

        HasSplice = false;
        SpanSpliceInfo.Clear();

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
                HasSplice = true;
        }

        Spans = spans;
    }
}
