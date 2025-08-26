using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TeklaResultsInterrogator.Core;
using TeklaResultsInterrogator.Utils;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Sections;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;
using TSD.API.Remoting.UserDefinedAttributes;
using static TeklaResultsInterrogator.Utils.Utils;

namespace TeklaResultsInterrogator.Commands
{
    /// <summary>
    /// Simplified command to generate a summary of steel columns with basic lift information.
    /// </summary>
    public class SteelColumnIntegrityForces : SolverInterrogator
    {
        /// <summary>
        /// Determines if this command should be shown in the menu.
        /// </summary>
        public override bool ShowInMenu() => true;

        /// <summary>
        /// Constructor sets up output and requested member type.
        /// </summary>
        public SteelColumnIntegrityForces()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        /// <summary>
        /// Organizes steel column spans and detects splices.
        /// </summary>
        public class ColumnSpansSteel
        {
            public IMember ParentMember { get; }
            public List<IMemberSpan> Spans { get; private set; } = new();
            public bool HasSplice { get; private set; }
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

                // Check for explicit splice data in stack data
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

                // Detect splices from span names and UDAs
                await DetectSplicesFromNamesAsync(spans);

                // Detect splices from gaps in span numbering
                DetectSplicesFromSpanNumbers(spans);

                Spans = spans;
            }

            private async Task DetectSplicesFromNamesAsync(List<IMemberSpan> spans)
            {
                var spliceKeywords = new[] { "splice", "connection", "joint", "lift", "piece", "stack" };
                foreach (var span in spans)
                {
                    try
                    {
                        bool nameIndicatesSplice = false;
                        double nameBasedSpliceOffset = span.Length.Value;

                        string spanName = span.Name?.ToLowerInvariant() ?? "";
                        if (spliceKeywords.Any(keyword => spanName.Contains(keyword)))
                        {
                            nameIndicatesSplice = true;
                        }

                        var udas = await span.GetUserDefinedAttributesAsync();
                        foreach (var uda in udas)
                        {
                            if (uda is IUserDefinedTextAttribute textUda)
                            {
                                string udaValue = textUda.Text?.ToLowerInvariant() ?? "";
                                string udaName = uda.AttributeDefinitionName?.ToLowerInvariant() ?? "";
                                if (spliceKeywords.Any(keyword => udaValue.Contains(keyword) || udaName.Contains(keyword)))
                                {
                                    nameIndicatesSplice = true;
                                    if (uda.AttributeDefinitionName?.ToLowerInvariant().Contains("offset") == true ||
                                        uda.AttributeDefinitionName?.ToLowerInvariant().Contains("distance") == true)
                                    {
                                        if (double.TryParse(textUda.Text, out double offsetValue))
                                        {
                                            nameBasedSpliceOffset = offsetValue;
                                        }
                                    }
                                }
                            }
                        }

                        if (nameIndicatesSplice)
                        {
                            HasSplice = true;
                            var existingInfo = SpanSpliceInfo.TryGetValue(span.Index, out var existing)
                                ? existing
                                : (false, 0.0);
                            double finalOffset = existingInfo.Item1 ? existingInfo.Item2 : nameBasedSpliceOffset;
                            SpanSpliceInfo[span.Index] = (true, finalOffset);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Error checking span {span.Index} for name-based splice indicators: {ex.Message}");
                    }
                }
            }

            private void DetectSplicesFromSpanNumbers(List<IMemberSpan> spans)
            {
                if (spans.Count <= 1) return;
                var spanIndices = spans.Select(s => s.Index).OrderBy(i => i).ToList();
                for (int i = 0; i < spanIndices.Count - 1; i++)
                {
                    int currentSpan = spanIndices[i];
                    int nextSpan = spanIndices[i + 1];
                    if (nextSpan - currentSpan > 1)
                    {
                        HasSplice = true;
                        var spanBeforeGap = spans.FirstOrDefault(s => s.Index == currentSpan);
                        if (spanBeforeGap != null)
                        {
                            var existingInfo = SpanSpliceInfo.TryGetValue(spanBeforeGap.Index, out var existing)
                                ? existing
                                : (false, spanBeforeGap.Length.Value);
                            double finalOffset = existingInfo.Item1 ? existingInfo.Item2 : spanBeforeGap.Length.Value;
                            SpanSpliceInfo[spanBeforeGap.Index] = (true, finalOffset);
                        }
                    }
                }
            }

            public List<ColumnLift> CreateLifts()
            {
                var lifts = new List<ColumnLift>();
                if (!HasSplice)
                {
                    var lift = new ColumnLift
                    {
                        Name = $"{ParentMember.Name}_L1",
                        Spans = new List<IMemberSpan>(Spans),
                        StartNode = Spans.First().StartMemberNode,
                        EndNode = Spans.Last().EndMemberNode,
                        Length = Spans.Sum(s => s.Length.Value)
                    };
                    lifts.Add(lift);
                    return lifts;
                }

                var currentLiftSpans = new List<IMemberSpan>();
                var liftCount = 1;
                var spliceSpanIndices = SpanSpliceInfo
                    .Where(si => si.Value.Item1)
                    .Select(si => si.Key)
                    .OrderBy(idx => idx)
                    .ToList();

                foreach (var span in Spans.OrderBy(s => s.Index))
                {
                    if (spliceSpanIndices.Contains(span.Index) && currentLiftSpans.Any())
                    {
                        var lift = new ColumnLift
                        {
                            Name = $"{ParentMember.Name}_L{liftCount}",
                            Spans = new List<IMemberSpan>(currentLiftSpans),
                            StartNode = currentLiftSpans.First().StartMemberNode,
                            EndNode = currentLiftSpans.Last().EndMemberNode,
                            Length = currentLiftSpans.Sum(s => s.Length.Value)
                        };
                        lifts.Add(lift);
                        currentLiftSpans.Clear();
                        liftCount++;
                    }
                    currentLiftSpans.Add(span);
                }

                if (currentLiftSpans.Any())
                {
                    var lift = new ColumnLift
                    {
                        Name = $"{ParentMember.Name}_L{liftCount}",
                        Spans = new List<IMemberSpan>(currentLiftSpans),
                        StartNode = currentLiftSpans.First().StartMemberNode,
                        EndNode = currentLiftSpans.Last().EndMemberNode,
                        Length = currentLiftSpans.Sum(s => s.Length.Value)
                    };
                    lifts.Add(lift);
                }

                return lifts;
            }
        }

        /// <summary>
        /// Represents a column lift (segment between splices).
        /// </summary>
        public class ColumnLift
        {
            public string Name { get; set; }
            public List<IMemberSpan> Spans { get; set; } = new List<IMemberSpan>();
            public IMemberNode StartNode { get; set; }
            public IMemberNode EndNode { get; set; }
            public double Length { get; set; }
        }

        /// <summary>
        /// Calculates integrity forces for each lift using splice offsets.
        /// </summary>
        private async Task<Dictionary<string, double>> CalculateIntegrityForcesWithSpliceOffsets(
            IMember member,
            List<ColumnLift> lifts,
            ILoadingCase integrityForceCase,
            bool reduced,
            ColumnSpansSteel columnSpans)
        {
            var integrityForces = new Dictionary<string, double>();
            if (lifts.Count <= 1)
            {
                return integrityForces;
            }

            try
            {
                IMemberLoading memberLoading = await member.GetLoadingAsync(integrityForceCase.Id, RequestedAnalysisType, LoadingResultType.Base);
                double valCon = ConversionFactor(LoadingValueType.Force);

                integrityForces[lifts[0].Name] = 0.0;

                double startNodeForce = await GetForceAtLiftStart(memberLoading, lifts[0], reduced) * valCon;

                var spliceForces = new List<double>();
                foreach (var lift in lifts)
                {
                    foreach (var span in lift.Spans)
                    {
                        if (columnSpans.SpanSpliceInfo.ContainsKey(span.Index) &&
                            columnSpans.SpanSpliceInfo[span.Index].HasSplice)
                        {
                            double spliceOffset = columnSpans.SpanSpliceInfo[span.Index].SpliceOffset;
                            double spliceForce = await GetLoadingValueAtPosition(
                                memberLoading,
                                LoadingValueType.Force,
                                LoadingDirection.Axial,
                                spliceOffset,
                                reduced,
                                span.Index) * valCon;
                            spliceForces.Add(spliceForce);
                        }
                    }
                }

                for (int i = 1; i < lifts.Count; i++)
                {
                    if (i == 1 && spliceForces.Count > 0)
                    {
                        integrityForces[lifts[i].Name] = startNodeForce - spliceForces[0];
                    }
                    else if (i - 1 < spliceForces.Count && i - 2 >= 0 && i - 2 < spliceForces.Count)
                    {
                        integrityForces[lifts[i].Name] = spliceForces[i - 2] - spliceForces[i - 1];
                    }
                    else
                    {
                        integrityForces[lifts[i].Name] = 0.0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calculating integrity forces with splice offsets: {ex.Message}");
                foreach (var lift in lifts)
                {
                    integrityForces[lift.Name] = 0.0;
                }
            }

            return integrityForces;
        }

        private async Task<double> GetForceAtLiftStart(IMemberLoading memberLoading, ColumnLift lift, bool reduced)
        {
            var firstSpan = lift.Spans.First();
            double position = 0.0;
            return await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, position, reduced, firstSpan.Index);
        }

        private static async Task<double> GetLoadingValueAtPosition(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            double positionMm,
            bool reduced,
            int spanIndex)
        {
            var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);
            IEnumerable<ILoadingValue> values = await loading.GetValueAsync(option, spanIndex, positionMm);
            values = values.OrderByDescending(lv => lv.Value);

            if (values.Any())
                return values.First().Value;
            return 0.0;
        }

        /// <summary>
        /// Main execution method for the command.
        /// </summary>
        public override async Task ExecuteAsync()
        {
            await InitializeAsync();
            if (Flag) return;

            Stopwatch stopwatch = Stopwatch.StartNew();
            int bufferSize = 65536 * 2;

            FancyWriteLine("Loading Summary:", TextColor.Title);
            Console.WriteLine("Unpacking loading data...");
            Console.WriteLine($"{AllLoadcases!.Count} loadcases found, {SolvedCases!.Count} solved.");
            Console.WriteLine($"{AllCombinations!.Count} load combinations found, {SolvedCombinations!.Count} solved.");
            Console.WriteLine($"{AllEnvelopes!.Count} load envelopes found, {SolvedEnvelopes!.Count} solved.\n");

            var loadingCases = AskLoading(SolvedCases, SolvedCombinations, SolvedEnvelopes);
            bool reduced = AskReduced();

            FancyWriteLine("\nMember summary:", TextColor.Title);
            Console.WriteLine("Unpacking member data...");
            var steelColumns = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction))).ToList();

            string filterField = AskUser("What UDA field to filter on?");
            string filterValue = AskUser("What UDA value to filter on?");

            Console.WriteLine($"{AllMembers.Count} structural members found in model.");
            Console.WriteLine($"{steelColumns.Count} steel columns found.");

            var levels = (await Model!.GetLevelsAsync()).ToList();

            FancyWriteLine("Organizing Column Lifts...", TextColor.Title);
            var steelColumnSpans = new List<ColumnSpansSteel>();

            foreach (var column in steelColumns)
            {
                var colSpans = new ColumnSpansSteel(column);
                await colSpans.OrganizeSpansAsync();
                steelColumnSpans.Add(colSpans);
            }

            // Prepare output CSV file
            string file1 = SaveDirectory + @"SteelColumnIntegrityForces_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material," +
                           "Start Node,X_StartNode,Y_StartNode,Z_StartNode," +
                           "End Node,X_EndNode,Y_EndNode,Z_EndNode," +
                           "Lift Length [ft],Integrity Force [k]\n";

            File.WriteAllText(file1, "");
            File.AppendAllText(file1, header1);

            FancyWriteLine("Writing Integrity Forces...", TextColor.Title);

            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
            {
                var integrityForceCase = loadingCases.FirstOrDefault(lc =>
                    lc.Name.Equals("Integrity Force", StringComparison.CurrentCultureIgnoreCase) ||
                    lc.Name.Contains("Integrity", StringComparison.CurrentCultureIgnoreCase));

                foreach (var columnSpans in steelColumnSpans)
                {
                    var member = columnSpans.ParentMember;
                    string memberName = member.Name;
                    var lifts = columnSpans.CreateLifts();

                    var integrityForces = new Dictionary<string, double>();
                    if (integrityForceCase != null && columnSpans.HasSplice)
                    {
                        integrityForces = await CalculateIntegrityForcesWithSpliceOffsets(member, lifts, integrityForceCase, reduced, columnSpans);
                    }

                    foreach (var lift in lifts)
                    {
                        var firstSpanforID = lift.Spans.First();
                        Guid id = firstSpanforID.Id;
                        string partMark = lifts.Count == 1 ? member.Name : firstSpanforID.Name;

                        // Get start and end node construction points for the lift
                        int startNodeIdx = lift.StartNode.ConstructionPointIndex.Value;
                        var startPoints = await Model.GetConstructionPointsAsync(new List<int> { startNodeIdx });
                        var startPoint = startPoints.First();
                        double startX = startPoint.Coordinates.Value.X * 0.00328084;
                        double startY = startPoint.Coordinates.Value.Y * 0.00328084;
                        double startZ = startPoint.Coordinates.Value.Z * 0.00328084;

                        var startPlaneIds = startPoints
                            .Where(p => p.PlaneInfo.Value.Type == TSD.API.Remoting.Common.EntityType.HorizontalConstructionPlane)
                            .Select(p => p.PlaneInfo.Value.Index);
                        string startLevelName = startPlaneIds.Any()
                            ? (await Model.GetLevelsAsync(startPlaneIds)).First().Name
                            : $"~{levels.OrderBy(l => Math.Abs(startPoint.Coordinates.Value.Z - l.Level.Value)).First().Name}";

                        int endNodeIdx = lift.EndNode.ConstructionPointIndex.Value;
                        var endPoints = await Model.GetConstructionPointsAsync(new List<int> { endNodeIdx });
                        var endPoint = endPoints.First();
                        double endX = endPoint.Coordinates.Value.X * 0.00328084;
                        double endY = endPoint.Coordinates.Value.Y * 0.00328084;
                        double endZ = endPoint.Coordinates.Value.Z * 0.00328084;

                        var endPlaneIds = endPoints
                            .Where(p => p.PlaneInfo.Value.Type == TSD.API.Remoting.Common.EntityType.HorizontalConstructionPlane)
                            .Select(p => p.PlaneInfo.Value.Index);
                        string endLevelName = endPlaneIds.Any()
                            ? (await Model.GetLevelsAsync(endPlaneIds)).First().Name
                            : $"~{levels.OrderBy(l => Math.Abs(endPoint.Coordinates.Value.Z - l.Level.Value)).First().Name}";

                        // Get section and material info from the first span in the lift
                        var firstSpan = lift.Spans.First();
                        string sectionName = "Unknown";
                        string materialName = "Unknown";
                        if (firstSpan.ElementSection.Value != null)
                        {
                            var elementSection = (IMemberSection)firstSpan.ElementSection.Value;
                            var physicalSection = (ISection)elementSection.PhysicalSection.Value;
                            sectionName = physicalSection.LongName;
                        }
                        if (firstSpan.Material?.Value != null)
                        {
                            materialName = firstSpan.Material.Value.Name;
                        }

                        double lengthFt = lift.Length * 0.00328084;
                        string startNodeName = $"{startNodeIdx}";
                        string endNodeName = $"{endNodeIdx}";

                        // Check UDA filter for this lift (check first span)
                        var udas = await firstSpan.GetUserDefinedAttributesAsync();
                        if (!string.IsNullOrEmpty(filterValue))
                        {
                            bool match = udas.Any(c =>
                                (c as IUserDefinedTextAttribute)?.Text.Equals(filterValue, StringComparison.CurrentCultureIgnoreCase) == true &&
                                c?.AttributeDefinitionName.Equals(filterField, StringComparison.CurrentCultureIgnoreCase) == true);
                            if (!match) continue;
                        }

                        double integrityForce = 0.0;
                        if (integrityForceCase != null && integrityForces.ContainsKey(lift.Name))
                        {
                            integrityForce = integrityForces[lift.Name];
                        }

                        // Write CSV line for this lift
                        string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(memberName)},{EscapeCsvValue(lift.Name)},{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)}," +
                                    $"{EscapeCsvValue(startNodeName)},{startX:F3},{startY:F3},{startZ:F3}," +
                                    $"{EscapeCsvValue(endNodeName)},{endX:F3},{endY:F3},{endZ:F3}," +
                                    $"{lengthFt:F3},{integrityForce}";

                        sw1.WriteLine(line);
                    }
                }
            }

            FancyWriteLine("Saved to: ", file1, "", TextColor.Path);
            double sizeKB = Math.Round(new FileInfo(file1).Length / 1024.0, 2);
            Console.WriteLine($"File size: {sizeKB} KB");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;
            Check();
        }

        /// <summary>
        /// Escapes CSV values: adds quotes if contains comma, quotes or line breaks, and doubles any embedded quotes.
        /// </summary>
        private static string EscapeCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }
            return value;
        }
    }
}