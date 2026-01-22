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
    /// Command to calculate steel column shortening (PL/AE) with basic lift information.
    /// 
    /// UNIT SYSTEM:
    /// - Forces: API value × ConversionFactor → kips
    /// - Lengths: API returns mm → converted to inches for formula
    /// - Areas: API always returns mm² → converted to in² for formula
    /// - Modulus: Hard-coded 29,000 ksi (steel only)
    /// 
    /// Formula: Shortening (inches) = P(kips) × L(in) / [A(in²) × E(ksi)]
    /// </summary>
    public class SteelColumnShortening : SolverInterrogator
    {
        private const double STEEL_MODULUS_E = 29000.0; // ksi - constant for all steel
        private const double MM_TO_INCHES = 0.0393701;  // 1 mm = 0.0393701 inches
        private const double MM2_TO_IN2 = 0.00155;      // 1 mm² = 0.00155 in²
        private const double MM_TO_FEET = 0.00328084;   // 1 mm = 0.00328084 feet

        /// <summary>
        /// Determines if this command should be shown in the menu.
        /// </summary>
        public override bool ShowInMenu() => true;

        /// <summary>
        /// Constructor sets up output and requested member type.
        /// </summary>
        public SteelColumnShortening()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        /// <summary>
        /// Represents span-level shortening calculation details.
        /// </summary>
        private class SpanShorteningDetail
        {
            public double Force { get; set; }
            public double LengthFt { get; set; }
            public double Shortening { get; set; }
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
        /// Gets the cross-sectional area in square inches.
        /// API always returns area in mm², so convert to in².
        /// </summary>
        private async Task<double> GetCrossSectionalArea(ColumnLift lift)
        {
            try
            {
                var firstSpan = lift.Spans.First();
                if (firstSpan.ElementSection.Value is IMemberSection elementSection)
                {
                    if (elementSection.PhysicalSection.Value is ISection physicalSection)
                    {
                        // API returns area in mm², convert to in²
                        double areaMm2 = physicalSection.CrossSectionalArea;
                        double areaIn2 = areaMm2 * MM2_TO_IN2;

                        return areaIn2 > 0 ? areaIn2 : 0.0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error getting cross-sectional area for lift {lift.Name}: {ex.Message}");
            }

            return 0.0;
        }

        /// <summary>
        /// Calculates total column shortening for a lift by summing shortening of each span.
        /// Gets force at start of each span and calculates individual span shortening.
        /// Returns total shortening and details for each span.
        /// </summary>
        private async Task<(double TotalShortening, List<SpanShorteningDetail> SpanDetails)> CalculateLiftShorteningWithDetails(IMember member, ColumnLift lift, ILoadingCase loadingCase, bool reduced, double area)
        {
            var spanDetails = new List<SpanShorteningDetail>();
            if (area <= 0) return (0.0, spanDetails);

            try
            {
                IMemberLoading memberLoading = await member.GetLoadingAsync(loadingCase.Id, RequestedAnalysisType, LoadingResultType.Base);
                double valCon = ConversionFactor(LoadingValueType.Force);

                double totalShortening = 0.0;

                // Calculate shortening for each span in the lift
                foreach (var span in lift.Spans)
                {
                    // Get force at start of span (force entering this span)
                    double spanForce = await GetLoadingValueAtPosition(
                        memberLoading, LoadingValueType.Force, LoadingDirection.Axial, 0.0, reduced, span.Index) * valCon;

                    // Convert span length from mm to inches and feet
                    double spanLengthInches = span.Length.Value * MM_TO_INCHES;
                    double spanLengthFt = span.Length.Value * MM_TO_FEET;

                    // Calculate shortening for this span: (P × L) / (A × E)
                    double spanShortening = (Math.Abs(spanForce) * spanLengthInches) / (area * STEEL_MODULUS_E);

                    spanDetails.Add(new SpanShorteningDetail
                    {
                        Force = Math.Abs(spanForce),
                        LengthFt = spanLengthFt,
                        Shortening = spanShortening
                    });

                    totalShortening += spanShortening;
                }

                return (totalShortening, spanDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error calculating span-by-span shortening for lift {lift.Name}: {ex.Message}");
                return (0.0, spanDetails);
            }
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
            values = values.OrderByDescending(lv => Math.Abs(lv.Value));

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

            // Find maximum number of spans across all lifts to determine column count
            int maxSpanCount = 0;
            foreach (var columnSpans in steelColumnSpans)
            {
                var lifts = columnSpans.CreateLifts();
                foreach (var lift in lifts)
                {
                    maxSpanCount = Math.Max(maxSpanCount, lift.Spans.Count);
                }
            }

            // Prepare output CSV file
            string file1 = SaveDirectory + @"SteelColumnShortening_" + OutputFileName + ".csv";

            // Build dynamic header with span columns based on actual max spans found
            var headerParts = new List<string>
            {
                "Tekla GUID",
                "Part Mark",
                "UDA Filter",
                "Member Name",
                "Lift Name",
                "Start Level",
                "End Level",
                "Shape",
                "Material",
                "Start Node",
                "X_StartNode",
                "Y_StartNode",
                "Z_StartNode",
                "End Node",
                "X_EndNode",
                "Y_EndNode",
                "Z_EndNode",
                "Lift Length [ft]",
                "Cross Section Area [in²]",
                "Shortening [in]"
            };

            // Add span columns based on actual max spans in model
            for (int i = 1; i <= maxSpanCount; i++)
            {
                headerParts.Add($"Force[k] Span {i}");
                headerParts.Add($"Length[ft] Span {i}");
                headerParts.Add($"Shortening[in] Span {i}");
            }

            string header1 = string.Join(",", headerParts) + "\n";

            File.WriteAllText(file1, header1);

            FancyWriteLine("Calculating Column Shortening...", TextColor.Title);

            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
            {
                // Use the first loading case for shortening calculations
                var loadingCase = loadingCases.First();

                foreach (var columnSpans in steelColumnSpans)
                {
                    var member = columnSpans.ParentMember;
                    string memberName = member.Name;
                    var lifts = columnSpans.CreateLifts();

                    for (int liftIndex = 0; liftIndex < lifts.Count; liftIndex++)
                    {
                        var lift = lifts[liftIndex];

                        var firstSpanforID = lift.Spans.First();
                        Guid id = firstSpanforID.Id;
                        string partMark = lifts.Count == 1 ? member.Name : firstSpanforID.Name;

                        // Get start and end node construction points for the lift
                        int startNodeIdx = lift.StartNode.ConstructionPointIndex.Value;
                        var startPoints = await Model.GetConstructionPointsAsync(new List<int> { startNodeIdx });
                        var startPoint = startPoints.First();
                        double startX = startPoint.Coordinates.Value.X * MM_TO_FEET;
                        double startY = startPoint.Coordinates.Value.Y * MM_TO_FEET;
                        double startZ = startPoint.Coordinates.Value.Z * MM_TO_FEET;

                        var startPlaneIds = startPoints
                            .Where(p => p.PlaneInfo.Value.Type == TSD.API.Remoting.Common.EntityType.HorizontalConstructionPlane)
                            .Select(p => p.PlaneInfo.Value.Index);
                        string startLevelName = startPlaneIds.Any()
                            ? (await Model.GetLevelsAsync(startPlaneIds)).First().Name
                            : $"~{levels.OrderBy(l => Math.Abs(startPoint.Coordinates.Value.Z - l.Level.Value)).First().Name}";

                        int endNodeIdx = lift.EndNode.ConstructionPointIndex.Value;
                        var endPoints = await Model.GetConstructionPointsAsync(new List<int> { endNodeIdx });
                        var endPoint = endPoints.First();
                        double endX = endPoint.Coordinates.Value.X * MM_TO_FEET;
                        double endY = endPoint.Coordinates.Value.Y * MM_TO_FEET;
                        double endZ = endPoint.Coordinates.Value.Z * MM_TO_FEET;

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

                        double lengthFt = lift.Length * MM_TO_FEET;
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

                        // Calculate shortening with correct force logic
                        double area = await GetCrossSectionalArea(lift);
                        var (shortening, spanDetails) = await CalculateLiftShorteningWithDetails(member, lift, loadingCase, reduced, area);

                        // Build base line with lift data
                        string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(memberName)},{EscapeCsvValue(lift.Name)},{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)}," +
                                      $"{EscapeCsvValue(startNodeName)},{startX:F3},{startY:F3},{startZ:F3}," +
                                      $"{EscapeCsvValue(endNodeName)},{endX:F3},{endY:F3},{endZ:F3}," +
                                      $"{lengthFt:F3},{area:F3},{shortening:F4}";

                        // Add span data columns
                        for (int spanIdx = 0; spanIdx < maxSpanCount; spanIdx++)
                        {
                            if (spanIdx < spanDetails.Count)
                            {
                                var span = spanDetails[spanIdx];
                                line += $",{span.Force:F1},{span.LengthFt:F2},{span.Shortening:F4}";
                            }
                            else
                            {
                                // No data for this span, add zeros
                                line += ",0,0,0";
                            }
                        }

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