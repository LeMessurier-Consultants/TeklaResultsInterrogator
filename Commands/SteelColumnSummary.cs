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
    /// Command to generate a summary of steel columns, including splice detection and lift organization.
    /// </summary>
    public class SteelColumnSummary : SolverInterrogator
    {
        /// <summary>
        /// Determines if this command should be shown in the menu.
        /// </summary>
        public override bool ShowInMenu() => true;

        /// <summary>
        /// Constructor sets up output and requested member type.
        /// </summary>
        public SteelColumnSummary()
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
            // Stores splice info for each span index: (HasSplice, SpliceOffset)
            public Dictionary<int, (bool HasSplice, double SpliceOffset)> SpanSpliceInfo { get; } = new();

            /// <summary>
            /// Constructor for ColumnSpansSteel.
            /// </summary>
            public ColumnSpansSteel(IMember parentMember)
            {
                ParentMember = parentMember;
            }

            /// <summary>
            /// Organizes spans and detects splices using multiple methods.
            /// </summary>
            public async Task OrganizeSpansAsync()
            {
                var spans = (await ParentMember.GetSpanAsync())
                    .OrderBy(s => s.Index)
                    .ToList();
                HasSplice = false;
                SpanSpliceInfo.Clear();

                // First pass: Check for explicit splice data in stack data.
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

                // Second pass: Detect splices from span names and UDAs.
                await DetectSplicesFromNamesAsync(spans);

                // Third pass: Detect splices from gaps in span numbering.
                DetectSplicesFromSpanNumbers(spans);

                Spans = spans;
            }

            /// <summary>
            /// Detects splices based on keywords in span names and UDAs.
            /// </summary>
            private async Task DetectSplicesFromNamesAsync(List<IMemberSpan> spans)
            {
                var spliceKeywords = new[] { "splice", "connection", "joint", "lift", "piece", "stack" };
                foreach (var span in spans)
                {
                    try
                    {
                        bool nameIndicatesSplice = false;
                        double nameBasedSpliceOffset = span.Length.Value; // Default offset is end of span

                        // Check span name for splice keywords
                        string spanName = span.Name?.ToLowerInvariant() ?? "";
                        if (spliceKeywords.Any(keyword => spanName.Contains(keyword)))
                        {
                            nameIndicatesSplice = true;
                        }

                        // Check UDAs for splice-related information
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
                                    // Try to extract splice offset from UDA if numeric
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

                        // If detected, update splice info
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

            /// <summary>
            /// Detects splices based on gaps in span numbering.
            /// </summary>
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

            /// <summary>
            /// Creates lifts (column segments between splices) from organized spans.
            /// </summary>
            public List<ColumnLift> CreateLifts()
            {
                var lifts = new List<ColumnLift>();
                if (!HasSplice)
                {
                    // No splices: one lift for the whole column
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

                // Splices exist: create lifts between splices
                var currentLiftSpans = new List<IMemberSpan>();
                var liftCount = 1;
                var spliceSpanIndices = SpanSpliceInfo
                    .Where(si => si.Value.Item1)
                    .Select(si => si.Key)
                    .OrderBy(idx => idx)
                    .ToList();

                foreach (var span in Spans.OrderBy(s => s.Index))
                {
                    // If this span starts with a splice, close previous lift
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

                // Add final lift if any spans remain
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
            public double Length { get; set; } // Total length in mm
        }

        /// <summary>
        /// Gets the min and max eccentric moment in a lift for a given direction.
        /// </summary>
        private static async Task<(double max, double min)> GetMinMaxEccentricMomentInLift(
            IMemberLoading loading,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            double maxValue = double.MinValue;
            double minValue = double.MaxValue;
            const double Nmm_to_kft = 0.000000737562149277; // Conversion factor

            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;
                for (int i = 0; i <= samplePoints; i++)
                {
                    double position = (i * spanLength) / samplePoints;
                    var option = LoadingValueOptions.StaticValue(LoadingValueType.EccentricityMoment, direction, reduced);
                    IEnumerable<ILoadingValue> values = await loading.GetValueAsync(option, span.Index, position);

                    foreach (var lv in values)
                    {
                        double value = lv.Value * Nmm_to_kft;
                        if (value > maxValue) maxValue = value;
                        if (value < minValue) minValue = value;
                    }
                }
            }

            // Fallback if no data
            if (maxValue == double.MinValue) maxValue = 0.0;
            if (minValue == double.MaxValue) minValue = 0.0;

            return (maxValue, minValue);
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

                // First lift always has zero integrity force
                integrityForces[lifts[0].Name] = 0.0;

                // Get start node force
                double startNodeForce = await GetForceAtLiftStart(memberLoading, lifts[0], reduced) * valCon;

                // Get forces at all splice locations
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

                // Calculate integrity forces for each lift
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

        /// <summary>
        /// Gets the force at the start of a lift.
        /// </summary>
        private async Task<double> GetForceAtLiftStart(IMemberLoading memberLoading, ColumnLift lift, bool reduced)
        {
            var firstSpan = lift.Spans.First();
            double position = 0.0;
            return await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, position, reduced, firstSpan.Index);
        }

        /// <summary>
        /// Gets the force at the end of a lift.
        /// </summary>
        private async Task<double> GetForceAtLiftEnd(IMemberLoading memberLoading, ColumnLift lift, bool reduced)
        {
            var lastSpan = lift.Spans.Last();
            double position = lastSpan.Length.Value;
            return await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, position, reduced, lastSpan.Index);
        }

        /// <summary>
        /// Finds which span contains a given cumulative length and the relative position within that span.
        /// </summary>
        private static (int spanIndex, double relativePosition) GetSpanAndPositionFromCumulativeLength(List<ColumnLift> lifts, double cumulativeLength)
        {
            double currentLength = 0.0;
            foreach (var lift in lifts)
            {
                foreach (var span in lift.Spans)
                {
                    if (currentLength + span.Length.Value >= cumulativeLength)
                    {
                        double relativePosition = cumulativeLength - currentLength;
                        return (span.Index, relativePosition);
                    }
                    currentLength += span.Length.Value;
                }
            }
            // Fallback to last span
            var lastLift = lifts.Last();
            var lastSpan = lastLift.Spans.Last();
            return (lastSpan.Index, lastSpan.Length.Value);
        }

        /// <summary>
        /// Gets a loading value at a specific position in a span.
        /// </summary>
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

            stopwatch.Stop();
            var loadingCases = AskLoading(SolvedCases, SolvedCombinations, SolvedEnvelopes);
            bool reduced = AskReduced();
            stopwatch.Start();

            // Member Data
            FancyWriteLine("\nMember summary:", TextColor.Title);
            Console.WriteLine("Unpacking member data...");
            var steelColumns = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction))).ToList();

            string filterField = AskUser("What UDA field to filter on?");
            string filterValue = AskUser("What UDA value to filter on?");

            Console.WriteLine($"{AllMembers.Count} structural members found in model.");
            Console.WriteLine($"{steelColumns.Count} steel columns found.");

            var levels = (await Model!.GetLevelsAsync()).ToList();
            double timeUnpack = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Loading and member data unpacked in {timeUnpack} seconds.\n");

            FancyWriteLine("Organizing Column Lifts...", TextColor.Title);
            var steelColumnSpans = new List<ColumnSpansSteel>();

            // For each steel column, organize spans and lifts, and report splice info
            foreach (var column in steelColumns)
            {
                var colSpans = new ColumnSpansSteel(column);
                await colSpans.OrganizeSpansAsync();

                Console.WriteLine($"Column {column.Name}:");
                Console.WriteLine($"  -> Has splice? {(colSpans.HasSplice ? "Yes" : "No")}");
                Console.WriteLine($"  -> {colSpans.Spans.Count} spans total");

                var lifts = colSpans.CreateLifts();
                Console.WriteLine($"  -> {lifts.Count} lifts created");

                // Debug output for splice detection
                if (colSpans.HasSplice)
                {
                    Console.WriteLine($"  -> Splice locations:");
                    foreach (var spliceInfo in colSpans.SpanSpliceInfo.Where(si => si.Value.Item1))
                    {
                        var span = colSpans.Spans.FirstOrDefault(s => s.Index == spliceInfo.Key);
                        string spanName = span?.Name ?? "Unknown";
                        double offsetInches = spliceInfo.Value.Item2 / 25.4; // mm to inches
                        Console.WriteLine($"     Span {spliceInfo.Key} ({spanName}): splice at {offsetInches:F1}\" from start");
                    }

                    // Show lift breakdown
                    Console.WriteLine($"  -> Lift breakdown:");
                    for (int i = 0; i < lifts.Count; i++)
                    {
                        var lift = lifts[i];
                        double lengthFt = lift.Length * 0.00328084;
                        var spanNames = string.Join(", ", lift.Spans.Select(s => s.Name));
                        Console.WriteLine($"     {lift.Name}: {lengthFt:F1}ft ({spanNames})");
                    }
                }

                steelColumnSpans.Add(colSpans);
            }

            double endWatch = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Column spans organized in {Math.Round(endWatch - timeUnpack, 3)} seconds.\n");

            // Prepare output CSV file
            string file1 = SaveDirectory + @"SteelColumnSummary_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material," +
                             "Start Node,Start Node Fixity,X_StartNode,Y_StartNode,Z_StartNode," +
                             "End Node,End Node Fixity,X_EndNode,Y_EndNode,Z_EndNode," +
                             "Lift Length [ft],Span Rotation [deg],Loading Name," +
                             "Column Axial Max [k],Column Axial Min [k],Column Major Moment Max [k-ft]," +
                             "Column Major Shear Max [k],Column Minor Moment Max [k-ft],Column Minor Shear Max [k]," +
                             "Column Ecc Mz Max [k-ft],Column Ecc Mz Min [k-ft],Column Ecc My Max [k-ft],Column Ecc My Min [k-ft]," +
                             "Integrity Force [k],Has Splice?\n";

            File.WriteAllText(file1, "");
            File.AppendAllText(file1, header1);

            FancyWriteLine("Writing internal forces table for lifts...", TextColor.Title);

            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
            {
                // Find the Integrity Force load combination
                var integrityForceCase = loadingCases.FirstOrDefault(lc =>
                    lc.Name.Equals("Integrity Force", StringComparison.CurrentCultureIgnoreCase) ||
                    lc.Name.Contains("Integrity", StringComparison.CurrentCultureIgnoreCase));

                foreach (var columnSpans in steelColumnSpans)
                {
                    var member = columnSpans.ParentMember;
                    string memberName = member.Name;
                    bool hasSplice = columnSpans.HasSplice;
                    string hasSpliceText = hasSplice ? "Yes" : "No";
                    var lifts = columnSpans.CreateLifts();

                    // Calculate integrity forces for this column if needed
                    var integrityForces = new Dictionary<string, double>();
                    if (integrityForceCase != null && hasSplice)
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
                        double rotationDeg = firstSpan.RotationAngle.Value * (180.0 / Math.PI);

                        // Get node fixity information
                        string startNodeFixity = GetNodeFixityDescription(firstSpan.StartReleases.Value);
                        string endNodeFixity = GetNodeFixityDescription(lift.Spans.Last().EndReleases.Value);

                        string startNodeName = $"{startNodeIdx}";
                        string endNodeName = $"{endNodeIdx}";

                        foreach (var loadingCase in loadingCases)
                        {
                            // Check UDA filter for this lift (check first span)
                            var udas = await firstSpan.GetUserDefinedAttributesAsync();
                            if (!string.IsNullOrEmpty(filterValue))
                            {
                                bool match = udas.Any(c =>
                                    (c as IUserDefinedTextAttribute)?.Text.Equals(filterValue, StringComparison.CurrentCultureIgnoreCase) == true &&
                                    c?.AttributeDefinitionName.Equals(filterField, StringComparison.CurrentCultureIgnoreCase) == true);
                                if (!match) continue;
                            }

                            IMemberLoading memberLoading = await member.GetLoadingAsync(loadingCase.Id, RequestedAnalysisType, LoadingResultType.Base);

                            // Get min/max forces for the entire lift
                            var (axialMax, axialMin) = await GetMinMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, lift, reduced);
                            var majorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, lift, reduced);
                            var majorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Major, lift, reduced);
                            var minorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, lift, reduced);
                            var minorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, lift, reduced);

                            // Eccentric moments
                            var (eccMzMax, eccMzMin) = await GetMinMaxEccentricMomentInLift(memberLoading, LoadingDirection.Major, lift, reduced);
                            var (eccMyMax, eccMyMin) = await GetMinMaxEccentricMomentInLift(memberLoading, LoadingDirection.Minor, lift, reduced);

                            // Get integrity force for this lift (only for Integrity Force load case)
                            double integrityForce = 0.0;
                            if (loadingCase == integrityForceCase && integrityForces.ContainsKey(lift.Name))
                            {
                                integrityForce = integrityForces[lift.Name];
                            }

                            // Write CSV line for this lift and loading case
                            string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(memberName)},{EscapeCsvValue(lift.Name)},{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)}," +
                                          $"{EscapeCsvValue(startNodeName)},{EscapeCsvValue(startNodeFixity)},{startX:F3},{startY:F3},{startZ:F3}," +
                                          $"{EscapeCsvValue(endNodeName)},{EscapeCsvValue(endNodeFixity)},{endX:F3},{endY:F3},{endZ:F3}," +
                                          $"{lengthFt:F3},{rotationDeg:F3},{EscapeCsvValue(loadingCase.Name)}," +
                                          $"{axialMax},{axialMin},{majorMomentMax},{majorShearMax},{minorMomentMax},{minorShearMax}," +
                                          $"{eccMzMax},{eccMzMin},{eccMyMax},{eccMyMin},{integrityForce},{EscapeCsvValue(hasSpliceText)}";

                            sw1.WriteLine(line);
                        }
                    }
                }
            }

            FancyWriteLine("Saved to: ", file1, "", TextColor.Path);
            double sizeKB = Math.Round(new FileInfo(file1).Length / 1024.0, 2);
            Console.WriteLine($"File size: {sizeKB} KB");
            double timeCSV = Math.Round(stopwatch.Elapsed.TotalSeconds - endWatch, 3);
            Console.WriteLine($"Steel Column lifts table written in {timeCSV} seconds.\n");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;
            Check();
        }

        /// <summary>
        /// Gets the maximum force in a lift for a given direction and value type.
        /// </summary>
        private static async Task<double> GetMaxForceInLift(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            double maxValue = 0.0;
            double valCon = ConversionFactor(valueType);

            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;
                for (int i = 0; i <= samplePoints; i++)
                {
                    double position = (i * spanLength) / samplePoints;
                    var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);
                    IEnumerable<ILoadingValue> values = await loading.GetValueAsync(option, span.Index, position);

                    if (values.Any())
                    {
                        double value = Math.Abs(values.OrderByDescending(lv => Math.Abs(lv.Value)).First().Value * valCon);
                        if (value > maxValue)
                            maxValue = value;
                    }
                }
            }

            return maxValue;
        }

        /// <summary>
        /// Gets the min and max force in a lift for a given direction and value type.
        /// </summary>
        private static async Task<(double max, double min)> GetMinMaxForceInLift(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            double maxValue = 0.0;
            double minValue = double.MaxValue;
            double valCon = ConversionFactor(valueType);

            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;
                for (int i = 0; i <= samplePoints; i++)
                {
                    double position = (i * spanLength) / samplePoints;
                    var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);
                    IEnumerable<ILoadingValue> values = await loading.GetValueAsync(option, span.Index, position);

                    if (values.Any())
                    {
                        double value = values.OrderByDescending(lv => Math.Abs(lv.Value)).First().Value * valCon;
                        if (value > 0 && value > maxValue)
                            maxValue = value;
                        if (value < minValue)
                            minValue = value;
                    }
                }
            }

            if (minValue == double.MaxValue) minValue = 0.0;
            return (maxValue, minValue);
        }

        /// <summary>
        /// Returns a string describing the node fixity based on releases.
        /// </summary>
        private static string GetNodeFixityDescription(ISpanReleases releases)
        {
            if (releases == null) return "Fixed";
            var fixityParts = new List<string>();
            try
            {
                var hasTransX = HasTranslationalRelease(releases, "X");
                var hasTransY = HasTranslationalRelease(releases, "Y");
                var hasTransZ = HasTranslationalRelease(releases, "Z");
                var hasRotX = HasRotationalRelease(releases, "X");
                var hasRotY = HasRotationalRelease(releases, "Y");
                var hasRotZ = HasRotationalRelease(releases, "Z");

                if (!hasTransX) fixityParts.Add("Tx");
                if (!hasTransY) fixityParts.Add("Ty");
                if (!hasTransZ) fixityParts.Add("Tz");
                if (!hasRotX) fixityParts.Add("Rx");
                if (!hasRotY) fixityParts.Add("Ry");
                if (!hasRotZ) fixityParts.Add("Rz");

                if (fixityParts.Count == 6)
                    return "Fixed";
                else if (fixityParts.Count == 0)
                    return "Pinned";
                else
                    return string.Join("-", fixityParts);
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// Checks if a translational release exists in a given direction.
        /// </summary>
        private static bool HasTranslationalRelease(ISpanReleases releases, string direction)
        {
            try
            {
                var type = releases.GetType();
                var property = type.GetProperty($"Translation{direction}") ??
                              type.GetProperty($"Translation{direction}Released") ??
                              type.GetProperty($"IsTranslation{direction}Released");

                if (property != null)
                {
                    var value = property.GetValue(releases);
                    if (value is bool boolValue)
                        return boolValue;
                    else if (value != null && value.GetType().GetProperty("Value") != null)
                        return (bool)value.GetType().GetProperty("Value").GetValue(value);
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Checks if a rotational release exists in a given direction.
        /// </summary>
        private static bool HasRotationalRelease(ISpanReleases releases, string direction)
        {
            try
            {
                var type = releases.GetType();
                var property = type.GetProperty($"Rotation{direction}") ??
                              type.GetProperty($"Rotation{direction}Released") ??
                              type.GetProperty($"IsRotation{direction}Released");

                if (property != null)
                {
                    var value = property.GetValue(releases);
                    if (value is bool boolValue)
                        return boolValue;
                    else if (value != null && value.GetType().GetProperty("Value") != null)
                        return (bool)value.GetType().GetProperty("Value").GetValue(value);
                }
            }
            catch { }
            return false;
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