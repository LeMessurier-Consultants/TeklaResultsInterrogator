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
    /// Generates envelope forces (max/min axial, moments, shears) for steel column lifts
    /// </summary>
    public class SteelColumnEnvelopes : SolverInterrogator
    {
        public override bool ShowInMenu() => true;

        public SteelColumnEnvelopes()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        #region Helper Classes

        /// <summary>
        /// Organizes steel column spans and detects splices
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

            /// <summary>
            /// Organizes spans and detects splices through multiple methods
            /// </summary>
            public async Task OrganizeSpansAsync()
            {
                var spans = (await ParentMember.GetSpanAsync())
                    .OrderBy(s => s.Index)
                    .ToList();

                HasSplice = false;
                SpanSpliceInfo.Clear();

                // First pass: Check existing splice data from API
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

                // Second pass: Check for name-based splice detection
                await DetectSplicesFromNamesAsync(spans);

                // Third pass: Check for span number discontinuity
                DetectSplicesFromSpanNumbers(spans);

                Spans = spans;
            }

            /// <summary>
            /// Detects splices based on span names and UDAs
            /// </summary>
            private async Task DetectSplicesFromNamesAsync(List<IMemberSpan> spans)
            {
                var spliceKeywords = new[] { "splice", "connection", "joint", "lift", "piece", "stack" };

                foreach (var span in spans)
                {
                    try
                    {
                        bool nameIndicatesSplice = false;
                        double nameBasedSpliceOffset = span.Length.Value;

                        // Check span name
                        string spanName = span.Name?.ToLowerInvariant() ?? "";
                        if (spliceKeywords.Any(keyword => spanName.Contains(keyword)))
                        {
                            nameIndicatesSplice = true;
                        }

                        // Check UDAs
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

                                    // Try to extract splice offset
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

            /// <summary>
            /// Detects splices based on span numbering gaps
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
            /// Creates column lifts based on splice locations
            /// </summary>
            public List<ColumnLift> CreateLifts()
            {
                var lifts = new List<ColumnLift>();

                if (!HasSplice)
                {
                    // No splice - entire column is one lift
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

                // Has splices - create lifts between splices
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

                // Add final lift
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
        /// Represents a column lift (section between splices)
        /// </summary>
        public class ColumnLift
        {
            public string Name { get; set; }
            public List<IMemberSpan> Spans { get; set; } = new List<IMemberSpan>();
            public IMemberNode StartNode { get; set; }
            public IMemberNode EndNode { get; set; }
            public double Length { get; set; } // Total length in mm
        }

        #endregion

        #region Force Calculation Methods

        /// <summary>
        /// Gets max absolute value of force in a lift
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
        /// Gets min and max values of force in a lift (preserving sign)
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

        #endregion

        #region Utility Methods

        /// <summary>
        /// Gets node fixity description string
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
        /// Escapes CSV values properly
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

        #endregion

        public override async Task ExecuteAsync()
        {
            await InitializeAsync();
            if (Flag) return;

            Stopwatch stopwatch = Stopwatch.StartNew();
            int bufferSize = 65536 * 2;

            // Loading Summary
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

            // Organize Column Lifts
            FancyWriteLine("Organizing Column Lifts...", TextColor.Title);
            var steelColumnSpans = new List<ColumnSpansSteel>();

            foreach (var column in steelColumns)
            {
                var colSpans = new ColumnSpansSteel(column);
                await colSpans.OrganizeSpansAsync();
                steelColumnSpans.Add(colSpans);

                Console.WriteLine($"Column {column.Name}: {colSpans.Spans.Count} spans, {(colSpans.HasSplice ? "has splice" : "no splice")}");
            }

            // Prepare output CSV file
            string file1 = SaveDirectory + @"SteelColumnEnvelopes_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material," +
                             "Start Node,Start Node Fixity,X_StartNode,Y_StartNode,Z_StartNode," +
                             "End Node,End Node Fixity,X_EndNode,Y_EndNode,Z_EndNode," +
                             "Lift Length [ft],Span Rotation [deg],Loading Name," +
                             "Column Axial Max [k],Column Axial Min [k],Column Major Moment Max [k-ft]," +
                             "Column Major Shear Max [k],Column Minor Moment Max [k-ft],Column Minor Shear Max [k]\n";

            File.WriteAllText(file1, header1);

            FancyWriteLine("Writing steel Column forces table...", TextColor.Title);

            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
            {
                foreach (var columnSpans in steelColumnSpans)
                {
                    var member = columnSpans.ParentMember;
                    string memberName = member.Name;
                    var lifts = columnSpans.CreateLifts();

                    foreach (var lift in lifts)
                    {
                        var firstSpan = lift.Spans.First();

                        // Check UDA filter
                        var udas = await firstSpan.GetUserDefinedAttributesAsync();
                        if (!string.IsNullOrEmpty(filterValue))
                        {
                            bool match = udas.Any(c =>
                                (c as IUserDefinedTextAttribute)?.Text.Equals(filterValue, StringComparison.CurrentCultureIgnoreCase) == true &&
                                c?.AttributeDefinitionName.Equals(filterField, StringComparison.CurrentCultureIgnoreCase) == true);
                            if (!match) continue;
                        }

                        // Get basic lift information
                        Guid id = firstSpan.Id;
                        string partMark = lifts.Count == 1 ? member.Name : firstSpan.Name;

                        // Get node coordinates and levels
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

                        // Get section and material
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

                        // Get node fixity
                        string startNodeFixity = GetNodeFixityDescription(firstSpan.StartReleases.Value);
                        string endNodeFixity = GetNodeFixityDescription(lift.Spans.Last().EndReleases.Value);

                        string startNodeName = $"{startNodeIdx}";
                        string endNodeName = $"{endNodeIdx}";

                        // Process each loading case
                        foreach (var loadingCase in loadingCases)
                        {
                            IMemberLoading memberLoading = await member.GetLoadingAsync(loadingCase.Id, RequestedAnalysisType, LoadingResultType.Base);

                            // Get envelope forces
                            var (axialMax, axialMin) = await GetMinMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, lift, reduced);
                            var majorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, lift, reduced);
                            var majorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Major, lift, reduced);
                            var minorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, lift, reduced);
                            var minorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, lift, reduced);

                            // Write row
                            string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(memberName)},{EscapeCsvValue(lift.Name)},{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)}," +
                                          $"{EscapeCsvValue(startNodeName)},{EscapeCsvValue(startNodeFixity)},{startX:F3},{startY:F3},{startZ:F3}," +
                                          $"{EscapeCsvValue(endNodeName)},{EscapeCsvValue(endNodeFixity)},{endX:F3},{endY:F3},{endZ:F3}," +
                                          $"{lengthFt:F3},{rotationDeg:F3},{EscapeCsvValue(loadingCase.Name)}," +
                                          $"{axialMax},{axialMin},{majorMomentMax},{majorShearMax},{minorMomentMax},{minorShearMax}";

                            sw1.WriteLine(line);
                        }
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
    }
}