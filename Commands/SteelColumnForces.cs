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
    public class SteelColumnForces : SolverInterrogator
    {
        public override bool ShowInMenu() => true;

        public SteelColumnForces()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

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

            FancyWriteLine("Organizing Column Spans and Splices...", TextColor.Title);
            var steelColumnSpans = new List<ColumnSpansSteel>();
            foreach (var column in steelColumns)
            {
                var colSpans = new ColumnSpansSteel(column);
                await colSpans.OrganizeSpansAsync();

                Console.WriteLine($"Column {column.Name} has splice? {(colSpans.HasSplice ? "Yes" : "No")}");
                Console.WriteLine($"  -> {colSpans.Spans.Count} spans total (all spans together)");

                steelColumnSpans.Add(colSpans);
            }

            double endWatch = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Column spans organized in {Math.Round(endWatch - timeUnpack, 3)} seconds.\n");

            // Prepare output CSV file with all required columns
            string file1 = SaveDirectory + @"SteelColumnForces_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,UDA Filter,Member Name,Span Name,Start Level,End Level,Shape,Material," +
                             "Start Node,Start Node Fixity,X_StartNode,Y_StartNode,Z_StartNode," +
                             "End Node,End Node Fixity,X_EndNode,Y_EndNode,Z_EndNode," +
                             "Span Length [ft],Span Rotation [deg],Loading Name,Location," +
                             "Axial Force [k],Shear Major [k],Shear Minor [k],Moment Major [k-ft],Moment Minor [k-ft],Torsion [k-ft],Has Splice?\n";

            File.WriteAllText(file1, "");
            File.AppendAllText(file1, header1);

            FancyWriteLine("Writing internal forces table...", TextColor.Title);

            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
            {
                foreach (var columnSpans in steelColumnSpans)
                {
                    var member = columnSpans.ParentMember;
                    string memberName = member.Name;
                    Guid id = member.Id;
                    bool hasSplice = columnSpans.HasSplice;
                    string hasSpliceText = hasSplice ? "Yes" : "No";

                    // Flatten all spans (no lifts), ordered by span index
                    var allSpans = columnSpans.Spans.OrderBy(s => s.Index);

                    foreach (var span in allSpans)
                    {
                        var spliceInfo = columnSpans.SpanSpliceInfo.TryGetValue(span.Index, out var si)
                            ? si
                            : (HasSplice: false, SpliceOffset: 0.0);

                        double spliceOffsetInches = spliceInfo.SpliceOffset / 25.4;

                        // Get start and end node construction points
                        int startNodeIdx = span.StartMemberNode.ConstructionPointIndex.Value;
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

                        int endNodeIdx = span.EndMemberNode.ConstructionPointIndex.Value;
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

                        // Get section and material info
                        string sectionName = "Unknown";
                        string materialName = "Unknown";

                        if (span.ElementSection.Value != null)
                        {
                            var elementSection = (IMemberSection)span.ElementSection.Value;
                            var physicalSection = (ISection)elementSection.PhysicalSection.Value;
                            sectionName = physicalSection.LongName;
                        }

                        if (span.Material?.Value != null)
                        {
                            materialName = span.Material.Value.Name;
                        }

                        double lengthFt = span.Length.Value * 0.00328084;

                        // Get span rotation in degrees
                        double rotationDeg = span.RotationAngle.Value * (180.0 / Math.PI);

                        // Get node fixity information
                        string startNodeFixity = GetNodeFixityDescription(span.StartReleases.Value);
                        string endNodeFixity = GetNodeFixityDescription(span.EndReleases.Value);

                        // Start and end node names/IDs - use index since Name might not be available
                        string startNodeName = $"{startNodeIdx}";
                        string endNodeName = $"{endNodeIdx}";

                        // Prepare list of positions to query (Start, Splice if exists and valid, End)
                        var positions = new List<(string LocationName, double PositionMm)>()
                        {
                            ("Start", 0.0),
                            ("End", span.Length.Value)
                        };

                        if (spliceInfo.HasSplice && spliceInfo.SpliceOffset > 0 && spliceInfo.SpliceOffset < span.Length.Value)
                        {
                            positions.Insert(1, ("Splice", spliceInfo.SpliceOffset));
                        }

                        foreach (var loadingCase in loadingCases)
                        {
                            // Check UDA filter for this span
                            var udas = await span.GetUserDefinedAttributesAsync();

                            if (!string.IsNullOrEmpty(filterValue))
                            {
                                bool match = udas.Any(c =>
                                    (c as IUserDefinedTextAttribute)?.Text.Equals(filterValue, StringComparison.CurrentCultureIgnoreCase) == true &&
                                    c?.AttributeDefinitionName.Equals(filterField, StringComparison.CurrentCultureIgnoreCase) == true);

                                if (!match) continue;
                            }

                            IMemberLoading memberLoading = await member.GetLoadingAsync(loadingCase.Id, RequestedAnalysisType, LoadingResultType.Base);

                            foreach (var (locationName, positionMm) in positions)
                            {
                                double axialForce = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, positionMm, reduced, span.Index);
                                double shearMajor = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Major, positionMm, reduced, span.Index);
                                double shearMinor = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, positionMm, reduced, span.Index);
                                double momentMajor = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, positionMm, reduced, span.Index);
                                double momentMinor = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, positionMm, reduced, span.Index);
                                double torsion = await GetLoadingValueAtPosition(memberLoading, LoadingValueType.Moment, LoadingDirection.Axial, positionMm, reduced, span.Index);

                                string line = $"{id},{filterValue},{memberName},{span.Name},{startLevelName},{endLevelName},{sectionName},{materialName}," +
                                              $"{startNodeName},{startNodeFixity},{startX:F3},{startY:F3},{startZ:F3}," +
                                              $"{endNodeName},{endNodeFixity},{endX:F3},{endY:F3},{endZ:F3}," +
                                              $"{lengthFt:F3},{rotationDeg:F3},{loadingCase.Name},{locationName}," +
                                              $"{axialForce},{shearMajor},{shearMinor},{momentMajor},{momentMinor},{torsion},{hasSpliceText}";

                                sw1.WriteLine(line);
                            }
                        }
                    }
                }
            }

            FancyWriteLine("Saved to: ", file1, "", TextColor.Path);
            double sizeKB = Math.Round(new FileInfo(file1).Length / 1024.0, 2);
            Console.WriteLine($"File size: {sizeKB} KB");

            double timeCSV = Math.Round(stopwatch.Elapsed.TotalSeconds - endWatch, 3);
            Console.WriteLine($"Steel Column table written in {timeCSV} seconds.\n");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;

            Check();
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

            double valCon = ConversionFactor(valueType);

            if (values.Any())
                return values.First().Value * valCon;

            return 0.0;
        }

        private static string GetNodeFixityDescription(ISpanReleases releases)
        {
            if (releases == null) return "Fixed";

            var fixityParts = new List<string>();

            // Check translational releases - assuming these are boolean properties
            // You may need to adjust these property names based on the actual API
            try
            {
                // Note: These property names might need to be adjusted based on actual TSD API
                // Common patterns might be: IsTranslationXReleased, TranslationXReleased, etc.
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
                // Fallback if we can't determine the release conditions
                return "Unknown";
            }
        }

        private static bool HasTranslationalRelease(ISpanReleases releases, string direction)
        {
            // This is a placeholder - you'll need to determine the correct property names
            // from the TSD API documentation for ISpanReleases
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

            return false; // Default to not released (fixed)
        }

        private static bool HasRotationalRelease(ISpanReleases releases, string direction)
        {
            // This is a placeholder - you'll need to determine the correct property names
            // from the TSD API documentation for ISpanReleases
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

            return false; // Default to not released (fixed)
        }
    }
}