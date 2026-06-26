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
using static TeklaResultsInterrogator.Utils.ConsoleUtils;

namespace TeklaResultsInterrogator.Commands
{
    /// <summary>
    /// Generates envelope eccentricity moments (max/min) for steel column lifts.
    /// </summary>
    public class SteelColumnEccentricityMoments : SolverInterrogator
    {
        /// <inheritdoc/>
        public override bool ShowInMenu() => true;

        /// <summary>Initializes a new instance of the <see cref="SteelColumnEccentricityMoments"/> class.</summary>
        public SteelColumnEccentricityMoments()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        /// <summary>
        /// Executes the command to retrieve and export eccentricity moments.
        /// </summary>
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
            var reduced = AskReduced();

            // Member Data
            FancyWriteLine("\nMember summary:", TextColor.Title);
            Console.WriteLine("Unpacking member data...");

            string? filterField = AskUser("What UDA field to filter on?");
            string? filterValue = AskUser("What UDA value to filter on?");

            stopwatch.Start();
            var steelColumns = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction))).ToList();

            Console.WriteLine($"{AllMembers!.Count} structural members found in model.");
            Console.WriteLine($"{steelColumns.Count} steel columns found.");

            // Organize Levels
            var rawLevels = await Model!.GetLevelsAsync();
            var levels = new List<IHorizontalConstructionPlane>();
            foreach (var item in rawLevels)
            {
                if (item is IHorizontalConstructionPlane hcp) levels.Add(hcp);
            }

            double timeUnpack = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Loading and member data unpacked in {timeUnpack} seconds.\n");

            FancyWriteLine("Organizing Column Data...", TextColor.Title);

            // Phase 1: Organize Spans and Collect Indices
            var columnData = new System.Collections.Concurrent.ConcurrentBag<(IMember Member, ColumnSpansSteel Spans)>();
            var allPointIndices = new System.Collections.Concurrent.ConcurrentBag<int>();

            var preTasks = new List<Task>();
            object consoleLock = new();

            foreach (var col in steelColumns)
            {
                preTasks.Add(Task.Run(async () =>
                {
                    var colSpans = new ColumnSpansSteel(col);
                    await colSpans.OrganizeSpansAsync();
                    columnData.Add((col, colSpans));

                    lock (consoleLock)
                    {
                        var lifts = colSpans.CreateLifts();
                        Console.WriteLine($"Column {col.Name}: {colSpans.Spans.Count} spans, {lifts.Count} lifts");
                    }

                    // Collect indices from lifts
                    foreach (var lift in colSpans.CreateLifts())
                    {
                        if (lift.StartNode?.ConstructionPointIndex != null)
                            allPointIndices.Add(lift.StartNode.ConstructionPointIndex.Value);
                        if (lift.EndNode?.ConstructionPointIndex != null)
                            allPointIndices.Add(lift.EndNode.ConstructionPointIndex.Value);
                    }
                }));
            }
            await Task.WhenAll(preTasks);

            // Phase 2: Batch Fetch Construction Points
            var uniqueIndices = allPointIndices.Distinct().ToList();
            Console.WriteLine($"\nFetching coordinates for {uniqueIndices.Count} unique points...");
            var pointsList = await Model.GetConstructionPointsAsync(uniqueIndices);
            var pointsDict = pointsList.ToDictionary(p => p.Index, p => p);

            FancyWriteLine("Querying Steel Column Eccentricity Moments (Parallel)...", TextColor.Title);

            string file1 = SaveDirectory + @"SteelColumnEccentricityMoments_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material," +
                             "Start Node,X_StartNode,Y_StartNode,Z_StartNode," +
                             "End Node,X_EndNode,Y_EndNode,Z_EndNode," +
                             "Lift Length [ft],Loading Name," +
                             "Column Ecc Mz Max [k-ft],Column Ecc Mz Min [k-ft],Column Ecc My Max [k-ft],Column Ecc My Min [k-ft]\n";

            File.WriteAllText(file1, header1);

            // Phase 3: Process Logic (Parallel)
            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism };
            var results = new System.Collections.Concurrent.ConcurrentBag<List<string>>();

            // Convert and sort for smooth progress
            var columnDataList = columnData.OrderBy(x => x.Member.Name).ToList();

            ApiMetrics.Reset();

            using var progress = new ProgressBar(columnDataList.Count);
            await Parallel.ForEachAsync(columnDataList, parallelOptions, async (item, token) =>
            {
                var (member, spans) = item;
                var colLines = await ProcessColumnAsync(member, spans, loadingCases, reduced, filterField, filterValue, levels, pointsDict);
                results.Add(colLines);
                progress.Increment();
            });

            using (StreamWriter sw1 = new(file1, true, Encoding.UTF8, bufferSize))
            {
                foreach (var result in results)
                {
                    foreach (var line in result)
                    {
                        sw1.WriteLine(line);
                    }
                }
            }

            FancyWriteLine("Saved to: ", file1, "", TextColor.Path);
            double sizeKB = Math.Round(new FileInfo(file1).Length / 1024.0, 2);
            Console.WriteLine($"File size: {sizeKB} KB");

            // Report Metrics
            Console.WriteLine("\n--- API Diagnostics ---");
            Console.WriteLine($"GetLoadingAsync:     {ApiMetrics.LoadingCalls} calls, Avg: {(ApiMetrics.LoadingCalls > 0 ? (double)ApiMetrics.LoadingDuration / ApiMetrics.LoadingCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"GetValueAsync:       {ApiMetrics.ValueCalls} calls, Avg: {(ApiMetrics.ValueCalls > 0 ? (double)ApiMetrics.ValueDuration / ApiMetrics.ValueCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"Peak Concurrency:    {ApiMetrics.MaxConcurrency}");
            if (ApiMetrics.SemaphoreWaitCalls > 0)
            {
                Console.WriteLine($"Semaphore Waits:     {ApiMetrics.SemaphoreWaitCalls} calls, Avg: {(double)ApiMetrics.SemaphoreWaitDuration / ApiMetrics.SemaphoreWaitCalls / 10000.0:F3} ms, Max: {ApiMetrics.SemaphoreWaitMax / 10000.0:F3} ms");
            }
            Console.WriteLine("-----------------------\n");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;
            Check();
        }

        private async Task<List<string>> ProcessColumnAsync(
            IMember member,
            ColumnSpansSteel colSpans,
            List<ILoadingCase> loadingCases,
            bool reduced,
            string? filterField,
            string? filterValue,
            List<IHorizontalConstructionPlane> levels,
            Dictionary<int, IConstructionPoint> pointsDict)
        {
            var lines = new List<string>();
            var lifts = colSpans.CreateLifts();

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

                Guid id = firstSpan.Id;
                string partMark = lifts.Count == 1 ? member.Name : firstSpan.Name;

                // Geometry and Levels
                int startNodeIdx = lift.StartNode!.ConstructionPointIndex.Value;
                IConstructionPoint? startPoint = pointsDict.ContainsKey(startNodeIdx) ? pointsDict[startNodeIdx] : null;
                double startX = 0, startY = 0, startZ = 0;
                string startLevelName = "Unknown";

                if (startPoint != null)
                {
                    startX = MmToFt(startPoint.Coordinates.Value.X);
                    startY = MmToFt(startPoint.Coordinates.Value.Y);
                    startZ = MmToFt(startPoint.Coordinates.Value.Z);
                    startLevelName = GetLevelName(startPoint.Coordinates.Value.Z, startPoint, levels);
                }

                int endNodeIdx = lift.EndNode!.ConstructionPointIndex.Value;
                IConstructionPoint? endPoint = pointsDict.ContainsKey(endNodeIdx) ? pointsDict[endNodeIdx] : null;
                double endX = 0, endY = 0, endZ = 0;
                string endLevelName = "Unknown";

                if (endPoint != null)
                {
                    endX = MmToFt(endPoint.Coordinates.Value.X);
                    endY = MmToFt(endPoint.Coordinates.Value.Y);
                    endZ = MmToFt(endPoint.Coordinates.Value.Z);
                    endLevelName = GetLevelName(endPoint.Coordinates.Value.Z, endPoint, levels);
                }

                string sectionName = "Unknown";
                string materialName = "Unknown";
                if (firstSpan.ElementSection.Value is IMemberSection elementSection &&
                    elementSection.PhysicalSection.Value is ISection physicalSection)
                {
                    sectionName = physicalSection.LongName;
                }
                if (firstSpan.Material?.Value != null)
                {
                    materialName = firstSpan.Material.Value.Name;
                }
                double lengthFt = MmToFt(lift.Length);

                foreach (var loadingCase in loadingCases)
                {
                    IMemberLoading memberLoading;
                    var swWait = Stopwatch.StartNew();
                    await ApiLimiter.WaitAsync();
                    swWait.Stop();
                    ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                    var sw = Stopwatch.StartNew();
                    try
                    {
                        memberLoading = await member.GetLoadingAsync(loadingCase.Id, RequestedAnalysisType, LoadingResultType.Base);
                    }
                    finally
                    {
                        sw.Stop();
                        ApiLimiter.Release();
                    }
                    ApiMetrics.RecordLoading(sw.ElapsedTicks);

                    // Sequential to avoid API congestion
                    var (eccMzMax, eccMzMin) = await GetMinMaxEccentricMomentInLift(memberLoading, LoadingDirection.Major, lift, reduced);
                    var (eccMyMax, eccMyMin) = await GetMinMaxEccentricMomentInLift(memberLoading, LoadingDirection.Minor, lift, reduced);

                    string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(member.Name)},{EscapeCsvValue(lift.Name)}," +
                                $"{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)}," +
                                $"{startNodeIdx},{startX:F3},{startY:F3},{startZ:F3}," +
                                $"{endNodeIdx},{endX:F3},{endY:F3},{endZ:F3}," +
                                $"{lengthFt:F3},{EscapeCsvValue(loadingCase.Name)}," +
                                $"{eccMzMax},{eccMzMin},{eccMyMax},{eccMyMin}";
                    lines.Add(line);
                }
            }
            return lines;
        }

        private static async Task<(double max, double min)> GetMinMaxEccentricMomentInLift(
            IMemberLoading loading,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            double Nmm_to_kft = ConversionFactor(LoadingValueType.Moment);
            var allValues = new List<double>();

            // Sequential processing to avoid API congestion
            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;

                for (int i = 0; i <= samplePoints; i++)
                {
                    double pos = (i * spanLength) / samplePoints;
                    try
                    {
                        var option = LoadingValueOptions.StaticValue(LoadingValueType.EccentricityMoment, direction, reduced);
                        IEnumerable<ILoadingValue> vals;
                        var swWait = Stopwatch.StartNew();
                        await ApiLimiter.WaitAsync();
                        swWait.Stop();
                        ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                        var sw = Stopwatch.StartNew();
                        ApiMetrics.IncrementActiveValueCalls();
                        try
                        {
                            vals = await loading.GetValueAsync(option, span.Index, pos);
                        }
                        finally
                        {
                            ApiMetrics.DecrementActiveValueCalls();
                            sw.Stop();
                            ApiLimiter.Release();
                        }
                        ApiMetrics.RecordValue(sw.ElapsedTicks);

                        if (vals.Any())
                        {
                            allValues.Add(vals.MaxBy(v => Math.Abs(v.Value))!.Value * Nmm_to_kft);
                        }
                    }
                    catch { }
                }
            }

            double globalMax = allValues.DefaultIfEmpty(0.0).Max();
            double globalMin = allValues.DefaultIfEmpty(0.0).Min();

            return (globalMax, globalMin);
        }


    }
}
