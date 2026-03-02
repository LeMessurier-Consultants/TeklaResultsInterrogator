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
    /// Generates envelope forces (max/min axial, moments, shears) for steel column lifts
    /// </summary>
    public class SteelColumnEnvelopes : SolverInterrogator
    {
        /// <inheritdoc/>
        public override bool ShowInMenu() => true;

        /// <summary>Initializes a new instance of the <see cref="SteelColumnEnvelopes"/> class.</summary>
        public SteelColumnEnvelopes()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        /// <summary>
        /// Executes the command to retrieve and envelope steel column forces.
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
            bool reduced = AskReduced();

            // Unpacking member data
            FancyWriteLine("\nMember summary:", TextColor.Title);
            Console.WriteLine("Unpacking member data...");

            var steelColumns = AskAndFilterMembers(true, true);

            string? filterField = AskUser("What UDA field to filter on?");
            string? filterValue = AskUser("What UDA value to filter on?");

            stopwatch.Start();

            Console.WriteLine($"{AllMembers!.Count} structural members found in model.");
            Console.WriteLine($"{steelColumns.Count} steel columns found.");

            var levels = (await Model!.GetLevelsAsync()).ToList();
            double timeUnpack = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Loading and member data unpacked in {timeUnpack} seconds.\n");

            // Organize Column Lifts & Pre-Fetch Geometry
            FancyWriteLine("Organizing Column Lifts & Geometry...", TextColor.Title);

            // Phase 1: Organize Spans and Collect Indices
            var columnData = new System.Collections.Concurrent.ConcurrentBag<(IMember Member, ColumnSpansSteel Spans, List<ColumnLift> Lifts)>();
            var allPointIndices = new System.Collections.Concurrent.ConcurrentBag<int>();

            var preTasks = new List<Task>();
            object consoleLock = new();

            foreach (var col in steelColumns)
            {
                preTasks.Add(Task.Run(async () =>
                {
                    var colSpans = new ColumnSpansSteel(col);
                    await colSpans.OrganizeSpansAsync();
                    var lifts = colSpans.CreateLifts();

                    columnData.Add((col, colSpans, lifts));

                    // Log progress
                    lock (consoleLock)
                    {
                        string spliceText = colSpans.HasSplice ? "has splice" : "no splice";
                        Console.WriteLine($"Column {col.Name}: {colSpans.Spans.Count} spans, {spliceText}");
                    }

                    // Collect indices (Start AND End nodes)
                    foreach (var lift in lifts)
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

            FancyWriteLine("\nQuerying Steel Column Forces (Parallel)...", TextColor.Title);

            // Setting up file
            string file1 = SaveDirectory + @"SteelColumnEnvelopes_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material,Section Area [in2]," +
                 "Start Node,Start Node Fixity,X_StartNode,Y_StartNode,Z_StartNode," +
                 "End Node,End Node Fixity,X_EndNode,Y_EndNode,Z_EndNode," +
                 "Lift Length [ft],Span Rotation [deg],Loading Name," +
                 "Column Axial Max [k],Column Axial Min [k],Column Major Moment Max [k-ft]," +
                 "Column Major Shear Max [k],Column Minor Moment Max [k-ft],Column Minor Shear Max [k]\n";
            File.WriteAllText(file1, header1);

            var results = new System.Collections.Concurrent.ConcurrentBag<List<string>>();

            // Convert to List for smoother partitioning (Sort by Name for consistency)
            var columnDataList = columnData.OrderBy(x => x.Member.Name).ToList();

            // -----------------------------------------------------------------------------------------
            // PRODUCTION: Bounded Parallelism (Memory-Safe)
            // Parallel.ForEachAsync limits concurrent columns to prevent OOM on large models.
            // Inner loops use sequential awaits to prevent API congestion; ApiLimiter caps active calls.
            // -----------------------------------------------------------------------------------------
            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = SolverInterrogator.MaxDegreeOfParallelism };

            ApiMetrics.Reset();

            using var progress = new ProgressBar(columnDataList.Count);
            await Parallel.ForEachAsync(columnDataList, parallelOptions, async (data, token) =>
            {
                var (col, colSpans, lifts) = data;
                var colLines = await ProcessColumnAsync(col, lifts, loadingCases, RequestedAnalysisType, reduced, filterField, filterValue, levels, pointsDict);
                results.Add(colLines);
                progress.Increment();
            });
            double endWatch = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);

            FancyWriteLine("Writing steel Column forces table...", TextColor.Title);

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
            double timeCSV = Math.Round(stopwatch.Elapsed.TotalSeconds - endWatch, 3);
            Console.WriteLine($"Steel Column table written in {timeCSV} seconds.\n");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;
            Check();

            // Report Metrics
            Console.WriteLine("\n--- API Diagnostics ---");
            Console.WriteLine($"GetLoadingAsync: {ApiMetrics.LoadingCalls} calls, Avg: {(ApiMetrics.LoadingCalls > 0 ? (double)ApiMetrics.LoadingDuration / ApiMetrics.LoadingCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"GetValueAsync:   {ApiMetrics.ValueCalls} calls, Avg: {(ApiMetrics.ValueCalls > 0 ? (double)ApiMetrics.ValueDuration / ApiMetrics.ValueCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"Peak Concurrency (Points): {ApiMetrics.MaxConcurrency}");
            if (ApiMetrics.SemaphoreWaitCalls > 0)
            {
                Console.WriteLine($"Semaphore Waits: {ApiMetrics.SemaphoreWaitCalls} calls, Avg: {(double)ApiMetrics.SemaphoreWaitDuration / ApiMetrics.SemaphoreWaitCalls / 10000.0:F3} ms, Max: {ApiMetrics.SemaphoreWaitMax / 10000.0:F3} ms");
            }
            Console.WriteLine("-----------------------\n");
        }

        private static async Task<List<string>> ProcessColumnAsync(
            IMember member,
            List<ColumnLift> lifts,
            List<ILoadingCase> loadingCases,
            AnalysisType analysisType,
            bool reduced,
            string? filterField,
            string? filterValue,
            List<IHorizontalConstructionPlane> levels,
            Dictionary<int, IConstructionPoint> pointsDict)
        {
            var lines = new List<string>();

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

                // Get node coordinates and levels using cached dictionary
                // Start Node
                int startNodeIdx = lift.StartNode?.ConstructionPointIndex?.Value ?? -1;
                IConstructionPoint? startPoint = (startNodeIdx != -1 && pointsDict.ContainsKey(startNodeIdx)) ? pointsDict[startNodeIdx] : null;

                double startX = 0, startY = 0, startZ = 0;
                string startLevelName = "Unknown";
                string startNodeName = $"{startNodeIdx}";

                if (startPoint != null)
                {
                    startX = MmToFt(startPoint.Coordinates.Value.X);
                    startY = MmToFt(startPoint.Coordinates.Value.Y);
                    startZ = MmToFt(startPoint.Coordinates.Value.Z);
                    startLevelName = GetLevelName(startPoint.Coordinates.Value.Z, startPoint, levels);
                }

                // End Node
                int endNodeIdx = lift.EndNode?.ConstructionPointIndex?.Value ?? -1;
                IConstructionPoint? endPoint = (endNodeIdx != -1 && pointsDict.ContainsKey(endNodeIdx)) ? pointsDict[endNodeIdx] : null;

                double endX = 0, endY = 0, endZ = 0;
                string endLevelName = "Unknown";
                string endNodeName = $"{endNodeIdx}";

                if (endPoint != null)
                {
                    endX = MmToFt(endPoint.Coordinates.Value.X);
                    endY = MmToFt(endPoint.Coordinates.Value.Y);
                    endZ = MmToFt(endPoint.Coordinates.Value.Z);
                    endLevelName = GetLevelName(endPoint.Coordinates.Value.Z, endPoint, levels);
                }

                // Get section and material
                string sectionName = "Unknown";
                string materialName = "Unknown";
                double sectionArea = 0.0;

                if (firstSpan.ElementSection.Value != null)
                {
                    var elementSection = (IMemberSection)firstSpan.ElementSection.Value;
                    var physicalSection = (ISection)elementSection.PhysicalSection.Value;
                    sectionName = physicalSection.LongName ?? "Unknown";

                    // Get section area and convert from mm² to in²
                    sectionArea = MmSqToInSq(physicalSection.CrossSectionalArea);
                }

                if (firstSpan.Material?.Value != null)
                {
                    materialName = firstSpan.Material.Value.Name ?? "Unknown";
                }

                double lengthFt = MmToFt(lift.Length);
                double rotationDeg = RadToDeg(firstSpan.RotationAngle?.Value ?? 0.0);

                // Fixity
                string startNodeFixity = GetNodeFixityDescription(firstSpan.StartReleases.Value);
                string endNodeFixity = GetNodeFixityDescription(lift.Spans.Last().EndReleases.Value);

                foreach (var loadingCase in loadingCases)
                {
                    // Get envelope forces 
                    var swWait = Stopwatch.StartNew();
                    await ApiLimiter.WaitAsync();
                    swWait.Stop();
                    ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                    IMemberLoading memberLoading;
                    var swL = Stopwatch.StartNew();
                    try
                    {
                        memberLoading = await member.GetLoadingAsync(loadingCase.Id, analysisType, LoadingResultType.Base);
                    }
                    finally
                    {
                        swL.Stop();
                        ApiMetrics.RecordLoading(swL.ElapsedTicks);
                        ApiLimiter.Release();
                    }

                    // Get envelope forces (sequential to avoid API congestion)
                    var (axialMax, axialMin) = await GetMinMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, lift, reduced);

                    double majorMomentMax = 0;
                    double majorShearMax = 0;
                    double minorMomentMax = 0;
                    double minorShearMax = 0;

                    // For envelopes, query moments/shears from child combinations
                    // (Envelopes don't store moments/shears directly, only forces)
                    if (loadingCase is IEnvelope envelope && envelope.CombinationIds != null)
                    {
                        try
                        {
                            majorMomentMax = await GetMaxMomentFromEnvelope(member, envelope, LoadingDirection.Major, lift, reduced, analysisType);
                            majorShearMax = await GetMaxShearFromEnvelope(member, envelope, LoadingDirection.Major, lift, reduced, analysisType);
                            minorMomentMax = await GetMaxMomentFromEnvelope(member, envelope, LoadingDirection.Minor, lift, reduced, analysisType);
                            minorShearMax = await GetMaxShearFromEnvelope(member, envelope, LoadingDirection.Minor, lift, reduced, analysisType);
                        }
                        catch
                        {
                            // If envelope querying fails, moments stay at 0
                        }
                    }
                    else if (!(loadingCase is IEnvelope))
                    {
                        // For regular cases/combos, query directly
                        majorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, lift, reduced);
                        majorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Major, lift, reduced);
                        minorMomentMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, lift, reduced);
                        minorShearMax = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, lift, reduced);
                    }

                    string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(member.Name)},{EscapeCsvValue(lift.Name)}," +
                          $"{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)},{sectionArea:F3}," +
                          $"{EscapeCsvValue(startNodeName)},{EscapeCsvValue(startNodeFixity)},{startX:F3},{startY:F3},{startZ:F3}," +
                          $"{EscapeCsvValue(endNodeName)},{EscapeCsvValue(endNodeFixity)},{endX:F3},{endY:F3},{endZ:F3}," +
                          $"{lengthFt:F3},{rotationDeg:F3},{EscapeCsvValue(loadingCase.Name)}," +
                          $"{axialMax},{axialMin},{majorMomentMax},{majorShearMax},{minorMomentMax},{minorShearMax}";

                    lines.Add(line);
                }
            }
            return lines;
        }

        private static async Task<double> GetMaxMomentFromEnvelope(IMember member, IEnvelope envelope, LoadingDirection direction, ColumnLift lift, bool reduced, AnalysisType analysisType)
        {
            double maxMoment = 0;
            var momentLock = new object();

            try
            {
                // Create list of tasks for parallel execution
                var tasks = new List<Task>();

                for (int i = 0; envelope.CombinationIds != null && i < envelope.CombinationIds.Count; i++)
                {
                    int index = i; // Capture for closure
                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            var combId = envelope.CombinationIds[index].Value;
                            var memberLoading = await member.GetLoadingAsync(combId, analysisType, LoadingResultType.Base);
                            var moment = await GetMaxForceInLift(memberLoading, LoadingValueType.Moment, direction, lift, reduced);

                            lock (momentLock)
                            {
                                if (moment > maxMoment)
                                    maxMoment = moment;
                            }
                        }
                        catch { }
                    }));
                }

                await Task.WhenAll(tasks);
            }
            catch { }

            return maxMoment;
        }

        private static async Task<double> GetMaxShearFromEnvelope(IMember member, IEnvelope envelope, LoadingDirection direction, ColumnLift lift, bool reduced, AnalysisType analysisType)
        {
            double maxShear = 0;
            var shearLock = new object();

            try
            {
                // Create list of tasks for parallel execution
                var tasks = new List<Task>();

                for (int i = 0; envelope.CombinationIds != null && i < envelope.CombinationIds.Count; i++)
                {
                    int index = i; // Capture for closure
                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            var combId = envelope.CombinationIds[index].Value;
                            var memberLoading = await member.GetLoadingAsync(combId, analysisType, LoadingResultType.Base);
                            var shear = await GetMaxForceInLift(memberLoading, LoadingValueType.Force, direction, lift, reduced);

                            lock (shearLock)
                            {
                                if (shear > maxShear)
                                    maxShear = shear;
                            }
                        }
                        catch { }
                    }));
                }

                await Task.WhenAll(tasks);
            }
            catch { }

            return maxShear;
        }

        private static async Task<double> GetMaxForceInLift(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            var allValues = new List<double>();

            // Sequential processing to avoid API congestion
            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;
                var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);

                for (int i = 0; i <= samplePoints; i++)
                {
                    double pos = (i * spanLength) / samplePoints;
                    try
                    {
                        var swWait = Stopwatch.StartNew();
                        await ApiLimiter.WaitAsync();
                        swWait.Stop();
                        ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                        try
                        {
                            ApiMetrics.IncrementActiveValueCalls();
                            var swV = Stopwatch.StartNew();

                            var vals = await loading.GetValueAsync(option, span.Index, pos);

                            swV.Stop();
                            ApiMetrics.RecordValue(swV.ElapsedTicks);
                            ApiMetrics.DecrementActiveValueCalls();

                            if (vals.Any())
                            {
                                allValues.Add(Math.Abs(vals.MaxBy(v => Math.Abs(v.Value))?.Value ?? 0.0));
                            }
                        }
                        finally
                        {
                            ApiLimiter.Release();
                        }
                    }
                    catch { }
                }
            }

            return allValues.DefaultIfEmpty(0.0).Max() * ConversionFactor(valueType);
        }

        private static async Task<(double max, double min)> GetMinMaxForceInLift(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            ColumnLift lift,
            bool reduced)
        {
            var allValues = new List<double>();

            // Sequential processing to avoid API congestion
            foreach (var span in lift.Spans)
            {
                int samplePoints = 10;
                double spanLength = span.Length.Value;
                var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);

                for (int i = 0; i <= samplePoints; i++)
                {
                    double pos = (i * spanLength) / samplePoints;
                    try
                    {
                        var swWait = Stopwatch.StartNew();
                        await ApiLimiter.WaitAsync();
                        swWait.Stop();
                        ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                        try
                        {
                            ApiMetrics.IncrementActiveValueCalls();
                            var swV = Stopwatch.StartNew();

                            var vals = await loading.GetValueAsync(option, span.Index, pos);

                            swV.Stop();
                            ApiMetrics.RecordValue(swV.ElapsedTicks);
                            ApiMetrics.DecrementActiveValueCalls();

                            if (vals.Any())
                            {
                                // Add all values to capture both positive and negative extremes
                                foreach (var v in vals)
                                {
                                    allValues.Add(v.Value);
                                }
                            }
                        }
                        finally
                        {
                            ApiLimiter.Release();
                        }
                    }
                    catch { }
                }
            }

            double globalMax = allValues.Count > 0 ? allValues.Max() : 0.0;
            double globalMin = allValues.Count > 0 ? allValues.Min() : 0.0;

            double valCon = ConversionFactor(valueType);
            return (globalMax * valCon, globalMin * valCon);
        }

    }
}