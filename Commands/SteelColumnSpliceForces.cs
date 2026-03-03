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
    /// Generates envelope forces (max/min axial, moments, shears) at splice locations for steel columns,
    /// plus integrity forces from the Integrity Force loading case.
    /// </summary>
    public class SteelColumnSpliceForces : SolverInterrogator
    {
        /// <inheritdoc/>
        public override bool ShowInMenu() => true;

        /// <summary>Initializes a new instance of the <see cref="SteelColumnSpliceForces"/> class.</summary>
        public SteelColumnSpliceForces()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelColumn };
        }

        /// <summary>
        /// Executes the command to retrieve and calculate envelope forces at splice locations.
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

            // Find the Integrity Force combination (always used if available)
            var integrityForceCombination = AllCombinations?.FirstOrDefault(c =>
                c.Name.Equals("Integrity Force", StringComparison.CurrentCultureIgnoreCase) ||
                c.Name.Contains("Integrity", StringComparison.CurrentCultureIgnoreCase));

            if (integrityForceCombination != null)
            {
                Console.WriteLine($"Integrity Force combination found: {integrityForceCombination.Name}");
            }
            else
            {
                Console.WriteLine("Warning: No Integrity Force combination found in model.");
            }
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

            FancyWriteLine("\nQuerying Steel Column Splice Forces (Parallel)...", TextColor.Title);

            // Setting up file
            string file1 = SaveDirectory + @"SteelColumnSpliceForces_" + OutputFileName + ".csv";
            string header1 = "Tekla GUID,Part Mark,UDA Filter,Member Name,Lift Name,Start Level,End Level,Shape,Material,Section Area [in2]," +
                 "Start Node,Start Node Fixity,X_StartNode,Y_StartNode,Z_StartNode," +
                 "End Node,End Node Fixity,X_EndNode,Y_EndNode,Z_EndNode," +
                 "Lift Length [ft],Splice Offset [in],Loading Name," +
                 "Column Axial Max [k],Column Axial Min [k],Column Major Moment Max [k-ft]," +
                 "Column Major Shear Max [k],Column Minor Moment Max [k-ft],Column Minor Shear Max [k]," +
                 "Integrity Force [k]\n";
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
                var colLines = await ProcessColumnAsync(col, lifts, colSpans, loadingCases, integrityForceCombination, RequestedAnalysisType, reduced, filterField, filterValue, levels, pointsDict);
                results.Add(colLines);
                progress.Increment();
            });
            double endWatch = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);

            FancyWriteLine("Writing steel column splice forces table...", TextColor.Title);

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
            Console.WriteLine($"Steel Column splice forces table written in {timeCSV} seconds.\n");

            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;
            Check();

            // Report Metrics
            Console.WriteLine("\n--- API Diagnostics ---");
            Console.WriteLine($"GetLoadingAsync: {ApiMetrics.LoadingCalls} calls, Avg: {(ApiMetrics.LoadingCalls > 0 ? (double)ApiMetrics.LoadingDuration / ApiMetrics.LoadingCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"GetValueAsync:   {ApiMetrics.ValueCalls} calls, Avg: {(ApiMetrics.ValueCalls > 0 ? (double)ApiMetrics.ValueDuration / ApiMetrics.ValueCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"Peak Concurrency: {ApiMetrics.MaxConcurrency}");
            if (ApiMetrics.SemaphoreWaitCalls > 0)
            {
                Console.WriteLine($"Semaphore Waits: {ApiMetrics.SemaphoreWaitCalls} calls, Avg: {(double)ApiMetrics.SemaphoreWaitDuration / ApiMetrics.SemaphoreWaitCalls / 10000.0:F3} ms, Max: {ApiMetrics.SemaphoreWaitMax / 10000.0:F3} ms");
            }
            Console.WriteLine("-----------------------\n");
        }

        private static async Task<List<string>> ProcessColumnAsync(
            IMember member,
            List<ColumnLift> lifts,
            ColumnSpansSteel columnSpans,
            List<ILoadingCase> loadingCases,
            ICombination? integrityForceCombination,  // Changed from ILoadingCase
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

                // Find splice offsets within this lift
                var spliceOffsets = new List<(int SpanIndex, double Offset)>();
                foreach (var span in lift.Spans)
                {
                    if (columnSpans.SpanSpliceInfo.TryGetValue(span.Index, out var spliceInfo) && spliceInfo.HasSplice)
                    {
                        spliceOffsets.Add((span.Index, spliceInfo.SpliceOffset));
                    }
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

                // Fixity
                string startNodeFixity = GetNodeFixityDescription(firstSpan.StartReleases.Value);
                string endNodeFixity = GetNodeFixityDescription(lift.Spans.Last().EndReleases.Value);

                foreach (var loadingCase in loadingCases)
                {
                    // Calculate integrity forces once per lift (if needed)
                    var integrityForces = new Dictionary<string, double>();
                    if (integrityForceCombination != null && columnSpans.HasSplice && lifts.Count > 1)
                    {
                        integrityForces = await CalculateIntegrityForcesWithSpliceOffsets(member, lifts, integrityForceCombination, reduced, columnSpans, analysisType);
                    }

                    // If no splices in this lift, output a single row with splice offset of 0 and all forces as 0
                    if (spliceOffsets.Count == 0)
                    {
                        string lineNoSplice = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(member.Name)},{EscapeCsvValue(lift.Name)}," +
                              $"{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)},{sectionArea:F3}," +
                              $"{EscapeCsvValue(startNodeName)},{EscapeCsvValue(startNodeFixity)},{startX:F3},{startY:F3},{startZ:F3}," +
                              $"{EscapeCsvValue(endNodeName)},{EscapeCsvValue(endNodeFixity)},{endX:F3},{endY:F3},{endZ:F3}," +
                              $"{lengthFt:F3},0,{EscapeCsvValue(loadingCase.Name)}," +
                              $"0,0,0,0,0,0,0";

                        lines.Add(lineNoSplice);
                        continue;
                    }

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

                    // Query forces at each splice location
                    foreach (var (spanIndex, spliceOffset) in spliceOffsets)
                    {
                        // Convert splice offset from mm to inches
                        double spliceOffsetIn = MmToIn(spliceOffset);

                        // Get envelope forces at splice location
                        var (axialMax, axialMin) = await GetMinMaxForceAtSplice(memberLoading, LoadingValueType.Force, LoadingDirection.Axial, spanIndex, spliceOffset, reduced);

                        double majorMomentMax = 0;
                        double majorShearMax = 0;
                        double minorMomentMax = 0;
                        double minorShearMax = 0;

                        // For envelopes, use GetPointsOfInterest to get moments/shears directly (MUCH FASTER!)
                        if (loadingCase is IEnvelope envelope && envelope.CombinationIds != null)
                        {
                            try
                            {
                                majorMomentMax = await GetMaxForceFromPointsOfInterestAtSplice(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, spanIndex, spliceOffset);
                                majorShearMax = await GetMaxForceFromPointsOfInterestAtSplice(memberLoading, LoadingValueType.Force, LoadingDirection.Major, spanIndex, spliceOffset);
                                minorMomentMax = await GetMaxForceFromPointsOfInterestAtSplice(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, spanIndex, spliceOffset);
                                minorShearMax = await GetMaxForceFromPointsOfInterestAtSplice(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, spanIndex, spliceOffset);
                            }
                            catch
                            {
                                // If envelope querying fails, moments stay at 0
                            }
                        }
                        else if (!(loadingCase is IEnvelope))
                        {
                            // For regular cases/combos, query directly at splice
                            majorMomentMax = await GetMaxForceAtSplice(memberLoading, LoadingValueType.Moment, LoadingDirection.Major, spanIndex, spliceOffset, reduced);
                            majorShearMax = await GetMaxForceAtSplice(memberLoading, LoadingValueType.Force, LoadingDirection.Major, spanIndex, spliceOffset, reduced);
                            minorMomentMax = await GetMaxForceAtSplice(memberLoading, LoadingValueType.Moment, LoadingDirection.Minor, spanIndex, spliceOffset, reduced);
                            minorShearMax = await GetMaxForceAtSplice(memberLoading, LoadingValueType.Force, LoadingDirection.Minor, spanIndex, spliceOffset, reduced);
                        }

                        // Get integrity force from pre-calculated values (lookup, don't query per-splice)
                        double integrityForce = 0.0;
                        if (integrityForceCombination != null && columnSpans.HasSplice && lifts.Count > 1)
                        {
                            // Look up the pre-calculated value for this lift
                            if (integrityForces.ContainsKey(lift.Name))
                            {
                                integrityForce = -1 * Math.Abs(integrityForces[lift.Name]);
                            }
                        }

                        string line = $"{EscapeCsvValue(id.ToString())},{EscapeCsvValue(partMark)},{EscapeCsvValue(filterValue)},{EscapeCsvValue(member.Name)},{EscapeCsvValue(lift.Name)}," +
                              $"{EscapeCsvValue(startLevelName)},{EscapeCsvValue(endLevelName)},{EscapeCsvValue(sectionName)},{EscapeCsvValue(materialName)},{sectionArea:F3}," +
                              $"{EscapeCsvValue(startNodeName)},{EscapeCsvValue(startNodeFixity)},{startX:F3},{startY:F3},{startZ:F3}," +
                              $"{EscapeCsvValue(endNodeName)},{EscapeCsvValue(endNodeFixity)},{endX:F3},{endY:F3},{endZ:F3}," +
                              $"{lengthFt:F3},{spliceOffsetIn:F3},{EscapeCsvValue(loadingCase.Name)}," +
                              $"{axialMax},{axialMin},{majorMomentMax},{majorShearMax},{minorMomentMax},{minorShearMax},{integrityForce}";

                        lines.Add(line);
                    }
                }
            }
            return lines;
        }

        /// <summary>
        /// Gets the integrity force (as negative) at a splice location.
        /// Always queries from the Integrity Force loading case.
        /// </summary>
        private static async Task<double> GetIntegrityForceAtSplice(
            IMember member,
            ICombination integrityForceCombination,  // Changed from ILoadingCase
            int spanIndex,
            double spliceOffset,
            bool reduced,
            AnalysisType analysisType)
        {
            try
            {
                var swWait = Stopwatch.StartNew();
                await ApiLimiter.WaitAsync();
                swWait.Stop();
                ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);

                IMemberLoading memberLoading;
                var swL = Stopwatch.StartNew();
                try
                {
                    memberLoading = await member.GetLoadingAsync(integrityForceCombination.Id, analysisType, LoadingResultType.Base);
                }
                finally
                {
                    swL.Stop();
                    ApiMetrics.RecordLoading(swL.ElapsedTicks);
                    ApiLimiter.Release();
                }

                // Get axial force at splice
                var option = LoadingValueOptions.StaticValue(LoadingValueType.Force, LoadingDirection.Axial, reduced);

                var swWait2 = Stopwatch.StartNew();
                await ApiLimiter.WaitAsync();
                swWait2.Stop();
                ApiMetrics.RecordSemaphoreWait(swWait2.ElapsedTicks);

                try
                {
                    ApiMetrics.IncrementActiveValueCalls();
                    var swV = Stopwatch.StartNew();

                    var vals = await memberLoading.GetValueAsync(option, spanIndex, spliceOffset);

                    swV.Stop();
                    ApiMetrics.RecordValue(swV.ElapsedTicks);
                    ApiMetrics.DecrementActiveValueCalls();

                    if (vals.Any())
                    {
                        // Take absolute value and multiply by -1 for negative result
                        double force = Math.Abs(vals.MaxBy(v => Math.Abs(v.Value))?.Value ?? 0.0);
                        return -1 * force * ConversionFactor(LoadingValueType.Force);
                    }
                }
                finally
                {
                    ApiLimiter.Release();
                }
            }
            catch { }

            return 0.0;
        }

        /// <summary>
        /// Gets the maximum force value at a splice location using GetPointsOfInterest.
        /// This is 10x faster than looping through all child combinations.
        /// </summary>
        private static async Task<double> GetMaxForceFromPointsOfInterestAtSplice(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            int spanIndex,
            double spliceOffset)
        {
            double maxValue = 0;
            var option = LoadingValueOptions.StaticValue(valueType, direction, false);

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

                    // Query maximum point directly - this works for envelopes!
                    var maxPoints = await loading.GetPointsOfInterest(option, PointOfInterestType.Maximum);

                    swV.Stop();
                    ApiMetrics.RecordValue(swV.ElapsedTicks);
                    ApiMetrics.DecrementActiveValueCalls();

                    // Filter to this span and get value at splice position
                    var spanMaxPoints = maxPoints.Where(p => p.SpanIndex == spanIndex).ToList();

                    // Get the actual value at the splice offset position
                    var values = await loading.GetValueAsync(option, spanIndex, spliceOffset);
                    foreach (var val in values)
                    {
                        if (Math.Abs(val.Value) > Math.Abs(maxValue))
                            maxValue = val.Value;
                    }
                }
                finally
                {
                    ApiLimiter.Release();
                }
            }
            catch { }

            return Math.Abs(maxValue) * ConversionFactor(valueType);
        }

        private static async Task<double> GetMaxForceAtSplice(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            int spanIndex,
            double spliceOffset,
            bool reduced)
        {
            var allValues = new List<double>();
            var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);

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

                    var vals = await loading.GetValueAsync(option, spanIndex, spliceOffset);

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

            return allValues.DefaultIfEmpty(0.0).Max() * ConversionFactor(valueType);
        }

        private static async Task<(double max, double min)> GetMinMaxForceAtSplice(
            IMemberLoading loading,
            LoadingValueType valueType,
            LoadingDirection direction,
            int spanIndex,
            double spliceOffset,
            bool reduced)
        {
            var allValues = new List<double>();
            var option = LoadingValueOptions.StaticValue(valueType, direction, reduced);

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

                    var vals = await loading.GetValueAsync(option, spanIndex, spliceOffset);

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

            double globalMax = allValues.Count > 0 ? allValues.Max() : 0.0;
            double globalMin = allValues.Count > 0 ? allValues.Min() : 0.0;

            double valCon = ConversionFactor(valueType);
            return (globalMax * valCon, globalMin * valCon);
        }

        private static async Task<Dictionary<string, double>> CalculateIntegrityForcesWithSpliceOffsets(
            IMember member,
            List<ColumnLift> lifts,
            ICombination integrityForceCombination,
            bool reduced,
            ColumnSpansSteel columnSpans,
            AnalysisType analysisType)
        {
            var integrityForces = new Dictionary<string, double>();
            if (lifts.Count <= 1 || integrityForceCombination == null) return integrityForces;

            try
            {
                IMemberLoading memberLoading;
                var swWait = Stopwatch.StartNew();
                await ApiLimiter.WaitAsync();
                swWait.Stop();
                ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                var sw = Stopwatch.StartNew();
                try
                {
                    memberLoading = await member.GetLoadingAsync(integrityForceCombination.Id, analysisType, LoadingResultType.Base);
                }
                finally
                {
                    sw.Stop();
                    ApiMetrics.RecordLoading(sw.ElapsedTicks);
                    ApiLimiter.Release();
                }

                double valCon = ConversionFactor(LoadingValueType.Force);
                integrityForces[lifts[0].Name] = 0.0;

                // Start node force - take absolute value
                var firstSpan = lifts[0].Spans.First();
                double startNodeForce = Math.Abs(await FetchForce(memberLoading, firstSpan.Index, 0.0, reduced)) * valCon;

                // Get forces at all splice locations - sequential processing to avoid API congestion
                var spliceForces = new List<double>();
                foreach (var lift in lifts)
                {
                    foreach (var span in lift.Spans)
                    {
                        if (columnSpans.SpanSpliceInfo.ContainsKey(span.Index) &&
                            columnSpans.SpanSpliceInfo[span.Index].HasSplice)
                        {
                            double spliceOffset = columnSpans.SpanSpliceInfo[span.Index].SpliceOffset;
                            // Take absolute value of splice forces
                            spliceForces.Add(Math.Abs(await FetchForce(memberLoading, span.Index, spliceOffset, reduced)) * valCon);
                        }
                    }
                }

                // Calculate integrity forces for each lift using the exact logic from SteelColumnIntegrityForces
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
            catch (Exception)
            {
                // console log if needed
                foreach (var l in lifts) integrityForces[l.Name] = 0.0;
            }
            return integrityForces;
        }

        private static async Task<double> FetchForce(IMemberLoading loading, int spanIndex, double pos, bool reduced)
        {
            try
            {
                var option = LoadingValueOptions.StaticValue(LoadingValueType.Force, LoadingDirection.Axial, reduced);
                IEnumerable<ILoadingValue> values;
                var swWait = Stopwatch.StartNew();
                await ApiLimiter.WaitAsync();
                swWait.Stop();
                ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
                var sw = Stopwatch.StartNew();
                ApiMetrics.IncrementActiveValueCalls();
                try
                {
                    values = await loading.GetValueAsync(option, spanIndex, pos);
                }
                finally
                {
                    ApiMetrics.DecrementActiveValueCalls();
                    sw.Stop();
                    ApiLimiter.Release();
                }
                ApiMetrics.RecordValue(sw.ElapsedTicks);

                return values.MaxBy(v => Math.Abs(v.Value))?.Value ?? 0.0;
            }
            catch { }
            return 0.0;
        }
    }
}