using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeklaResultsInterrogator.Core;
using TeklaResultsInterrogator.Utils;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Sections;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;
using static TeklaResultsInterrogator.Utils.ConsoleUtils;

namespace TeklaResultsInterrogator.Commands
{
    /// <summary>
    /// Interrogates Steel Beam forces, calculating maximums and values at stations.
    /// </summary>
    public class SteelBeamForces : SolverInterrogator
    {
        /// <inheritdoc/>
        public override bool ShowInMenu() { return true; }

        /// <summary>Initializes a new instance of the <see cref="SteelBeamForces"/> class.</summary>
        public SteelBeamForces()
        {
            HasOutput = true;
            RequestedMemberType = [MemberConstruction.SteelBeam, MemberConstruction.CompositeBeam];
        }

        string GetMemberLevelNameAsync(IMember member)
        {
            int constructionPointIndex = member.MemberNodes.Value.First().Value.ConstructionPointIndex.Value;
            IEnumerable<IConstructionPoint> constructionPoints = Model!.GetConstructionPointsAsync([constructionPointIndex]).Result;
            int planeId = constructionPoints.First().PlaneInfo.Value.Index;
            IEnumerable<IHorizontalConstructionPlane> levels = Model.GetLevelsAsync([planeId]).Result;
            string levelName;
            if (levels.Any())
            {
                levelName = levels.First().Name;
            }
            else
            {
                levelName = "Not Associated";
            }
            return levelName;
        }
        async Task<List<string>> GetMemberSpanInfoAsync(String levelName, IMember member, IMemberSpan span, int subdivisions, List<ILoadingCase> loadingCases, Boolean reduced)
        {
            List<string> output = [];
            Guid id = member.Id;
            string name = member.Name;
            string spanName = span.Name;

            double length = span.Length.Value;
            double lengthFt = MmToFt(length); // Converting from [mm] to [ft]
            double rot = Math.Round(RadToDeg(span.RotationAngle.Value), 3); // Converting from [rad] to [deg]
            IMemberSection section = (IMemberSection)span.ElementSection.Value;
            string sectionName = section.PhysicalSection.Value.LongName;
            string materialGrade = span.Material.Value.Name;

            int startNodeIdx = span.StartMemberNode.ConstructionPointIndex.Value;
            string startNodeFixity = span.StartReleases.Value.DegreeOfFreedom.Value.ToString();
            if (GetProperty(span.StartReleases.Value.Cantilever) == true)
            {
                startNodeFixity += " (Cantilever end)";
            }
            startNodeFixity = startNodeFixity.Replace(',', '|');
            int endNodeIdx = span.EndMemberNode.ConstructionPointIndex.Value;
            string endNodeFixity = span.EndReleases.Value.DegreeOfFreedom.Value.ToString();
            if (GetProperty(span.EndReleases.Value.Cantilever) == true)
            {
                endNodeFixity += " (Cantilever end)";
            }
            endNodeFixity = endNodeFixity.Replace(',', '|');

            string spanLineOnly = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}",
                id, name, levelName, sectionName, materialGrade, spanName,
                startNodeIdx, startNodeFixity, endNodeIdx, endNodeFixity, lengthFt, rot);

            if (subdivisions == 0)
            {
                output.Add(spanLineOnly);
            }

            else
            {
                List<MaxSpanInfo> completedMaxSpanTasks = [];
                List<List<PointSpanInfo>> completedPointSpanInfoTasks = [];

                // Sequential processing to avoid API congestion
                foreach (ILoadingCase loadingCase in loadingCases)
                {

                    SpanResults spanResults = new(span, subdivisions, loadingCase, reduced, RequestedAnalysisType, member);

                    if (subdivisions >= 1)
                    {
                        // Getting maximum internal forces and displacements
                        completedMaxSpanTasks.Add(await spanResults.GetMaxima());
                    }

                    if (subdivisions >= 2)
                    {
                        // Getting internal forces and displacements at each station
                        completedPointSpanInfoTasks.Add(await spanResults.GetStations());
                    }
                }

                //Process MaxSpanInfo results
                foreach (MaxSpanInfo task in completedMaxSpanTasks)
                {
                    MaxSpanInfo maxSpanInfo = task;
                    string maxLine = spanLineOnly + "," + String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}",
                           maxSpanInfo.LoadName, "MAXIMA",
                           maxSpanInfo.ShearMajor.Value,
                           maxSpanInfo.ShearMinor.Value,
                           maxSpanInfo.MomentMajor.Value,
                           maxSpanInfo.MomentMinor.Value,
                           maxSpanInfo.AxialForce.Value,
                           maxSpanInfo.Torsion.Value,
                           maxSpanInfo.DeflectionMajor.Value,
                           maxSpanInfo.DeflectionMinor.Value);
                    output.Add(maxLine);
                }
                //Process MaxSpanInfo results
                foreach (List<PointSpanInfo> pointSpanInfoList in completedPointSpanInfoTasks)
                {
                    foreach (PointSpanInfo info in pointSpanInfoList)
                    {
                        string posLine = spanLineOnly + "," + String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}",
                                            info.LoadName, info.Position, info.ShearMajor, info.ShearMinor, info.MomentMajor, info.MomentMinor,
                                            info.AxialForce, info.Torsion, info.DeflectionMajor, info.DeflectionMinor);
                        output.Add(posLine);
                    }
                }
            }
            return output;
        }
        /// <summary>
        /// Executes the Steel Beam Forces interrogation, writing results to CSV.
        /// </summary>
        public override async Task ExecuteAsync()
        {
            // Initialize parents
            await InitializeAsync();

            if (Flag) return;

            // Reset API metrics for this command run
            ApiMetrics.Reset();

            // Data setup and diagnostics initialization
            Stopwatch stopwatch = Stopwatch.StartNew();
            int bufferSize = 65536 * 2;

            // Unpacking loading data
            LogLoadingSummary();

            stopwatch.Stop();
            List<ILoadingCase> loadingCases = AskLoading(SolvedCases, SolvedCombinations, SolvedEnvelopes);
            bool reduced = AskReduced();

            List<IMember> steelBeams = AskAndFilterMembers(true, true);

            stopwatch.Start();
            Console.WriteLine($"{AllMembers!.Count} structural members found in model.");
            Console.WriteLine($"{steelBeams.Count} steel beams found.");

            double timeUnpack = Math.Round(stopwatch.Elapsed.TotalSeconds, 3);
            Console.WriteLine($"Loading and member data unpacked in {timeUnpack} seconds.\n");

            // Extracting internal forces
            FancyWriteLine("Retrieving internal forces...", TextColor.Title);
            stopwatch.Stop();
            int subdivisions = AskPoints(20);  // Setting maximum number of stations to 20
            stopwatch.Start();
            FancyWriteLine($"Asked for {subdivisions} points.", TextColor.Warning);

            // Setting up file
            string file1 = SaveDirectory + @"SteelBeamForces_" + OutputFileName + ".csv";
            string header1 = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}\n",
                "Tekla GUID", "Member Name", "Level", "Shape", "Material", "Span Name",
                "Start Node", "Start Node Fixity", "End Node", "End Node Fixity",
                "Span Length [ft]", "Span Rotation [deg]",
                "Loading Name", "Position [ft]",
                "Shear Major [k]", "Shear Minor [k]", "Moment Major [k-ft]", "Moment Minor [k-ft]",
                "Axial Force [k]", "Torsion [k-ft]", "Deflection Major [in]", "Deflection Minor [in]");
            File.WriteAllText(file1, header1);

            List<string>[] completedTaskOutput;

            // Use global parallelism setting now that API usage is throttled via ApiLimiter
            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism };
            var results = new System.Collections.Concurrent.ConcurrentBag<List<string>>();


            // Phase 1: Collect span data (Parallel) with Throttling & Error Handling
            FancyWriteLine("Collecting span data...", TextColor.Title);
            var spanData = new System.Collections.Concurrent.ConcurrentBag<(IMember member, IMemberSpan span, string levelName)>();
            var exceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();

            using (var collectProgress = new ProgressBar(steelBeams.Count))
            {
                // Re-enable parallelism but respect the global API limit (Safety First!)
                await Parallel.ForEachAsync(steelBeams, parallelOptions, async (member, token) =>
                {
                    try
                    {
                        // Throttle API calls in Phase 1
                        await ApiLimiter.WaitAsync(token);
                        string levelName;
                        IEnumerable<IMemberSpan> spans;
                        try
                        {
                            levelName = GetMemberLevelNameAsync(member);
                            spans = await member.GetSpanAsync(cancellationToken: token);
                        }
                        finally
                        {
                            ApiLimiter.Release();
                        }

                        foreach (IMemberSpan span in spans)
                        {
                            spanData.Add((member, span, levelName));
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                        // Don't crash the whole loop, just log internally and continue
                    }
                    finally
                    {
                        collectProgress.Increment();
                    }
                });
            }

            if (!exceptions.IsEmpty)
            {
                FancyWriteLine($"\nWarning: {exceptions.Count} members failed to collect span data.", TextColor.Warning);
                foreach (var ex in exceptions.Take(5))
                {
                    FancyWriteLine($"Error: {ex.Message}", TextColor.Error);
                }
                if (exceptions.Count > 5) FancyWriteLine($"...and {exceptions.Count - 5} more.", TextColor.Error);
            }

            // Phase 2: Process forces (Parallel)
            var collectedSpans = spanData.ToList();
            FancyWriteLine($"Processing forces for {collectedSpans.Count} spans...", TextColor.Title);
            using var progress = new ProgressBar(collectedSpans.Count);

            // Adjusted parallelism to match API Limit to prevent 'batching' pauses
            // Previous setting (ProcCount * 8 = 128) caused consistent 0.75s pauses every 128 items.
            var smoothedParallelOptions = new ParallelOptions { MaxDegreeOfParallelism = SolverInterrogator.MaxDegreeOfParallelism, CancellationToken = parallelOptions.CancellationToken };

            await Parallel.ForEachAsync(collectedSpans, smoothedParallelOptions, async (item, token) =>
            {
                var spanLines = await GetMemberSpanInfoAsync(item.levelName, item.member, item.span, subdivisions, loadingCases, reduced);
                results.Add(spanLines);
                progress.Increment();
            });

            completedTaskOutput = [.. results];
            // Getting internal forces and writing table
            FancyWriteLine("\nWriting internal forces table...", TextColor.Title);
            double writeStart = stopwatch.Elapsed.TotalSeconds;
            using (StreamWriter sw1 = new(file1, true, Encoding.UTF8, bufferSize))
            {
                //process output to streamwriter
                foreach (List<string> outputLines in completedTaskOutput)
                {
                    foreach (string outputLine in outputLines)
                    {
                        sw1.WriteLine(outputLine);
                    }
                }
            }

            // Output diagnostics to console
            FancyWriteLine("Saved to: ", file1, "", TextColor.Path);
            double size1 = Math.Round((double)new FileInfo(file1).Length / 1024, 2);
            Console.WriteLine($"File size: {size1} KB");
            double writeTime = Math.Round(stopwatch.Elapsed.TotalSeconds - writeStart, 3);
            Console.WriteLine($"Steel Beam table written in {writeTime} seconds.\n");

            // Finish up
            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;

            Check();

            // Report Metrics
            Console.WriteLine("\n--- API Diagnostics ---");
            Console.WriteLine($"GetLoadingAsync:     {ApiMetrics.LoadingCalls} calls, Avg: {(ApiMetrics.LoadingCalls > 0 ? (double)ApiMetrics.LoadingDuration / ApiMetrics.LoadingCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"GetPointsOfInterest: {ApiMetrics.PoiCalls} calls, Avg: {(ApiMetrics.PoiCalls > 0 ? (double)ApiMetrics.PoiDuration / ApiMetrics.PoiCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"GetValueAsync:       {ApiMetrics.ValueCalls} calls, Avg: {(ApiMetrics.ValueCalls > 0 ? (double)ApiMetrics.ValueDuration / ApiMetrics.ValueCalls / 10000.0 : 0):F3} ms");
            Console.WriteLine($"Peak Concurrency:    {ApiMetrics.MaxConcurrency}");
            if (ApiMetrics.SemaphoreWaitCalls > 0)
            {
                Console.WriteLine($"Semaphore Waits:     {ApiMetrics.SemaphoreWaitCalls} calls, Avg: {(double)ApiMetrics.SemaphoreWaitDuration / ApiMetrics.SemaphoreWaitCalls / 10000.0:F3} ms (Max: {ApiMetrics.SemaphoreWaitMax / 10000.0:F3} ms)");
            }
            Console.WriteLine("-----------------------\n");

            return;
        }
    }
}
