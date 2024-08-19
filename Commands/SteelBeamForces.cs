﻿using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using TeklaResultsInterrogator.Core;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;
using TSD.API.Remoting.Sections;
using TeklaResultsInterrogator.Utils;
using static TeklaResultsInterrogator.Utils.Utils;
using System.Diagnostics.Metrics;
using System.Xml.Linq;
using System.Reflection.PortableExecutable;

namespace TeklaResultsInterrogator.Commands
{
    public class SteelBeamForces : ForceInterrogator
    {
        public SteelBeamForces()
        {
            HasOutput = true;
            RequestedMemberType = new List<MemberConstruction>() { MemberConstruction.SteelBeam, MemberConstruction.CompositeBeam};
            
        }

        string GetMemberLevelNameAsync(IMember member)
        {
            int constructionPointIndex = member.MemberNodes.Value.First().Value.ConstructionPointIndex.Value;
            IEnumerable<IConstructionPoint> constructionPoints = Model!.GetConstructionPointsAsync(new List<int>() { constructionPointIndex }).Result;
            int planeId = constructionPoints.First().PlaneInfo.Value.Index;
            IEnumerable<IHorizontalConstructionPlane> levels = Model.GetLevelsAsync(new List<int>() { planeId }).Result;
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
        async Task<List<string>> GetMemberSpanInfoAsync(String levelName, IMember member, IMemberSpan span, int subdivisions, List<ILoadingCase> loadingCases,  Boolean reduced)
        {
            List<string> output = new List<string>();
            Guid id = member.Id;
            string name = member.Name;
            string spanName = span.Name;
            int spanIdx = span.Index;
            double length = span.Length.Value;
            double lengthFt = length * 0.00328084; // Converting from [mm] to [ft]
            double rot = Math.Round(span.RotationAngle.Value * 57.2958, 3); // Converting from [rad] to [deg]
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
                List<Task<MaxSpanInfo>> maxSpanInfotasks = new List<Task<MaxSpanInfo>>();
                List<Task<List<PointSpanInfo>>> pointSpanInfotasks = new List<Task<List<PointSpanInfo>>>();

                foreach (ILoadingCase loadingCase in loadingCases)
                {
                    string loadName = loadingCase.Name.Replace(',', '`');
                    SpanResults spanResults = new SpanResults(span, subdivisions, loadingCase, reduced, AnalysisType, member);

                    if (subdivisions >= 1)
                    {
                        // Getting maximum internal forces and displacements and locations in parallel
                        maxSpanInfotasks.Add(Task.Run(() => spanResults.GetMaxima()));
                    }

                    if (subdivisions >= 2)
                    {
                        // Getting internal forces and displacements at each station in parallel
                        pointSpanInfotasks.Add(Task.Run(() => spanResults.GetStations()));
                    }
                }
                // Wait for all maxSpanInfo tasks to finish
                MaxSpanInfo[] completedMaxSpanTasks = await Task.WhenAll(maxSpanInfotasks); //execute in parallel
                // Wait for all pointSpanInfo tasks to finish
                List<PointSpanInfo>[] completedPointSpanInfoTasks = await Task.WhenAll(pointSpanInfotasks); //execute in parallel

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
        public override async Task ExecuteAsync()
        {
            // Initialize parents
            await InitializeAsync();

            // Check for null properties
            if (Flag)
            {
                return;
            }

            // Data setup and diagnostics initialization
            Stopwatch stopwatch = Stopwatch.StartNew();
            int bufferSize = 65536*2;

            // Unpacking loading data
            FancyWriteLine("Loading Summary:", TextColor.Title);
            Console.WriteLine("Unpacking loading data...");
            Console.WriteLine($"{AllLoadcases!.Count} loadcases found, {SolvedCases!.Count} solved.");
            Console.WriteLine($"{AllCombinations!.Count} load combinations found, {SolvedCombinations!.Count} solved.");
            Console.WriteLine($"{AllEnvelopes!.Count} load envelopes found, {SolvedEnvelopes!.Count} solved.\n");

            stopwatch.Stop();
            List<ILoadingCase> loadingCases = AskLoading(SolvedCases, SolvedCombinations, SolvedEnvelopes);
            bool reduced = AskReduced();
            stopwatch.Start();

            // Unpacking member data
            FancyWriteLine("\nMember summary:", TextColor.Title);
            Console.WriteLine("Unpacking member data...");

            List<IMember> steelBeams = new List<IMember>();

            bool? GravityOnlyState = AskGravityOnly();
            bool? AutoDesignState = AskAutoDesign();
            if (GravityOnlyState == null & AutoDesignState == null)
            {
                steelBeams = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction))).ToList();
            }
            else if (AutoDesignState == null)
            {
                steelBeams = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction)) & GetProperty(c.Data.Value.GravityOnly) == GravityOnlyState).ToList();
            }
            else if (GravityOnlyState == null)
            {
                steelBeams = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction)) & GetProperty(c.Data.Value.AutoDesign) == AutoDesignState).ToList();
            }
            else
            {
                steelBeams = AllMembers!.Where(c => RequestedMemberType.Contains(GetProperty(c.Data.Value.Construction)) & GetProperty(c.Data.Value.AutoDesign) == AutoDesignState & GetProperty(c.Data.Value.GravityOnly) == GravityOnlyState).ToList();
            };

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
            double start1 = timeUnpack;
            string file1 = SaveDirectory + @"SteelBeamForces_" + FileName + ".csv";
            string header1 = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}\n",
                "Tekla GUID", "Member Name", "Level", "Shape", "Material", "Span Name",
                "Start Node", "Start Node Fixity", "End Node", "End Node Fixity",
                "Span Length [ft]", "Span Rotation [deg]",
                "Loading Name", "Position [ft]",
                "Shear Major [k]", "Shear Minor [k]", "Moment Major [k-ft]", "Moment Minor [k-ft]",
                "Axial Force [k]", "Torsion [k-ft]", "Deflection Major [in]", "Deflection Minor [in]");
            File.WriteAllText(file1, "");
            File.AppendAllText(file1, header1);

            List<string>[] completedTaskOutput;
            List<Task<List<string>>> tasks = new();
            foreach (IMember member in steelBeams)
            {
                string levelName = GetMemberLevelNameAsync(member);
                IEnumerable<IMemberSpan> spans = await member.GetSpanAsync(); //should not be long running so await 
                foreach (IMemberSpan span in spans)
                {
                    tasks.Add(Task.Run(() => GetMemberSpanInfoAsync(levelName, member, span, subdivisions, loadingCases, reduced)));
                } 
            }
            completedTaskOutput = await Task.WhenAll(tasks); //execute in parallel
            // Getting internal forces and writing table
            FancyWriteLine("\nWriting internal forces table...", TextColor.Title);
            using (StreamWriter sw1 = new StreamWriter(file1, true, Encoding.UTF8, bufferSize))
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
            double time1 = Math.Round(stopwatch.Elapsed.TotalSeconds - start1, 3);
            Console.WriteLine($"Steel Beam table written in {time1} seconds.\n");

            // Finish up
            stopwatch.Stop();
            ExecutionTime = stopwatch.Elapsed.TotalSeconds;

            Check();

            return;
        }
    }
}
