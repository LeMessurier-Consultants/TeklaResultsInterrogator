﻿using System.Diagnostics;
using System.Text;
using TSD.API.Remoting;
using TSD.API.Remoting.Document;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;
using TSD.API.Remoting.Common;
using AnalysisType = TSD.API.Remoting.Solver.AnalysisType;
using Google.Protobuf.WellKnownTypes;
using MathNet.Numerics.Providers.SparseSolver;
using TeklaResultsInterrogator.Utils;
using static TeklaResultsInterrogator.Utils.Utils;

namespace TeklaResultsInterrogator.Core
{
    public class SolverInterrogator : BaseInterrogator
    {
        
        protected TSD.API.Remoting.Solver.IModel? SolverModel { get; private set; }
        protected AnalysisType RequestedAnalysisType { get; private set; }
        protected List<ILoadcase>? AllLoadcases { get; private set; }
        protected List<ILoadcase>? SolvedCases { get; private set; }
        protected List<ICombination>? AllCombinations { get; private set; }
        protected List<ICombination>? SolvedCombinations { get; private set; }
        protected List<IEnvelope>? AllEnvelopes { get; private set; }
        protected List<IEnvelope>? SolvedEnvelopes { get; private set; }
        protected List<IMember>? AllMembers { get; set; }

        protected List<MemberConstruction> RequestedMemberType = new List<MemberConstruction>();

        public SolverInterrogator() { }

        public override async Task InitializeAsync()  // to get solver model and other stuff here
        {
            // Set up
            Stopwatch stopwatch = Stopwatch.StartNew();
            await InitializeBaseAsync();

            // Get Model
            Console.WriteLine("Searching for analysis solver model...");
            if (Model == null)
            {
                FancyWriteLine("No model found!", TextColor.Error);
                Flag = true;
                return;
            }
            
            AnalysisType defaultAnalysisType = AnalysisType.FirstOrderLinear;// Default

            List<AnalysisType> solvedAnalysisTypes = new List<AnalysisType>();

            foreach (AnalysisType analysisType in System.Enum.GetValues(typeof(AnalysisType)))
            {
                if (analysisType!=AnalysisType.Unknown && analysisType!=AnalysisType.None)
                {
                    
                    IEnumerable<TSD.API.Remoting.Solver.IModel> solverModels2 = await Model.GetSolverModelsAsync(new[] { analysisType });
                    if (solverModels2.Any())
                    {
                        //FancyWriteLine(analysisType.ToString(), TextColor.Error);
                        solvedAnalysisTypes.Add(analysisType);
                    }
                }
            }

            int count = 0;

            foreach (AnalysisType analysisType in solvedAnalysisTypes)
            {
                count = count + 1;
                FancyWriteLine(count.ToString() + " - " + analysisType.ToString(), TextColor.Text);
            }
            string? readIn = AskUser("Input a number or hit Enter to use default 1st Order Linear. ");
            bool keepReading = int.TryParse(readIn, out int intOption);
            if (readIn != null && readIn != "" && keepReading == true)
            {
                RequestedAnalysisType = solvedAnalysisTypes[intOption-1];
            }
            else 
            {
                RequestedAnalysisType = defaultAnalysisType;
            }

            IEnumerable<TSD.API.Remoting.Solver.IModel> solverModels = await Model.GetSolverModelsAsync(new[] { defaultAnalysisType });
            if (!solverModels.Any())
            {
                FancyWriteLine("No solver models found!", TextColor.Error);
                Flag = true;
                return;
            }

            TSD.API.Remoting.Solver.IModel? solverModel = solverModels.FirstOrDefault();
            if (solverModel == null)
            {
                FancyWriteLine("No solver model found!", TextColor.Error);
                Flag = true;
                return;
            }
            SolverModel = solverModel;

            // Get Analysis Results
            Console.WriteLine("Searching for analysis results...");
            IAnalysisResults? solverResults = await SolverModel.GetResultsAsync();
            if (solverResults == null)
            {
                FancyWriteLine("No results found for requested analysis type!", TextColor.Error);
                Flag = true;
                return;
            }
            IAnalysis3DResults? analysis3Dresults = await solverResults.GetAnalysis3DAsync();
            if (analysis3Dresults == null)
            {
                FancyWriteLine("No 3-D analysis results found for requested analysis type!", TextColor.Error);
                Flag = true;
                return;
            }
            var solvedLoadingGuids = await analysis3Dresults.GetSolvedLoadingIdsAsync();
            if (!solvedLoadingGuids.Any())
            {
                FancyWriteLine("No solved loading GUIDs found!", TextColor.Error);
                Flag = true;
                return;
            }

            // Get members
            Console.WriteLine("Searching for members...");
            IEnumerable<IMember>? allMembers = await Model.GetMembersAsync(null);
            if (allMembers == null || !allMembers.Any())
            {
                FancyWriteLine("No members found!", TextColor.Error);
                Flag = true;
                return;
            }
            AllMembers = allMembers.ToList();

            // Get solved loadcases
            Console.WriteLine("Searching for solved loadcases...");
            IEnumerable<ILoadcase> loadingCases = await Model.GetLoadcasesAsync(null);
            if (!loadingCases.Any())
            {
                FancyWriteLine("No loadcases found!", TextColor.Warning);
            }
            AllLoadcases = loadingCases.Where(lc => lc.Name != "0 ").ToList();  // Eliminating "0" slab unit load and roof unit load loadcases
            List<ILoadcase> solvedCases = AllLoadcases.Where(c => solvedLoadingGuids.Contains(c.Id)).ToList();
            if (!solvedCases.Any())
            {
                FancyWriteLine("No solved loadcases found!", TextColor.Warning);
            }
            SolvedCases = solvedCases;

            // Get solved load combos
            Console.WriteLine("Searching for solved load combinations...");
            IEnumerable<ICombination> loadingCombinations = await Model.GetCombinationsAsync(null);
            if (!loadingCombinations.Any())
            {
                FancyWriteLine("No load combinations found!", TextColor.Warning);
            }
            AllCombinations = loadingCombinations.ToList();
            List<ICombination> solvedCombinations = loadingCombinations.Where(c => solvedLoadingGuids.Contains(c.Id)).ToList();
            if (!solvedCombinations.Any())
            {
                FancyWriteLine("No solved load combinations found!", TextColor.Warning);
            }
            SolvedCombinations = solvedCombinations;

            // Get solved load envelopes
            Console.WriteLine("Searching for solved load envelopes...");
            IEnumerable<IEnvelope> loadingEnvelopes = await Model.GetEnvelopesAsync(null);
            if (!loadingEnvelopes.Any())
            {
                FancyWriteLine("No load envelopes found!", TextColor.Warning);
                Flag = false;  // Do not raise flag for no envelopes
            }
            AllEnvelopes = loadingEnvelopes.ToList();
            List<IEnvelope> solvedEnvelopes = new List<IEnvelope>();
            foreach (IEnvelope envelope in AllEnvelopes)
            {
                List<TSD.API.Remoting.Common.Properties.IReadOnlyProperty<Guid>> combinationIds = envelope.CombinationIds.ToList();
                if (combinationIds.All(id => solvedLoadingGuids.Contains(id.Value)))
                {
                    solvedEnvelopes.Add(envelope);
                }
            }
            if (!solvedEnvelopes.Any())
            {
                FancyWriteLine("No solved load envelopes found!", TextColor.Warning);
            }
            SolvedEnvelopes = solvedEnvelopes;

            // Check to make sure there are solved load cases
            if (AllLoadcases.Count + AllCombinations.Count + AllEnvelopes.Count == 0)
            {
                FancyWriteLine("No loading found!", TextColor.Error);
                Flag = true;
                return;
            }
            if (SolvedCases.Count + SolvedCombinations.Count + SolvedEnvelopes.Count == 0)
            {
                FancyWriteLine("No solved loading found!", TextColor.Error);
                Flag = true;
                return;
            }

            // Finish up
            stopwatch.Stop();
            InitializationTime = stopwatch.Elapsed.TotalSeconds;
            Console.WriteLine($"Initialization completed in {Math.Round(InitializationTime, 3)} seconds.\n");

            Check();

            return;
        }




        public List<ILoadingCase> AskLoading(List<ILoadcase> solvedCases, List<ICombination> solvedCombinations, List<IEnvelope> solvedEnvelopes)
        {
            Dictionary<string, List<ILoadingCase>> loadingOptions = new Dictionary<string, List<ILoadingCase>>(StringComparer.InvariantCultureIgnoreCase);

            if (solvedCases.Count > 0)
            {
                loadingOptions.Add("Cases", solvedCases.Cast<ILoadingCase>().ToList());
            }
            if (solvedCombinations.Count > 0)
            {
                loadingOptions.Add("Combos", solvedCombinations.Cast<ILoadingCase>().ToList());
            }
            if (solvedEnvelopes.Count > 0)
            {
                loadingOptions.Add("Envelopes", solvedEnvelopes.Cast<ILoadingCase>().ToList());
            }

            List<ILoadingCase>? loadingCases = null;

            FancyWriteLine("Available loading conditions:", TextColor.Text);
            foreach (string condition in loadingOptions.Keys)
            {
                FancyWriteLine($"  {condition}", TextColor.Command);
            }

            do
            {
                string? readIn = AskUser("Choose an available loading condition: ");
                if (readIn != null && loadingOptions.ContainsKey(readIn))
                    {

                    loadingCases = loadingOptions[readIn];

                    FancyWriteLine("Available loading:", TextColor.Text);
                    foreach (var load in loadingCases.OrderBy(o => o.ReferenceIndex))
                    {
                        FancyWriteLine(load.Name, TextColor.Text);
                    }
                    readIn = AskUser("Input a number or hit Enter to get all: ");
                    if (readIn != null && readIn != "")
                    {
                        loadingCases = loadingCases.Where(load => load.ReferenceIndex.Equals(Convert.ToInt32(readIn))).ToList(); // This is a bad way of doing this because I am not checking if the integer is valid, but it let's not let perfect be the enemy of the good!
                    }
                }
                else
                {
                    FancyWriteLine("Loading Condition", $" {readIn}", " not found.", TextColor.Command);
                }
            } while (loadingCases == null);

            return loadingCases;
        }

        public bool? AskGravityOnly()
        {
            bool? GravityOnly = null;

            string? readIn = AskUser("Enter Y to query Gravity Only members, or N to query Lateral members, hit Enter to get All");
            do
            {
                if (readIn == "Y")
                {
                    GravityOnly = true;
                }
                else if (readIn == "N")
                {
                    GravityOnly = false;
                }
                else if (readIn == "")
                {
                    return null;
                }
                else
                {
                    FancyWriteLine("Input ", $"{readIn}", " not recognized. All members will be returned", TextColor.Command);
                }

            } while (GravityOnly == null);

            return (bool?)GravityOnly;
        }

        public bool? AskAutoDesign()
        {
            bool? AutoDesign = null;

            string? readIn = AskUser("Enter Y to query Autodesign members only, or N to query Non-Autodesign members, hit Enter to get All");
            if (readIn == "Y")
            {
                AutoDesign = true;
            }
            else if (readIn == "N")
            {
                AutoDesign = false;
            }
            else if (readIn == "")
            {
                AutoDesign = null;
            }
            else
            {
                FancyWriteLine("Input ", $"{readIn}", " not recognized. All members will be returned", TextColor.Command);
            }

            return (bool?)AutoDesign;
        }


        public int AskPoints(int maxPoints)
        {
            FancyWriteLine("Select the number of points along beam span at which forces and displacements will be calculated.", TextColor.Text);
            FancyWriteLine("Enter ", "1", " to return maxima only.", TextColor.Command);
            if (maxPoints >= 2)
            {
                FancyWriteLine("Enter ", "2", $" or greater (max. {maxPoints}) to subdivide spans.", TextColor.Command);
            }
            FancyWriteLine("Enter ", "0", " to ignore force and displacement data.", TextColor.Command);

            int numPoints = -1;

            do
            {
                string? readIn = AskUser($"Enter an integer between 0 and {maxPoints}: ");
                if (int.TryParse(readIn, out _))
                {
                    numPoints = int.Parse(readIn);
                }
                if (numPoints < 0 | numPoints > maxPoints)
                {
                    FancyWriteLine("Illegal input: ", $"{readIn}", "", TextColor.Command);
                }
            } while (numPoints < 0 | numPoints > maxPoints);

            return numPoints;
        }

        public bool AskReduced()
        {
            bool? reduced = null;
            do
            {
                string? readIn = AskUser("Enter N to query nonreduced forces, or hit Enter to get reduced forces (where applicable): ");
                if (readIn == "")
                {
                    reduced = true;
                }
                else if (readIn == "N")
                {
                    reduced = false;
                }
                else
                {
                    FancyWriteLine("Input ", $"{readIn}", " not recognized.", TextColor.Command);
                }
            } while (reduced == null);
            return (bool)reduced;
        }

    }
}
