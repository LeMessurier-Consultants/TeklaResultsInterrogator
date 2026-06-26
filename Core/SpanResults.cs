using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MathNet.Numerics;
using TeklaResultsInterrogator.Core;
using TeklaResultsInterrogator.Utils;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;
using static TeklaResultsInterrogator.Utils.ConsoleUtils;

namespace TeklaResultsInterrogator.Core
{


    /// <summary>
    /// Encapsulates results for a single member span, handling maxima and station-based data retrieval.
    /// </summary>
    public class SpanResults
    {
        /// <summary>Name of the span.</summary>
        public string Name { get; set; }
        /// <summary>Length of the span (mm).</summary>
        public double Length { get; set; }
        /// <summary>Whether to use reduced results.</summary>
        public bool Reduced { get; set; }
        /// <summary>Number of subdivisions/stations for detailed analysis.</summary>
        public int Subdivisions { get; set; }
        /// <summary>Associated loading case.</summary>
        public ILoadingCase Loading { get; set; }
        /// <summary>Associated analysis type.</summary>
        public AnalysisType AnalysisType { get; set; }
        /// <summary>The TSD span object.</summary>
        public IMemberSpan Span { get; set; }
        /// <summary>The parent member of the span.</summary>
        public IMember ParentMember { get; set; }

        private LoadingValueOptions ShearMajorValueOption { get; set; }
        private LoadingValueOptions ShearMinorValueOption { get; set; }
        private LoadingValueOptions MomentMajorValueOption { get; set; }
        private LoadingValueOptions MomentMinorValueOption { get; set; }
        private LoadingValueOptions AxialValueOption { get; set; }
        private LoadingValueOptions TorsionValueOption { get; set; }
        private LoadingValueOptions DeflectionMajorValueOption { get; set; }
        private LoadingValueOptions DeflectionMinorValueOption { get; set; }
        private LoadingValueOptions DisplacementMajorValueOption { get; set; }
        private LoadingValueOptions DisplacementMinorValueOption { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="SpanResults"/> class.
        /// </summary>
        /// <param name="span">The member span.</param>
        /// <param name="subdivisions">Number of subdivisions.</param>
        /// <param name="loading">Loading case.</param>
        /// <param name="reduced">Use reduced forces.</param>
        /// <param name="analysisType">Analysis type.</param>
        /// <param name="parentMember">Parent member.</param>
        public SpanResults(IMemberSpan span, int subdivisions, ILoadingCase loading, bool reduced, AnalysisType analysisType, IMember parentMember)
        {
            Span = span;
            Name = Span.Name;
            Length = Span.Length.Value;  // Note this is in base units [mm]
            Reduced = reduced;
            Subdivisions = subdivisions;
            Loading = loading;
            AnalysisType = analysisType;
            ParentMember = parentMember;

            ShearMajorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Force, LoadingDirection.Major, Reduced);
            ShearMinorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Force, LoadingDirection.Minor, Reduced);
            MomentMajorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Moment, LoadingDirection.Major, Reduced);
            MomentMinorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Moment, LoadingDirection.Minor, Reduced);
            AxialValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Force, LoadingDirection.Axial, Reduced);
            TorsionValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Moment, LoadingDirection.Axial, Reduced);
            DeflectionMajorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Deflection, LoadingDirection.Major, Reduced);
            DeflectionMinorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Deflection, LoadingDirection.Minor, Reduced);
            DisplacementMajorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Displacement, LoadingDirection.Major, Reduced);
            DisplacementMinorValueOption = LoadingValueOptions.StaticValue(LoadingValueType.Displacement, LoadingDirection.Minor, Reduced);
        }

        /// <summary>
        /// Calculates and retrieves the maximum values for forces and displacements along the span.
        /// </summary>
        /// <returns>A <see cref="MaxSpanInfo"/> object containing the maxima.</returns>
        public async Task<MaxSpanInfo> GetMaxima()
        {
            // Get Loading (instrumented)
            await SolverInterrogator.ApiLimiter.WaitAsync();
            var sw = Stopwatch.StartNew();
            IMemberLoading memberLoading;
            try
            {
                memberLoading = await ParentMember.GetLoadingAsync(Loading.Id, AnalysisType, LoadingResultType.Base);
            }
            finally
            {
                sw.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }
            ApiMetrics.RecordLoading(sw.ElapsedTicks);

            // Calculate maxima
            // Calculate maxima (Parallel)
            // Use Task.WhenAll to mask API latency (throttled by ApiLimiter)
            var tShearMajor = CalculateMaximum(memberLoading, ShearMajorValueOption);
            var tShearMinor = CalculateMaximum(memberLoading, ShearMinorValueOption);
            var tMomentMajor = CalculateMaximum(memberLoading, MomentMajorValueOption);
            var tMomentMinor = CalculateMaximum(memberLoading, MomentMinorValueOption);
            var tAxialForce = CalculateMaximum(memberLoading, AxialValueOption);
            var tTorsion = CalculateMaximum(memberLoading, TorsionValueOption);
            var tDeflectionMajor = CalculateMaximum(memberLoading, DeflectionMajorValueOption);
            var tDeflectionMinor = CalculateMaximum(memberLoading, DeflectionMinorValueOption);
            var tDisplacementMajor = CalculateMaximum(memberLoading, DisplacementMajorValueOption);
            var tDisplacementMinor = CalculateMaximum(memberLoading, DisplacementMinorValueOption);

            await Task.WhenAll(tShearMajor, tShearMinor, tMomentMajor, tMomentMinor, tAxialForce, tTorsion, tDeflectionMajor, tDeflectionMinor, tDisplacementMajor, tDisplacementMinor);

            MaxSpanInfoData shearMajor = tShearMajor.Result;
            MaxSpanInfoData shearMinor = tShearMinor.Result;
            MaxSpanInfoData momentMajor = tMomentMajor.Result;
            MaxSpanInfoData momentMinor = tMomentMinor.Result;
            MaxSpanInfoData axialForce = tAxialForce.Result;
            MaxSpanInfoData torsion = tTorsion.Result;
            MaxSpanInfoData deflectionMajor = tDeflectionMajor.Result;
            MaxSpanInfoData deflectionMinor = tDeflectionMinor.Result;
            MaxSpanInfoData displacementMajor = tDisplacementMajor.Result;
            MaxSpanInfoData displacementMinor = tDisplacementMinor.Result;

            // Instantiate and return MaxSpanInfo object
            return new MaxSpanInfo(Loading, shearMajor, shearMinor, momentMajor, momentMinor, axialForce, torsion, deflectionMajor, deflectionMinor, displacementMajor, displacementMinor);
        }

        private async Task<MaxSpanInfoData> CalculateMaximum(IMemberLoading loading, LoadingValueOptions option)
        {
            // Get Places of Interest (instrumented)
            List<IPointOfInterest> points = new();

            var swWait = Stopwatch.StartNew();
            await SolverInterrogator.ApiLimiter.WaitAsync();
            swWait.Stop();
            ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
            var swPoi = Stopwatch.StartNew();
            try
            {
                points.AddRange((await loading.GetPointsOfInterest(option, PointOfInterestType.Maximum)).ToList());
            }
            finally
            {
                swPoi.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }

            ApiMetrics.RecordPoi(swPoi.ElapsedTicks);

            swWait.Restart();
            await SolverInterrogator.ApiLimiter.WaitAsync();
            swWait.Stop();
            ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
            swPoi.Restart();
            try
            {
                points.AddRange((await loading.GetPointsOfInterest(option, PointOfInterestType.Minimum)).ToList());
            }
            finally
            {
                swPoi.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }
            ApiMetrics.RecordPoi(swPoi.ElapsedTicks);

            points = points.Where(p => p.SpanIndex == Span.Index).ToList();

            // Get Positions
            List<(int, double)> positions = new();
            foreach (IPointOfInterest point in points)
            {
                positions.Add((Span.Index, point.Position));
            }

            // Get Values in base units (instrumented)
            // Get Values in base units (instrumented)
            IEnumerable<ILoadingValue> values;
            swWait.Restart();
            await SolverInterrogator.ApiLimiter.WaitAsync();
            swWait.Stop();
            ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
            var swVal = Stopwatch.StartNew();
            ApiMetrics.IncrementActiveValueCalls();
            try
            {
                values = await loading.GetValueAsync(option, positions);
            }
            finally
            {
                ApiMetrics.DecrementActiveValueCalls();
                swVal.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }
            ApiMetrics.RecordValue(swVal.ElapsedTicks);

            double maxValue = 0;
            double maxPosition = 0;
            double minValue = 0;
            double minPosition = 0;

            // Convert units conversion factor
            double valCon = ConversionFactor(option.Type);

            if (values.Any())
            {
                var max = values.MaxBy(lv => lv.Value)!;
                maxValue = max.Value * valCon;
                maxPosition = MmToFt(max.Position);

                var min = values.MinBy(lv => lv.Value)!;
                minValue = min.Value * valCon;
                minPosition = MmToFt(min.Position);
            }

            return new MaxSpanInfoData(maxValue, maxPosition, minValue, minPosition);
        }

        /// <summary>
        /// Calculates force and displacement values at specified stations along the span.
        /// </summary>
        /// <returns>A list of <see cref="PointSpanInfo"/> containing data for each station.</returns>
        public async Task<List<PointSpanInfo>> GetStations()
        {
            // Calculate station positions
            IEnumerable<double> positions = Generate.LinearSpaced(Subdivisions, 0, Length);

            // Get Loading (instrumented)
            IMemberLoading memberLoading;
            var swWait = Stopwatch.StartNew();
            await SolverInterrogator.ApiLimiter.WaitAsync();
            swWait.Stop();
            ApiMetrics.RecordSemaphoreWait(swWait.ElapsedTicks);
            var sw = Stopwatch.StartNew();
            try
            {
                memberLoading = await ParentMember.GetLoadingAsync(Loading.Id, AnalysisType, LoadingResultType.Base);
            }
            finally
            {
                sw.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }
            ApiMetrics.RecordLoading(sw.ElapsedTicks);

            //Instantiate output list
            List<PointSpanInfo> stationData = new();

            foreach (double position in positions)
            {

                // Calculate values
                double shearMajor = await GetLoadingValues(memberLoading, ShearMajorValueOption, position);
                double shearMinor = await GetLoadingValues(memberLoading, ShearMinorValueOption, position);
                double momentMajor = await GetLoadingValues(memberLoading, MomentMajorValueOption, position);
                double momentMinor = await GetLoadingValues(memberLoading, MomentMinorValueOption, position);
                double axialForce = await GetLoadingValues(memberLoading, AxialValueOption, position);
                double torsion = await GetLoadingValues(memberLoading, TorsionValueOption, position);
                double deflectionMajor = await GetLoadingValues(memberLoading, DeflectionMajorValueOption, position);
                double deflectionMinor = await GetLoadingValues(memberLoading, DeflectionMinorValueOption, position);
                double displacementMajor = await GetLoadingValues(memberLoading, DisplacementMajorValueOption, position);
                double displacementMinor = await GetLoadingValues(memberLoading, DisplacementMinorValueOption, position);

                double positionFt = MmToFt(position); // Converting from [mm] to [ft]
                PointSpanInfo stationInfo = new(Loading, positionFt, shearMajor, shearMinor, momentMajor, momentMinor, axialForce, torsion, deflectionMajor, deflectionMinor, displacementMajor, displacementMinor);

                stationData.Add(stationInfo);
            }

            return stationData;
        }

        private async Task<double> GetLoadingValues(IMemberLoading loading, LoadingValueOptions option, double position)
        {
            // Get Values (instrumented)
            IEnumerable<ILoadingValue> values;
            await SolverInterrogator.ApiLimiter.WaitAsync();
            var sw = Stopwatch.StartNew();
            ApiMetrics.IncrementActiveValueCalls();
            try
            {
                values = await loading.GetValueAsync(option, Span.Index, position);
            }
            finally
            {
                ApiMetrics.DecrementActiveValueCalls();
                sw.Stop();
                SolverInterrogator.ApiLimiter.Release();
            }
            ApiMetrics.RecordValue(sw.ElapsedTicks);

            // Convert units
            double valCon = ConversionFactor(option.Type);

            double value = 0;
            if (values.Any())
            {
                value = values.MaxBy(lv => lv.Value)!.Value * valCon;
            }

            return value;
        }
    }
}
