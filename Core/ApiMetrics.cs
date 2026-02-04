using System.Threading;

namespace TeklaResultsInterrogator.Core
{
    /// <summary>
    /// Tracks TSD API usage statistics (Latency, Concurrency, Throttling).
    /// Shared across all commands to provide consistent diagnostics.
    /// </summary>
    public static class ApiMetrics
    {
        private static long _loadingCalls = 0;
        private static long _loadingDuration = 0;
        private static long _valueCalls = 0;
        private static long _valueDuration = 0;
        private static long _poiCalls = 0;
        private static long _poiDuration = 0;
        private static int _activeValueCalls = 0;
        private static int _maxConcurrency = 0;
        private static long _semaphoreWaitCalls = 0;
        private static long _semaphoreWaitDuration = 0;
        private static long _semaphoreWaitMax = 0;

        /// <summary>Total number of GetLoadingAsync calls.</summary>
        public static long LoadingCalls => Interlocked.Read(ref _loadingCalls);
        /// <summary>Total duration of GetLoadingAsync calls in ticks.</summary>
        public static long LoadingDuration => Interlocked.Read(ref _loadingDuration);
        /// <summary>Total number of GetValueAsync calls.</summary>
        public static long ValueCalls => Interlocked.Read(ref _valueCalls);
        /// <summary>Total duration of GetValueAsync calls in ticks.</summary>
        public static long ValueDuration => Interlocked.Read(ref _valueDuration);
        /// <summary>Total number of PointOfInterest calls.</summary>
        public static long PoiCalls => Interlocked.Read(ref _poiCalls);
        /// <summary>Total duration of PointOfInterest calls in ticks.</summary>
        public static long PoiDuration => Interlocked.Read(ref _poiDuration);
        /// <summary>Count of currently active GetValueAsync calls (concurrency).</summary>
        public static int ActiveValueCalls => _activeValueCalls;
        /// <summary>Maximum observed concurrency for GetValueAsync calls.</summary>
        public static int MaxConcurrency => _maxConcurrency;
        /// <summary>Total number of times execution waited on the semaphore.</summary>
        public static long SemaphoreWaitCalls => Interlocked.Read(ref _semaphoreWaitCalls);
        /// <summary>Total duration spent waiting on the semaphore in ticks.</summary>
        public static long SemaphoreWaitDuration => Interlocked.Read(ref _semaphoreWaitDuration);
        /// <summary>Maximum duration spent waiting on the semaphore in ticks.</summary>
        public static long SemaphoreWaitMax => Interlocked.Read(ref _semaphoreWaitMax);

        /// <summary>Records a Loading call and its duration.</summary>
        public static void RecordLoading(long ticks)
        {
            Interlocked.Increment(ref _loadingCalls);
            Interlocked.Add(ref _loadingDuration, ticks);
        }

        /// <summary>Records a Value call and its duration.</summary>
        public static void RecordValue(long ticks)
        {
            Interlocked.Increment(ref _valueCalls);
            Interlocked.Add(ref _valueDuration, ticks);
        }

        /// <summary>Records a Point of Interest call and its duration.</summary>
        public static void RecordPoi(long ticks)
        {
            Interlocked.Increment(ref _poiCalls);
            Interlocked.Add(ref _poiDuration, ticks);
        }

        /// <summary>Increments the active value call count and updates peak concurrency.</summary>
        public static int IncrementActiveValueCalls()
        {
            int current = Interlocked.Increment(ref _activeValueCalls);
            RecordConcurrency(current);
            return current;
        }

        /// <summary>Decrements the active value call count.</summary>
        public static void DecrementActiveValueCalls() => Interlocked.Decrement(ref _activeValueCalls);

        /// <summary>Records a semaphore wait event and its duration.</summary>
        public static void RecordSemaphoreWait(long ticks)
        {
            Interlocked.Increment(ref _semaphoreWaitCalls);
            Interlocked.Add(ref _semaphoreWaitDuration, ticks);
            RecordWait(ticks);
        }

        /// <summary>
        /// Atomically updates MaxConcurrency if the current value is higher.
        /// </summary>
        private static void RecordConcurrency(int current)
        {
            int initial, computed;
            do
            {
                initial = _maxConcurrency;
                if (current <= initial) break;
                computed = current;
            } while (Interlocked.CompareExchange(ref _maxConcurrency, computed, initial) != initial);
        }

        /// <summary>
        /// Atomically updates SemaphoreWaitMax if the current wait is higher.
        /// </summary>
        private static void RecordWait(long waitTicks)
        {
            long initial, computed;
            do
            {
                initial = Interlocked.Read(ref _semaphoreWaitMax);
                if (waitTicks <= initial) break;
                computed = waitTicks;
            } while (Interlocked.CompareExchange(ref _semaphoreWaitMax, computed, initial) != initial);
        }

        /// <summary>
        /// Resets all metrics to zero (e.g. at start of a command).
        /// </summary>
        public static void Reset()
        {
            Interlocked.Exchange(ref _loadingCalls, 0);
            Interlocked.Exchange(ref _loadingDuration, 0);
            Interlocked.Exchange(ref _valueCalls, 0);
            Interlocked.Exchange(ref _valueDuration, 0);
            Interlocked.Exchange(ref _poiCalls, 0);
            Interlocked.Exchange(ref _poiDuration, 0);
            Interlocked.Exchange(ref _activeValueCalls, 0);
            Interlocked.Exchange(ref _maxConcurrency, 0);
            Interlocked.Exchange(ref _semaphoreWaitCalls, 0);
            Interlocked.Exchange(ref _semaphoreWaitDuration, 0);
            Interlocked.Exchange(ref _semaphoreWaitMax, 0);
        }
    }
}
