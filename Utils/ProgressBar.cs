using System;
using System.Diagnostics;

namespace TeklaResultsInterrogator.Utils
{
    /// <summary>
    /// Thread-safe console progress bar for parallel operations.
    /// Updates in-place on a single line with percentage, count, and elapsed time.
    /// </summary>
    public class ProgressBar : IDisposable
    {
        private readonly int _total;
        private readonly int _barWidth;
        private readonly Stopwatch _stopwatch;
        private readonly object _lock = new();
        private int _current;
        private bool _disposed;
        private bool _completed;

        /// <summary>
        /// Initializes a new progress bar.
        /// </summary>
        /// <param name="total">Total number of items to process.</param>
        /// <param name="barWidth">Width of the progress bar in characters (default 30).</param>
        public ProgressBar(int total, int barWidth = 30)
        {
            _total = total;
            _barWidth = barWidth;
            _current = 0;
            _stopwatch = Stopwatch.StartNew();

            // Initial render
            Render();
        }

        /// <summary>
        /// Thread-safe increment of progress. Call after each item is processed.
        /// </summary>
        public void Increment()
        {
            lock (_lock)
            {
                _current++;
                Render();

                // Auto-complete when reaching 100%
                if (_current >= _total && !_completed)
                {
                    _stopwatch.Stop();
                    _completed = true;
                    Console.WriteLine(); // Move to next line
                }
            }
        }

        /// <summary>
        /// Renders the progress bar to the console.
        /// </summary>
        private void Render()
        {
            double percent = _total > 0 ? (double)_current / _total : 0;
            int filledWidth = (int)(percent * _barWidth);
            int emptyWidth = _barWidth - filledWidth;

            string bar = new string('█', filledWidth) + new string('░', emptyWidth);
            string elapsed = _stopwatch.Elapsed.ToString(@"mm\:ss");

            // Format: [████████░░░░░░░░░░░░░░░░░░░░░░] 42.5% (85/200) 01:23
            Console.Write($"\r[{bar}] {percent:P1} ({_current}/{_total}) {elapsed}   ");
        }

        /// <summary>
        /// Completes the progress bar, moves to next line.
        /// </summary>
        public void Complete()
        {
            lock (_lock)
            {
                if (_completed) return; // Already completed

                _stopwatch.Stop();
                _current = _total;
                Render();
                _completed = true;
                Console.WriteLine(); // Move to next line
            }
        }

        /// <summary>
        /// Disposes the progress bar, completing it if not already done.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                Complete();
                _disposed = true;
            }
        }
    }
}
