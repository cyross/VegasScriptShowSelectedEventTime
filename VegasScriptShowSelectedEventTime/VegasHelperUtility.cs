using System;

namespace VegasScriptShowSelectedEventTime
{
    internal class VegasHelperUtility
    {
        public static string NanoToTimestamp(long nanos)
        {
            TimeSpan span = new TimeSpan(nanos);
            return span.ToString("g");
        }

        public static long RoundNanos(long nanos)
        {
            return nanos + 500000 / 1000000 * 1000000;
        }
    }
}
