using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;

namespace ExcelWriter
{
#if DEBUG

    /// <summary>
    /// DEBUG CLASS - Provides Console.WriteLines IN DEBUG BUILD ONLY
    /// </summary>
    internal static class TempDiagnostics
    {
        private static DateTime resetTime = DateTime.MinValue;
        private static DateTime outputTime = DateTime.MinValue;

        /// <summary>
        /// DEBUG METHOD - Output a message to the console.
        /// </summary>
        public static void Output(string message)
        {
            DateTime currentTime = DateTime.Now;

            if (resetTime == DateTime.MinValue)
            {
                // This is first call - reset everything
                resetTime = currentTime;
                outputTime = currentTime;
            }

            // Calc time since last reset + last output
            TimeSpan overallTimeSpan = currentTime - resetTime;
            TimeSpan outputTimeSpan = currentTime - outputTime;

            // Update last output time
            outputTime = currentTime;

            string newMessage = string.Format("Op={0:mm'm'ss's'ffff'ms'},Elapsed={1:mm'm'ss's'ffff'ms'} - {2}", outputTimeSpan, overallTimeSpan, message);
            Console.WriteLine(newMessage);
        }

        /// <summary>
        /// DEBUG METHOD - Output a message to the console, resets a timestamp
        /// </summary>
        public static void Output(string message, bool reset)
        {
            if (reset) resetTime = DateTime.MinValue;
            Output(message);
        }
    }



#else
    /// <summary>
    /// DEBUG CLASS - Does nothing in Non-Debug mode.
    /// </summary>
    internal static class TempDiagnostics
    {
        /// <summary>
        /// DEBUG METHOD - Does nothing in Non-Debug mode.
        /// </summary>
        public static void Output(string message)
        {
            return;
        }

        /// <summary>
        /// DEBUG METHOD - Does nothing in Non-Debug mode.
        /// </summary>
        public static void Output(string message, bool reset)
        {
        }    
    }
#endif
}
