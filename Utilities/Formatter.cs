using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Utilities
{
    public static class Formatter
    {
        public static string Duration(int duration)
        {
            TimeSpan time = TimeSpan.FromSeconds(duration);
            if (time.Hours == 0 && time.Days == 0)
                return $"{time.Minutes}m {time.Seconds}s";
            else
                return $"{time.Hours + (time.Days * 24)}h {time.Minutes}m {time.Seconds}s";
        }

        public static string LastInitial(string name)
        {
            if (name == null) return "Invalid Name";
            if (name == "-- AVERAGE --") return name;
            if (name == "-- TOTAL --") return name;

            string[] names = name.Split(' ');
            if (names.Length > 1)
            {
                return names[0] + " " + names.Last()[0];
            }
            else
            {
                return names[0];
            }
        }
    }
}
