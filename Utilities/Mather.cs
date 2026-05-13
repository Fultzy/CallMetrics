using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Utilities
{
    public static class Mather
    {
        public static decimal Ratio(decimal value1, decimal value2)
        {
            if (value2 == 0) return 0;
            return Math.Round((decimal)value1 / value2, 2, MidpointRounding.AwayFromZero);
        }

        public static double DoubleAverage(double value1, double value2)
        {
            if (value2 == 0) return 0;
            return (double)value1 / value2 * 100;
        }

        public static int IntAverage(int value1, int value2)
        {
            if (value2 == 0) return 0;
            return (int)value1 / value2 * 100;
        }
    }
}