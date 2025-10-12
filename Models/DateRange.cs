using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    internal class DateRange
    {
        public DateTime StartDate { get; internal set; }
        public DateTime EndDate { get; internal set; }

        public override string ToString()
        {
            return $"{StartDate.ToShortDateString()} - {EndDate.ToShortDateString()}";
        }
    }
}
