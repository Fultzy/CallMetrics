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

        internal int WorkDays()
        {
            // exclude weekends and known holidays
            int workDays = 0;
            for (DateTime date = StartDate; date <= EndDate; date = date.AddDays(1))
            {
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    workDays++;
                }
            }

            return workDays;
        }
    }
}
