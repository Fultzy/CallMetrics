using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Controllers.Generators.WorkSheets
{
    internal class SupportMetricsHelper
    {
        public List<Rep> GetAllReps(List<Team> teams)
        {
            List<Rep> allReps = new();
            foreach (var team in teams)
            {
                allReps.AddRange((IEnumerable<Rep>)team.Members);
            }
            return allReps;

        }
    }
}
