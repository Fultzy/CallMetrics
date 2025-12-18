using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    public class TicketEntry
    {
        public int CallRecNumber { get; internal set; }
        public DateTime DateTime { get; internal set; }
        public string Status { get; internal set; }
        public string Description { get; internal set; }
        public string AssignedByName { get; internal set; }
        public string AssignedToName { get; internal set; }
    }
}
