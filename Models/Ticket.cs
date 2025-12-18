using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    public class Ticket
    {
        public string ClientName { get; internal set; }
        public DateTime DateTime { get; internal set; }
        public string CallRecNumber { get; internal set; }
        public List<TicketEntry> TicketEntries { get; internal set; }
    }
}
