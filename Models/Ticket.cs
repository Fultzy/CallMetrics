using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    public class Ticket
    {
        public string ClientName { get; internal set; } = "None";
        public DateTime DateTime { get; internal set; } = new();
        public string CallRecNumber { get; internal set; } = "0";
        public List<TicketEntry> TicketEntries { get; internal set; } = new();

        public bool IsResolved
        {
            get
            {
                if (TicketEntries == null || TicketEntries.Count == 0)
                    return false;

                return TicketEntries.Any(te => te.IsResolved);
            }
        }

        public int ResolutionTime // in seconds
        {
            get
            {
                if (!IsResolved || TicketEntries == null || TicketEntries.Count == 0)
                    return 0;

                var firstEntry = TicketEntries.OrderBy(te => te.DateTime).First();
                var lastEntry = TicketEntries.OrderBy(te => te.DateTime).Last();

                return (int)(lastEntry.DateTime - firstEntry.DateTime).TotalSeconds;
            }
        }


        public TicketEntry? PrimaryResolveEntry
        {
            get
            {
                if (!IsResolved || TicketEntries == null || TicketEntries.Count == 0)
                    return null;

                return TicketEntries
                    .Where(te => te.IsResolved)
                    .OrderBy(te => te.DateTime)
                    .FirstOrDefault();
            }
        }
    }
}
