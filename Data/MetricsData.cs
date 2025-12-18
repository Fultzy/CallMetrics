using CallMetrics.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Data
{
    public static class MetricsData
    {
        public static List<Rep> Reps = new();
        public static List<Call> Calls = new();
        public static List<Ticket> Tickets = new();

        public static bool AddRep(Rep newRep)
        {
            if (newRep == null) return false;
            if (Reps.Any(r => r.Name == newRep.Name)) return false;

            Reps.Add(newRep);
            return true;
        }

        public static void AddCalls(List<Call> newCalls)
        {
            foreach (var call in newCalls)
            {
                AddCall(call);
            }
        }

        public static bool AddCall(Call newCall)
        {
            if (newCall == null) return false;
            if (Calls.Any(c => c.DateTime == newCall.DateTime 
            && c.Duration == newCall.Duration)) return false;

            if (Reps.Any(r => r.Name == newCall.UserName))
            {
                var rep = Reps.First(r => r.Name == newCall.UserName);
                UpdateRepCallMetrics(rep, newCall);
            }
            else
            {
                var rep = new Rep
                {
                    Name = newCall.UserName,
                    Extension = newCall.UserExtention?.Trim() ?? "None"
                };

                UpdateRepCallMetrics(rep, newCall);
                Reps.Add(rep);  
            }

                Calls.Add(newCall);
            return true;
        }

        public static void AddTickets(List<Ticket> newTickets)
        {
            foreach (var ticket in newTickets)
            {
                AddTicket(ticket);
            }
        }

        public static bool AddTicket(Ticket newTicket)
        {
            if (newTicket == null) return false;
            if (Tickets.Any(t => t.CallRecNumber == newTicket.CallRecNumber)) return false;

            if (Reps.Any(r => r.Name == newTicket.ClientName))
            {
                var rep = Reps.First(r => r.Name == newTicket.ClientName);
                UpdateRepTicketMetrics(rep, newTicket);
            }
            else
            {
                var rep = new Rep
                {
                    Name = newTicket.ClientName,
                    Extension = "None"
                };
                UpdateRepTicketMetrics(rep, newTicket);
                Reps.Add(rep);
            }

            Tickets.Add(newTicket);
            return true;
        }

        public static void Clear()
        {
            Reps.Clear();
            Calls.Clear();
            Tickets.Clear();
        }

        private static void UpdateRepCallMetrics(Rep rep, Call call)
        {
            rep.TotalCalls += 1;
            rep.TotalDuration += call.Duration;
            if (call.CallType.ToLower().Contains("inbound"))
            {
                rep.InboundCalls += 1;
                rep.InboundDuration += call.Duration;
            }
            else if (call.CallType.ToLower().Contains("outbound"))
            {
                rep.OutboundCalls += 1;
                rep.OutboundDuration += call.Duration;
            }
            if (call.Duration > 30)
            {
                rep.CallsOver30 += 1;
            }
            if (call.Duration > 60)
            {
                rep.CallsOver60 += 1;
            }
            if (call.DateTime.DayOfWeek == DayOfWeek.Saturday || call.DateTime.DayOfWeek == DayOfWeek.Sunday)
            {
                rep.WeekendCalls += 1;
            }
            if (call.IsInternal())
            {
                rep.InternalCalls += 1;
            }
        }

        private static void UpdateRepTicketMetrics(Rep rep, Ticket ticket)
        {
            rep.TotalTickets += 1;
            if (ticket.DateTime.DayOfWeek == DayOfWeek.Saturday || ticket.DateTime.DayOfWeek == DayOfWeek.Sunday)
            {
                rep.WeekendTickets += 1;
            }
        }
    }
}
