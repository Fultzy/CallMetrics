using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

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


        private static Alias GetAliasForRepName(string repName)
        {
            foreach (var alias in Settings.Aliases)
            {
                if (alias.AliasedTo.Any(name => name == repName))
                {
                    return alias;
                }
            }
            return new Alias();
        }

        private static Alias CreateAliasForRepName(string repName)
        {
            var newAlias = new Alias
            {
                Name = repName,
                AliasedTo = new List<string> { repName }
            };

            Settings.Aliases.Add(newAlias);
            Settings.Save();
            return newAlias;
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
                var alias = GetAliasForRepName(newCall.UserName);
                if (alias.IsNull()) // only create alias from Call Records. 
                {
                    alias = CreateAliasForRepName(newCall.UserName);
                }

                var rep = new Rep
                {
                    Name = alias.Name,
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

            var repNameFromTicket = newTicket.TicketEntries.First().AssignedToName;

            var alias = GetAliasForRepName(repNameFromTicket);

            // Try direct match first
            var rep = Reps.FirstOrDefault(r => r.Name == repNameFromTicket);
            if (rep != null)
            {
                UpdateRepTicketMetrics(rep, newTicket);
            }
            else
            {
                // Try to resolve via alias mapping
                if (!alias.IsNull())
                {
                    // Check if a rep with the alias name already exists
                    rep = Reps.FirstOrDefault(r => r.Name == alias.Name);
                    if (rep != null)
                    {
                        UpdateRepTicketMetrics(rep, newTicket);
                    }
                    else
                    {
                        // Create new rep using canonical name from alias
                        rep = new Rep
                        {
                            Name = alias.Name,
                            Extension = "None"
                        };

                        UpdateRepTicketMetrics(rep, newTicket);
                        Reps.Add(rep);
                    }
                }
                else
                {
                    rep = new Rep
                    {
                        Name = repNameFromTicket,
                        Extension = "None"
                    };

                    UpdateRepTicketMetrics(rep, newTicket);
                    Reps.Add(rep);
                }

                
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

            if (Settings.InboundCallTypes.Contains(call.CallType.ToLower()))
            {
                rep.InboundCalls += 1;
                rep.InboundDuration += call.Duration;
            }
            else if (Settings.OutboundCallTypes.Contains(call.CallType.ToLower()))
            {
                rep.OutboundCalls += 1;
                rep.OutboundDuration += call.Duration;
            }

            if (call.Duration > 30 * 60) // 30 minutes  
            {
                rep.CallsOver30 += 1;
            }

            if (call.Duration > 60 * 60) // 60 minutes
            {
                rep.CallsOver60 += 1;
            }

            if (call.DateTime.DayOfWeek == DayOfWeek.Saturday || call.DateTime.DayOfWeek == DayOfWeek.Sunday)
            {
                rep.WeekendCalls += 1;
            }

            if (call.IsInternal()) // only used for Nextiva
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
