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
        public static Rep CreateTotalUser(List<Rep> reps)
        {
            Rep totalUser = new Rep();

            // add all the call records to the total rep
            totalUser.Name = "-- TOTAL --";

            // get the average metrics
            int totalTickets = 0;
            int weekendTickets = 0;

            int totalCalls = 0;
            int inboundCalls = 0;
            int outboundCalls = 0;

            int totalPhoneTime = 0;
            int inboundPhoneTime = 0;
            int outboundPhoneTime = 0;

            int callsOver30 = 0;
            int callsOver60 = 0;

            int weekendCalls = 0;
            int internalCalls = 0;

            int inboundOver30 = 0;
            int inboundOver60 = 0;
            int outboundOver30 = 0;
            int outboundOver60 = 0;

            foreach (var rep in reps)
            {
                if (rep.Name == "-- AVERAGE --") 
                    continue; // skip the average rep
                if (rep.Name == "-- TOTAL --") 
                    continue; // skip the total rep

                // add total tickets
                totalTickets += rep.TotalTickets;
                weekendTickets += rep.WeekendTickets;

                // add total calls
                totalCalls += rep.TotalCalls;
                inboundCalls += rep.InboundCalls;
                outboundCalls += rep.OutboundCalls;

                // add timing
                totalPhoneTime += rep.TotalDuration;
                inboundPhoneTime += rep.InboundDuration;
                outboundPhoneTime += rep.OutboundDuration;

                // add calls over 30 and 60 minutes
                callsOver30 += rep.CallsOver30;
                callsOver60 += rep.CallsOver60;

                internalCalls += rep.InternalCalls;
                weekendCalls += rep.WeekendCalls;

                inboundOver30 += rep.InboundCallsOver30;
                inboundOver60 += rep.InboundCallsOver60;
                outboundOver30 += rep.OutboundCallsOver30;
                outboundOver60 += rep.OutboundCallsOver60;
            }

            // set the Total rep metrics
            totalUser.TotalTickets = totalTickets;
            totalUser.WeekendTickets = weekendTickets;

            totalUser.TotalCalls = totalCalls;
            totalUser.InboundCalls = inboundCalls;
            totalUser.OutboundCalls = outboundCalls;

            totalUser.TotalDuration = totalPhoneTime;
            totalUser.InboundDuration = inboundPhoneTime;
            totalUser.OutboundDuration = outboundPhoneTime;

            totalUser.WeekendCalls = weekendCalls;
            totalUser.InternalCalls = internalCalls;

            totalUser.CallsOver30 = callsOver30;
            totalUser.CallsOver60 = callsOver60;

            totalUser.InboundCallsOver30 = inboundOver30;
            totalUser.InboundCallsOver60 = inboundOver60;
            totalUser.OutboundCallsOver30 = outboundOver30;
            totalUser.OutboundCallsOver60 = outboundOver60;


            return totalUser;
        }


        public static Rep CreateAverageUser(List<Rep> reps)
        {
            // create new rep
            Rep avgUser = new Rep();
            avgUser.Name = "-- AVERAGE --";

            // get the average metrics
            int userCount = reps.Count;

            if (userCount == 0)
            {
                Console.WriteLine("No genReps found");
                return avgUser;
            }

            int totalTickets = 0;
            int weekendTickets = 0;

            int totalCalls = 0;
            int inboundCalls = 0;
            int outboundCalls = 0;

            int totalPhoneTime = 0;
            int inboundPhoneTime = 0;
            int outboundPhoneTime = 0;

            int callsOver30 = 0;
            int callsOver60 = 0;

            int weekendCalls = 0;
            int internalCalls = 0;

            int inboundOver30 = 0;
            int inboundOver60 = 0;
            int outboundOver30 = 0;
            int outboundOver60 = 0;

            // add from all genReps
            foreach (var user in reps)
            {
                if (user.Name == "-- TOTAL --") continue; // skip the total rep
                if (user.Name == "-- AVERAGE --") continue; // skip the average rep

                // add total tickets
                totalTickets += user.TotalTickets;
                weekendTickets += user.WeekendTickets;

                // add total calls
                totalCalls += user.TotalCalls;
                weekendCalls += user.WeekendCalls;
                internalCalls += user.InternalCalls;
                inboundCalls += user.InboundCalls;
                outboundCalls += user.OutboundCalls;

                // add timing
                totalPhoneTime += user.TotalDuration;
                inboundPhoneTime += user.InboundDuration;
                outboundPhoneTime += user.OutboundDuration;

                // add calls over 30 and 60 minutes
                callsOver30 += user.CallsOver30;
                callsOver60 += user.CallsOver60;

                inboundOver30 += user.InboundCallsOver30;
                inboundOver60 += user.InboundCallsOver60;
                outboundOver30 += user.OutboundCallsOver30;
                outboundOver60 += user.OutboundCallsOver60;
            }

            // set the Total rep metrics
            avgUser.TotalTickets = totalTickets / userCount;
            avgUser.WeekendTickets = weekendTickets / userCount;

            avgUser.TotalCalls = totalCalls / userCount;
            avgUser.InboundCalls = inboundCalls / userCount;
            avgUser.OutboundCalls = outboundCalls / userCount;

            avgUser.TotalDuration = totalPhoneTime / userCount;
            avgUser.InboundDuration = inboundPhoneTime / userCount;
            avgUser.OutboundDuration = outboundPhoneTime / userCount;

            avgUser.WeekendCalls = weekendCalls / userCount;
            avgUser.InternalCalls = internalCalls / userCount;

            avgUser.CallsOver30 = callsOver30 / userCount;
            avgUser.CallsOver60 = callsOver60 / userCount;

            avgUser.InboundCallsOver30 = inboundOver30 / userCount;
            avgUser.InboundCallsOver60 = inboundOver60 / userCount;
            avgUser.OutboundCallsOver30 = outboundOver30 / userCount;
            avgUser.OutboundCallsOver60 = outboundOver60 / userCount;

            return avgUser;
        }

        public static double CalculateProgressSteps(int departmentCount, int repCount, int rankCount, int teamCount)
        {
            int steps = 0;

            if (departmentCount > 0)
            {
                steps += departmentCount + 1; // department rows plus the total row
            }

            steps += repCount + 2;            // user rows plus the total/average rows
            steps += rankCount * 8;           // eight ranking loops that each emit progress
            steps += teamCount;               // one progress notification per team table

            return steps == 0 ? 0 : 100.0 / steps;
        }

        public static string numToLetter(int num)
        {
            string letter = "";
            switch (num)
            {
                case 1:
                    letter = "A";
                    break;
                case 2:
                    letter = "B";
                    break;
                case 3:
                    letter = "C";
                    break;
                case 4:
                    letter = "D";
                    break;
                case 5:
                    letter = "E";
                    break;
                case 6:
                    letter = "F";
                    break;
                case 7:
                    letter = "G";
                    break;
                case 8:
                    letter = "H";
                    break;
                case 9:
                    letter = "I";
                    break;
                case 10:
                    letter = "J";
                    break;
                case 11:
                    letter = "K";
                    break;
                case 12:
                    letter = "L";
                    break;
                case 13:
                    letter = "M";
                    break;
                case 14:
                    letter = "N";
                    break;
                case 15:
                    letter = "O";
                    break;
                case 16:
                    letter = "P";
                    break;
                case 17:
                    letter = "Q";
                    break;
                case 18:
                    letter = "R";
                    break;
                case 19:
                    letter = "S";
                    break;
                case 20:
                    letter = "T";
                    break;
                case 21:
                    letter = "U";
                    break;
                case 22:
                    letter = "V";
                    break;
                case 23:
                    letter = "W";
                    break;
                case 24:
                    letter = "X";
                    break;
                case 25:
                    letter = "Y";
                    break;
                case 26:
                    letter = "Z";
                    break;
                default:
                    letter = "A";
                    break;
            }

            return letter;
        }

    }
}
