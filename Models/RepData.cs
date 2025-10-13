
namespace CallMetrics.Models
{
    
    public class RepData
    {
        public List<CallData> Calls = new();

        public int id = 0;
        public string Name { get; set; } = "None";
        public string Extention { get; set; } = "None";

        public int TotalCalls { get; set; } = 0;
        public int TotalDuration { get; set; } = 0;

        public int InboundCalls { get; set; } = 0;
        public int InboundDuration { get; set; } = 0;

        public int OutboundCalls { get; set; } = 0;
        public int OutboundDuration { get; set; } = 0;

        public int CallsOver30 { get; set; } = 0;
        public int CallsOver60 { get; set; } = 0;

        public int WeekendCalls { get; set; } = 0;
        public int InternalCalls { get; set; } = 0;


        public void AddCalls(List<CallData> newCalls)
        {
            foreach (CallData call in newCalls)
            {
                AddCall(call);
            }
        }

        public void AddCall(CallData newCall)
        {
            TotalCalls++;
            TotalDuration += newCall.Duration;

            if (newCall.Duration > 1800) CallsOver30++;
            if (newCall.Duration > 3600) CallsOver60++;

            if (newCall.IsWeekend() && newCall.IsInternal() == false) WeekendCalls++;
            if (newCall.IsInternal()) InternalCalls++;


            if (newCall.CallType == "Inbound call")
            {
                InboundCalls++;
                InboundDuration += newCall.Duration;
            }
            else
            {
                OutboundCalls++;
                OutboundDuration += newCall.Duration;
            }

            Calls.Add(newCall);
        }

        public string LastInitial()
        {
            if (Name == null) return "Invalid Name";
            if (Name == "-- AVERAGE --") return Name;
            if (Name == "-- TOTAL --") return Name;

            string[] name = Name.Split(' ');
            if (name.Length > 1)
            {
                return name[0] + " " + name[1][0];
            }
            else
            {
                return name[0];
            }
        }

        public string AverageCallTime()
        {
            if (TotalCalls == 0) return "0s";

            double averageTime = (double)TotalDuration / (double)TotalCalls;
            if (averageTime < 1) Console.WriteLine("OPPS!");
            int averageTimeInSeconds = Convert.ToInt32(averageTime);

            return FormatedDuration(averageTimeInSeconds);
        }

        public int AdjustedCalls()
        {
            return TotalCalls - InternalCalls - WeekendCalls;
        }

        public float AverageDuration()
        {
            if (TotalCalls == 0) return 0; // prevent divide by zero error
            return TotalDuration / TotalCalls;
        }

        public string Over30Percentage()
        {
            return ((float)CallsOver30 / (float)TotalCalls) * 100 + "%";
        }

        public float Over30PercentFloat()
        {
            if (TotalCalls == 0) return 0; // prevent divide by zero error
            return ((float)CallsOver30 / (float)TotalCalls) * 100;
        }

        public string Over60Percentage()
        {
            return ((float)CallsOver60 / (float)TotalCalls) * 100 + "%";
        }

        public float Over60PercentFloat()
        {
            if (TotalCalls == 0) return 0; // prevent divide by zero error
            return ((float)CallsOver60 / (float)TotalCalls) * 100;
        }

        public (double, double, double) OverSixtyMinutesCallsPercentages()
        {
            int overHourCalls = 0;
            int overHourInboundCalls = 0;
            int overHourOutboundCalls = 0;

            foreach (CallData call in Calls)
            {
                if (call.Duration > 3600)
                {
                    overHourCalls++;
                    if (call.CallType == "Inbound call")
                    {
                        overHourInboundCalls++;
                    }
                    else
                    {
                        overHourOutboundCalls++;
                    }
                }
            }
            return (
                Math.Round((double)overHourCalls / (double)TotalCalls * 100, 2),
                Math.Round((double)overHourInboundCalls / (double)InboundCalls * 100, 2),
                Math.Round((double)overHourOutboundCalls / (double)OutboundCalls * 100, 2)
                );
        }

        public (double, double, double) OverThirtyMinutesCallsPercentages()
        {
            int overThirtyMinuteCalls = 0;
            int overThirtyMinuteInboundCalls = 0;
            int overThirtyMinuteOutboundCalls = 0;

            foreach (CallData call in Calls)
            {
                if (call.Duration > 1800)
                {
                    overThirtyMinuteCalls++;
                    if (call.CallType == "Inbound call")
                    {
                        overThirtyMinuteInboundCalls++;
                    }
                    else
                    {
                        overThirtyMinuteOutboundCalls++;
                    }
                }
            }
            return (
                Math.Round((double)overThirtyMinuteCalls / (double)TotalCalls * 100, 2),
                Math.Round((double)overThirtyMinuteInboundCalls / (double)InboundCalls * 100, 2),
                Math.Round((double)overThirtyMinuteOutboundCalls / (double)OutboundCalls * 100, 2)
                );
        }

        public string FormatedDuration(int duration)
        {
            TimeSpan time = TimeSpan.FromSeconds(duration);
            if (time.Hours == 0 && time.Days == 0)
                return $"{time.Minutes}m {time.Seconds}s";
            else
                return $"{time.Hours + (time.Days * 24)}h {time.Minutes}m {time.Seconds}s";
        }
    }
}