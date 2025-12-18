


namespace CallMetrics.Models
{
    
    public class Rep
    {
        public string Name { get; set; } = "None";
        public string Extension { get; set; } = "None";

        // Tickets
        public int TotalTickets { get; set; } = 0;
        public int WeekendTickets { get; set; } = 0;
        


        // Calls
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

        public string LastInitial()
        {
            if (Name == null) return "Invalid Name";
            if (Name == "-- AVERAGE --") return Name;
            if (Name == "-- TOTAL --") return Name;

            string[] name = Name.Split(' ');
            if (name.Length > 1)
            {
                return name[0] + " " + name.Last()[0];
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

            return FormattedDuration(averageTimeInSeconds);
        }

        public int AdjustedCalls()
        {
            return TotalCalls - InternalCalls - WeekendCalls;
        }

        public float AverageDuration()
        {
            if (TotalCalls == 0) return 0; 
            return TotalDuration / TotalCalls;
        }

        public string Over30Percentage()
        {
            if (TotalCalls == 0) return "Divided by 0!";
            return ((float)CallsOver30 / (float)TotalCalls) * 100 + "%";
        }

        public float Over30PercentFloat()
        {
            if (TotalCalls == 0) return 0;
            return ((float)CallsOver30 / (float)TotalCalls) * 100;
        }

        public string Over60Percentage()
        {
            if (TotalCalls == 0) return "Divided by 0!";
            return ((float)CallsOver60 / (float)TotalCalls) * 100 + "%";
        }

        public float Over60PercentFloat()
        {
            if (TotalCalls == 0) return 0;
            return ((float)CallsOver60 / (float)TotalCalls) * 100;
        }

        public string FormattedDuration(int duration)
        {
            TimeSpan time = TimeSpan.FromSeconds(duration);
            if (time.Hours == 0 && time.Days == 0)
                return $"{time.Minutes}m {time.Seconds}s";
            else
                return $"{time.Hours + (time.Days * 24)}h {time.Minutes}m {time.Seconds}s";
        }
    }
}