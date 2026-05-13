


using CallMetrics.Utilities;

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

        public int InboundCallsOver30 { get; set; } = 0;
        public int InboundCallsOver60 { get; set; } = 0;
        public int OutboundCallsOver30 { get; set; } = 0;
        public int OutboundCallsOver60 { get; set; } = 0;

        public int WeekendCalls { get; set; } = 0;
        public int InternalCalls { get; set; } = 0;

        public int AverageCallTime()
        {
            if (TotalCalls == 0) return 0;

            double averageTime = (double)TotalDuration / (double)TotalCalls;
            int averageTimeInSeconds = Convert.ToInt32(averageTime);

            return averageTimeInSeconds;
        }

        public int AdjustedCalls()
        {
            return TotalCalls - InternalCalls - WeekendCalls;
        }

        public int AdjustedTickets()
        {
            return TotalTickets - WeekendTickets;
        }

        public float AverageDuration()
        {
            if (TotalCalls == 0) return 0;
            return TotalDuration / TotalCalls;
        }

        public double Over30Percentage()
        {
            return Mather.DoubleAverage(CallsOver30, TotalCalls);
        }

        public double Over60Percentage()
        {
            return Mather.DoubleAverage(CallsOver60, TotalCalls);
        }

        internal decimal CallsToTicketsRatio()
        {
            return Mather.Ratio(TotalCalls, TotalTickets);
        }

        internal double InboundOver30Percentage()
        {
            return Mather.DoubleAverage(InboundCallsOver30, InboundCalls);
        }

        internal double InboundOver60Percentage()
        {
            return Mather.DoubleAverage(InboundCallsOver60, InboundCalls);
        }

        internal double OutboundOver30Percentage()
        {
            return Mather.DoubleAverage(OutboundCallsOver30, OutboundCalls);
        }

        internal double OutboundOver60Percentage()
        {
            return Mather.DoubleAverage(OutboundCallsOver60, OutboundCalls);
        }

        internal float AverageInboundDuration()
        {
            return Mather.IntAverage(InboundDuration, InboundCalls);
        }

        internal float AverageOutboundDuration()
        {
            return Mather.IntAverage(OutboundDuration, OutboundCalls);
        }
    }
}