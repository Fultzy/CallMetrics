using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    public class CallData
    {
        public int id = 0;
        public int PhoneNumberID { get; set; }
        public int UserID { get; set; }


        public string CallType { get; set; }
        public int Duration { get; set; }
        public DateTime Time { get; set; }
        public string State { get; set; }


        public string UserName { get; set; }
        public string UserExtention { get; set; }
        public string TransferUser { get; set; }
        public string Caller { get; set; }


        public bool IsWeekend()
        {
            if (Time.DayOfWeek == DayOfWeek.Saturday || Time.DayOfWeek == DayOfWeek.Sunday)
                return true;
            else
                return false;
        }

        public bool IsInternal()
        {
            if (Caller.Length == UserExtention.Length)
                return true;
            else
                return false;
        }
    }
}
