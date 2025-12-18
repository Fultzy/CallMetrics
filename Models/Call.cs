using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
{
    public class Call
    {
        public int PhoneNumberID { get; set; }
        public int UserID { get; set; }


        public string CallType { get; set; }
        public int Duration { get; set; }
        public DateTime DateTime { get; set; }
        public string State { get; set; }


        public string UserName { get; set; }
        public string UserExtention { get; set; }
        public string TransferUser { get; set; }
        public string Caller { get; set; }

        public bool IsInternal()
        {
            return Caller.Length == UserExtention.Length;
        }
    }
}
