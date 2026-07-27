using CallMetrics.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Data
{
    [Serializable]
    public class SettingsData
    {
        public string Version;
        public DateTime LastSave;

        public string DefaultReportPath;
        public int RankedRepsCount;
        public bool AutoOpenReport;
        public ImportType TicketImportType;
        public ImportType CallImportType;

        public List<string> InboundCallTypes;
        public List<string> OutboundCallTypes;

        public List<Team> Teams;
        public List<Alias> Aliases;

        public SettingsData()
        {
            Version = "Old";
            LastSave = DateTime.Now;

            DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            RankedRepsCount = 10;
            AutoOpenReport = false;
            TicketImportType = ImportType.CallTracker;
            CallImportType = ImportType.Five9;

            InboundCallTypes = new List<string>();
            OutboundCallTypes = new List<string>();

            Teams = new List<Team>();
            Aliases = new List<Alias>();
        }

    }
}
