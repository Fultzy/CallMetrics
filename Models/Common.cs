using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Models
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

    public enum ImportType
    {
        CallTracker,
        Dynamics,
        Nextiva,
        Five9,
        None,
    }


    [Serializable]
    public class Team
    {
        public string Name;

        public bool IncludeInMetrics;
        public bool IsDepartment;
        public bool HideTeam;
        public bool IsExcluded;

        public List<string> Members;

        public Team()
        {
            Members = new List<string>();
        }

        public bool IsNull()
        {
            var noName = string.IsNullOrEmpty(Name);
            var noMembers = Members == null || Members.Count == 0;
            return noName && noMembers && !IncludeInMetrics && !IsDepartment && !HideTeam && !IsExcluded;
        }
    }

    [Serializable]
    public class Alias
    {
        public string Name;
        public List<string> AliasedTo;
        public ImportType AddedFrom;

        public Alias()
        {
            AliasedTo = new List<string>();
        }

        public bool IsNull()
        {
            var n = string.IsNullOrEmpty(Name);
            var a = AliasedTo == null || AliasedTo.Count == 0;

            return n && a;
        }

        public override string ToString()
        {
            if (AliasedTo == null) return string.Empty;
            return string.Join(", ", AliasedTo);
        }

        public bool For(string name)
        {
            if (name == null) throw new ArgumentNullException("name");
            if (AliasedTo == null) return false;

            return AliasedTo.Any(n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
