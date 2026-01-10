using CallMetrics.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Windows;
using System.IO;
using System.Drawing;


namespace CallMetrics.Utilities
{
    [Serializable]
    public static class Settings
    {
        // current running directory
        private static string Path = Directory.GetCurrentDirectory() + "/Settings.json";

        public static string DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public static int RankedRepsCount = 10;
        public static bool AutoOpenReport = false;
        public static ImportType TicketImportType = ImportType.CallTracker;
        public static ImportType CallImportType = ImportType.Nextiva;

        public static List<string> InboundCallTypes = new() { "Inbound" };
        public static List<string> OutboundCallTypes = new() { "Outbound" };

        public static List<Team> Teams = new();
        public static List<Alias> Aliases = new();


        public static void Load()
        {
            if (System.IO.File.Exists(Path))
            {
                try
                {
                    var json = System.IO.File.ReadAllText(Path);
                    var data = JsonConvert.DeserializeObject<SettingsData>(json);
                    
                    Teams = data.Teams;
                    RankedRepsCount = data.RankedRepsCount;
                    AutoOpenReport = data.AutoOpenReport;
                    DefaultReportPath = data.DefaultReportPath;
                    CallImportType = data.CallImportType;
                    
                    InboundCallTypes = data.InboundCallTypes;
                    OutboundCallTypes = data.OutboundCallTypes;
                    
                    TicketImportType = data.TicketImportType;
                    Aliases = data.Aliases;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading settings. Using defaults.\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                // no settings file, use defaults
                Teams = new List<Team>();
                RankedRepsCount = 10;
                AutoOpenReport = false;
                DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                Save();
            }
        }

        public static bool Save()
        {
            var directory = System.IO.Path.GetDirectoryName(Path);
            if (!System.IO.Directory.Exists(directory))
            {
                System.IO.Directory.CreateDirectory(directory);
            }

            try
            {
                var data = new SettingsData
                {
                    Teams = Teams,
                    RankedRepsCount = RankedRepsCount,
                    AutoOpenReport = AutoOpenReport,
                    DefaultReportPath = DefaultReportPath,
                    CallImportType = CallImportType,
                   
                    InboundCallTypes = InboundCallTypes,
                    OutboundCallTypes = OutboundCallTypes,
                    
                    TicketImportType = TicketImportType,
                    Aliases = Aliases,
                };

                var json = JsonConvert.SerializeObject(data, Formatting.Indented);
                System.IO.File.WriteAllText(Path, json);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving settings. RIP\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
    }


    [Serializable]
    public struct SettingsData
    {
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
            DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            RankedRepsCount = 10;
            AutoOpenReport = false;
            TicketImportType = ImportType.CallTracker;
            CallImportType = ImportType.Nextiva;
            
            InboundCallTypes = new List<string> { "Inbound" };
            OutboundCallTypes = new List<string> { "Outbound" };
            
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
    public struct Team
    {
        public string Name;

        public bool IncludeInMetrics;
        public bool IsDepartment;
        public bool HideTeam;

        public List<string> Members;

        public bool IsNull()
        {
            return this.Equals(default(Team));
        }
    }

    [Serializable]
    public struct Alias
    {
        public string Name;
        public List<string> AliasedTo;
        public ImportType AddedFrom;

        public bool IsNull()
        {
            var n = string.IsNullOrEmpty(Name);
            var a = AliasedTo == null || AliasedTo.Count == 0;

            return n && a;
        }

        public override string ToString()
        {
            return string.Join(", ", AliasedTo);
        }
    }


}

