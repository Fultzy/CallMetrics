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
        public static List<Team> Teams = new();


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
                    DefaultReportPath = DefaultReportPath
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
        public List<Team> Teams;

        public SettingsData()
        {
            DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            RankedRepsCount = 10;
            AutoOpenReport = false;
            Teams = new List<Team>();
        }
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


}

