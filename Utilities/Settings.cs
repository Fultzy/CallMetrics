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

        public static List<string> IgnoreTeamMetrics = new(); // team names to ignore
        public static Dictionary<string,List<string>> Teams = new(); // team name, list of rep names


        public static void Load()
        {
            if (System.IO.File.Exists(Path))
            {
                try
                {
                    var json = System.IO.File.ReadAllText(Path);
                    var data = JsonConvert.DeserializeObject<SettingsData>(json);
                    if (data.IgnoreTeamMetrics.Count > 0 || data.Teams.Count > 0)
                    {
                        IgnoreTeamMetrics = data.IgnoreTeamMetrics ?? new List<string>();
                        Teams = data.Teams ?? new Dictionary<string, List<string>>();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading settings. Using defaults.\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
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
                    IgnoreTeamMetrics = IgnoreTeamMetrics,
                    Teams = Teams,
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
}

[Serializable]
public struct SettingsData
{
    public List<string> IgnoreTeamMetrics;
    public Dictionary<string, List<string>> Teams;
}
