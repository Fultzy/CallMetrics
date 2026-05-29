using CallMetrics.Menus;
using CallMetrics.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace CallMetrics.Utilities
{
    [Serializable]
    public static class Settings
    {
        public static SettingsData Data;

        // paths for portable (working directory) and persistent (AppData) settings
        private static string BackupPath = Path.Combine(Directory.GetCurrentDirectory(), "Settings.json");
        private static string AppDataSettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CallMetrics", "Settings.json");

        static Settings()
        {
            Data = new SettingsData();
        }

        public static void Load()
        {
            try
            {
                var hasPortable = File.Exists(BackupPath);
                var hasAppData = File.Exists(AppDataSettingsPath);

                // If both exist, check for conflicts
                if (hasPortable && hasAppData)
                {
                    var portableText = File.ReadAllText(BackupPath);
                    var appDataText = File.ReadAllText(AppDataSettingsPath);

                    var backupSettings = JsonConvert.DeserializeObject<SettingsData>(portableText);
                    var appDataSettings = JsonConvert.DeserializeObject<SettingsData>(appDataText);

                    // if identical, just load one
                    if (string.Equals(portableText, appDataText, StringComparison.Ordinal))
                    {
                        Data = backupSettings;
                        return;
                    }

                    var prompt = new SettingsConflictWindow(backupSettings, appDataSettings);
                    
                    try
                    {
                        prompt.ShowDialog();
                    }
                    catch
                    {
                        // close the program
                        App.Current.Shutdown();
                    }
                    

                    if (prompt.ApplyNewSettings)
                    {
                        if (backupSettings != null) Data = backupSettings;

                        Backup(AppDataSettingsPath); // backup and replace other
                        File.Copy(BackupPath, AppDataSettingsPath, true);
                        return;
                    }
                    else if (!prompt.ApplyNewSettings)
                    {
                        if (appDataSettings != null) Data = appDataSettings;
                        
                        Backup(BackupPath); // backup and replace other
                        File.Copy(AppDataSettingsPath, BackupPath, true);
                        return;
                    }

                    Data = backupSettings;
                    return;
                }

                // If only one exists, load it and ensure the other is populated
                if (hasAppData)
                {
                    var json = File.ReadAllText(AppDataSettingsPath);
                    var appDataSettings = JsonConvert.DeserializeObject<SettingsData>(json);
                    
                    if (appDataSettings != null) Data = appDataSettings;

                    try
                    {
                        File.Copy(AppDataSettingsPath, BackupPath, true);
                    }
                    catch { }

                    return;
                }

                if (hasPortable)
                {
                    var json = File.ReadAllText(BackupPath);
                    var backupSettings = JsonConvert.DeserializeObject<SettingsData>(json);
                    
                    if (backupSettings != null) Data = backupSettings;

                    try
                    {
                        File.Copy(BackupPath, AppDataSettingsPath, true);
                    }
                    catch { }

                    return;
                }

                // no settings file, use defaults
                Data.Teams = new List<Team>();
                Data.RankedRepsCount = 10;
                Data.AutoOpenReport = false;
                Data.DefaultReportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                Save();
            }
            catch (Exception ex)
            {
                var msg = Logger.ExceptionLog("Error loading settings. Using defaults.\n" + ex.Message);
                MessageBox.Show(msg, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static void Backup(string filepath)
        {
            var backup = filepath + ".bak." + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            try { File.Copy(filepath, backup, true); } catch { }
        }

        public static void Reset()
        {
            Data = new SettingsData();
            Save();
        }

        public static bool Save()
        {
            var backupDir = System.IO.Path.GetDirectoryName(BackupPath);
            var portableDir = System.IO.Path.GetDirectoryName(AppDataSettingsPath);
            if (!System.IO.Directory.Exists(backupDir))
            {
                System.IO.Directory.CreateDirectory(backupDir);
            }

            if (!System.IO.Directory.Exists(portableDir))
            {
                System.IO.Directory.CreateDirectory(portableDir);
            }

            try
            {
                // BUGFIX: force inbound and outbound types as lowercase
                Data.InboundCallTypes = Data.InboundCallTypes.Select(t => t.ToLower()).ToList();
                Data.OutboundCallTypes = Data.OutboundCallTypes.Select(t => t.ToLower()).ToList();

                Data.Version = App.Version; // update version
                Data.LastSave = DateTime.Now;

                var json = JsonConvert.SerializeObject(Data, Formatting.Indented);

                System.IO.File.WriteAllText(BackupPath, json);
                System.IO.File.WriteAllText(AppDataSettingsPath, json);

                return true;
            }
            catch (Exception ex)
            {
                var msg = Logger.ExceptionLog("Error saving settings: " + ex.Message);
                MessageBox.Show(msg, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
    }
}

