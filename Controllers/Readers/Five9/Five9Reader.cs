using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CallMetrics.Controllers.Readers.Five9
{
    internal class Five9Reader
    {
        private Dictionary<string, int> HeaderIndices = new();
        public string[] RequiredHeaders =
        [
            "CALL TYPE",
            "AGENT NAME",
            "ANI", // callers number
            "TALK TIME",
            "DATE",
            "TIME",
            "SKILL",
        ];

        internal async Task<List<Call>> Start()
        {
            // open explorer to select file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.DefaultDirectory = Settings.DefaultReportPath;
            openFileDialog.DefaultExt = ".csv";
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string filePath = openFileDialog.FileName;
                return await Task.Run(() => Read(filePath));
            }

            return new List<Call>();
        }

        internal List<Call> Read(string filePath)
        {
            try
            {
                if (!System.IO.File.Exists(filePath))
                {
                    MessageBox.Show("File not found: " + filePath, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return new List<Call>();
                }

                var lines = System.IO.File.ReadAllLines(filePath);
                var calls = new List<Call>();

                // Validate header
                var headRes = ValidateHeader(lines[0].Split(','));
                if (headRes != "Valid Header")
                   throw new Exception(headRes);

                var dataLines = lines.Skip(1).ToList();
                calls = ProcessData(dataLines);
                
                return calls;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading Five9 report: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<Call>();
            }
        }

        public string ValidateHeader(string[] header)
        {
            var sb = new StringBuilder();
            HeaderIndices.Clear();
            foreach (var column in RequiredHeaders)
            {
                if (!header.Contains(column))
                {
                    sb.AppendLine($"Missing required column: {column}\n");
                }
                else
                {
                    HeaderIndices[column] = Array.IndexOf(header, column);
                }
            }

            if (sb.Length == 0)
            {
                return "Valid Header";
            }
            
            return sb.ToString();
        }

        internal List<Call> ProcessData(List<string> lines)
        {
            if (lines == null || lines.Count == 0)
            {
                return new List<Call>();
            }

            var calls = new List<Call>();

            foreach (var line in lines)
            {
                var columns = line.Split(',');
                var callType = columns[HeaderIndices["CALL TYPE"]];
                var agentName = columns[HeaderIndices["AGENT NAME"]];
                var ani = columns[HeaderIndices["ANI"]];
                var talkTimeStr = columns[HeaderIndices["TALK TIME"]];
                var dateStr = columns[HeaderIndices["DATE"]];
                var timeStr = columns[HeaderIndices["TIME"]];
                var skill = columns[HeaderIndices["SKILL"]];

                if (ParseDuration(talkTimeStr, out int talkTime) == -1)
                {
                    talkTime = 0;
                }

                DateTime callDateTime;
                if (!DateTime.TryParse($"{dateStr} {timeStr}", out callDateTime))
                {
                    callDateTime = DateTime.MinValue;
                }

                var call = new Call
                {
                    CallType = callType,
                    UserName = agentName,
                    Caller = ani,
                    Duration = talkTime,
                    DateTime = callDateTime,
                    UserExtention = "0",
                };

                
                calls.Add(call);

            }

            return calls;
        }

        private int ParseDuration(string durationStr, out int seconds)
        {
            // Expected format: HH:MM:SS
            seconds = 0;
            var parts = durationStr.Split(':');
            if (parts.Length != 3)
            {
                return -1;
            }
            if (int.TryParse(parts[0], out int hours) &&
                int.TryParse(parts[1], out int minutes) &&
                int.TryParse(parts[2], out int secs))
            {
                seconds = hours * 3600 + minutes * 60 + secs;
                return 0;
            }
            return -1;
        }
    }
}
