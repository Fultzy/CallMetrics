using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CallMetrics.Controllers.Readers.Nextiva
{
    public class NextivaReader
    {
        private List<Call> Calls = new();
        private Dictionary<string,int> Headers = new();

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

        public List<Call> Read(string filePath)
        {
            try
            {
                Calls = new List<Call>(); 
                string[] lines = File.ReadAllLines(filePath);

                // first line is the header
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    string[] columns = ParseCsvLine(line);

                    if (i == 0)
                    {
                        // map headers
                        for (int c = 0; c < columns.Length; c++)
                        {
                            Headers[columns[c]] = c;
                        }
                        continue;
                    }

                    ParseCallRow(columns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading file: " + ex.Message + "\n Try closing the damn file, huh?", "idk bro, good luck.. ", MessageBoxButton.YesNoCancel, MessageBoxImage.Error);
            }

            return Calls;
        }

        private string[] ParseCsvLine(string line)
        {
            List<string> columns = new List<string>();
            StringBuilder currentColumn = new StringBuilder();
            bool inQuotes = false;
            foreach (char c in line)
            {
                if (c == '\"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    columns.Add(currentColumn.ToString());
                    currentColumn.Clear();
                }
                else
                {
                    currentColumn.Append(c);
                }
            }
            columns.Add(currentColumn.ToString());
            return columns.ToArray();
        }


//Call Type	Transfer User	Name	DateTime	Duration	Direction	Answered	State	From	To	Internal	External
        public void ParseCallRow(string[] columns)
        {
            Call call = new();

            call.CallType = columns[Headers["Call Type"]];
            call.UserName = NextivaHelper.GetName(Headers, columns);
            call.DateTime = DateTime.Parse(columns[Headers["DateTime"]]);
            call.Duration = NextivaHelper.GetDurationInSeconds(columns[Headers["Duration"]]);
            call.State = columns[Headers["State"]];

            if (call.CallType == "Inbound call")
            {
                call.UserExtention = columns[Headers["To"]];
                call.Caller = columns[Headers["From"]];
            }
            else if (call.CallType == "Outbound call")
            {
                call.UserExtention = columns[Headers["From"]];
                call.Caller = columns[Headers["To"]];
            }

            Calls.Add(call);
        }


    }
}
