using CallMetrics.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CallMetrics.Controllers
{
    public class NextivaReportReader
    {
        private List<RepData> Reps = new();
        private Dictionary<string,int> Headers = new();

        public List<RepData> Read(string filePath)
        {
            try
            {
                Reps = new List<RepData>();
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
            finally
            {
                // idk man. just felt like it i guess.
            }

            return Reps;
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


//Call Type	Transfer User	User Name	Time	Duration	Direction	Answered	State	From	To	Internal	External
        public void ParseCallRow(string[] columns)
        {
            CallData call = new();

            call.CallType = columns[Headers["Call Type"]];

            if (call.CallType == "Inbound call")
            {
                call.UserName = columns[Headers["User Name"]];
                call.UserExtention = columns[Headers["To"]];
                call.Time = DateTime.Parse(columns[Headers["Time"]]);
                call.Duration = int.Parse(columns[Headers["Duration"]]);
                call.State = columns[Headers["State"]];
                call.Caller = columns[Headers["From"]];

                // find or create rep
                RepData rep = Reps.FirstOrDefault(r => r.Name == call.UserName);
                if (rep == null)
                {
                    rep = new RepData
                    {
                        id = Reps.Count + 1,
                        Name = call.UserName,
                        Extention = call.UserExtention
                    };
                    Reps.Add(rep);
                }

                rep.AddCall(call);
            }
            else if (call.CallType == "Outbound call")
            {
                call.UserName = columns[Headers["User Name"]];
                call.UserExtention = columns[Headers["From"]];
                call.Time = DateTime.Parse(columns[Headers["Time"]]);
                call.Duration = int.Parse(columns[Headers["Duration"]]);
                call.State = columns[Headers["State"]];
                call.Caller = columns[Headers["To"]];

                // find or create rep
                RepData rep = Reps.FirstOrDefault(r => r.Name == call.UserName);
                if (rep == null)
                {
                    rep = new RepData
                    {
                        id = Reps.Count + 1,
                        Name = call.UserName,
                        Extention = call.UserExtention
                    };
                    Reps.Add(rep);
                }

                rep.AddCall(call);
            }
        }
    }
}
