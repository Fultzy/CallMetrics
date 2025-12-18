using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CallMetrics.Controllers.Readers.Nextiva
{
    public static class NextivaHelper
    {
        public static string GetName(Dictionary<string, int> headers, string[] row)
        {
            string[] options = new string[]
            {
                "User Name",
                "Name",
                "User",
                "Agent Name"
            };

            foreach (string option in options)
            {
                if (headers.ContainsKey(option))
                {
                    return row[headers[option]];
                }
            }
            throw new Exception("No valid 'Name' header found in Nextiva report.");
        }

        public static int GetDurationInSeconds(string input)
        {
            input = input.Trim().ToLower();

            // Case 1: "1h 10m 23s"
            if (input.Contains('h') || input.Contains('m') || input.Contains('s'))
            {
                int totalSeconds = 0;
                var matches = Regex.Matches(input, @"(\d+\.?\d*)\s*([hms])");

                foreach (Match match in matches)
                {
                    double value = double.Parse(match.Groups[1].Value);
                    char unit = match.Groups[2].Value[0];

                    if (unit == 'h') totalSeconds += (int)(value * 3600);
                    else if (unit == 'm') totalSeconds += (int)(value * 60);
                    else if (unit == 's') totalSeconds += (int)value;
                }

                return totalSeconds;
            }

            // Case 2:"1234.0"
            if (input.Contains('.'))
                return (int)float.Parse(input);

            // Case 3: "1234"
            return int.Parse(input);
        }

        
    }
}
