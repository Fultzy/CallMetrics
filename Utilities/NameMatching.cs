using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Utilities
{
    public static class NameMatching
    {
        public static List<string> GetSimilarNames(string targetName, List<string> nameList, int threshold = 1)
        {
            List<string> similarNames = new List<string>();
            foreach (var name in nameList)
            {
                if (AreNamesSimilar(targetName, name))
                {
                    similarNames.Add(name);
                }
            }
            return similarNames;
        }

        public static bool AreNamesSimilar(string name1, string name2)
        {
            if (string.IsNullOrWhiteSpace(name1) || string.IsNullOrWhiteSpace(name2))
                return false;

            name1 = name1.Trim().ToLower();
            name2 = name2.Trim().ToLower();

            // Exact match
            if (name1 == name2)
                return true;

            // Check if one name is a substring of the other
            if (name1.Contains(name2) || name2.Contains(name1))
                return true;

            // Check initials match
            var initials1 = GetInitials(name1);
            var initials2 = GetInitials(name2);
            if (initials1 == initials2)
                return true;

            return false;
        }

        private static object GetInitials(string name1)
        {
            var parts = name1.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder initials = new StringBuilder();
            foreach (var part in parts)
            {
                initials.Append(part[0]);
            }

            return initials.ToString();
        }
    }
}
