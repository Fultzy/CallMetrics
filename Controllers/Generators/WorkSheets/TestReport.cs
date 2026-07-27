using CallMetrics.Controllers.Generators.WorkSheets.RepMetrics;
using CallMetrics.Data;
using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Controllers.Generators.WorkSheets
{
    internal class TestReport
    {
        public Worksheet Create(List<Rep> importReps, Worksheet worksheet)
        {
            /////////////////////////////////////////// WorkSheet Setup
            // setup worksheet
            worksheet.Name = "Test Report";
            Settings.Load();

           

            Console.WriteLine(" Done!");
            return worksheet;
        }
    }
}
