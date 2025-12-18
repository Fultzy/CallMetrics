using CallMetrics.Controllers.Generators.WorkSheets;
using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Controllers.Generators
{
    public class MetricsReport
    {
        private Application excelApp;
        private Workbook workbook;
        private Worksheet worksheet;

        public event EventHandler<int> ReportProgressChanged;

        public void Generate(List<Rep> reps, string directoryPath)
        {
            try
            {
                var generator = new SupportRepMetrics();
                generator.ProgressChanged += (s, e) => ReportProgressChanged?.Invoke(this, e);

                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // setup and create file
                string fileName = @"\SupportMetrics_" + UniqueTimeCode();
                workbook = excelApp.Workbooks.Add(Type.Missing);

                // create support metrics report worksheet
                worksheet = (Worksheet)workbook.ActiveSheet;
                worksheet = generator.Create(reps, worksheet);

                // save the workbook
                workbook.SaveAs(directoryPath + fileName + ".xlsx");
                workbook.Close(false);
                excelApp.Quit();

                ReportProgressChanged.Invoke(this, 100);

                if (Settings.AutoOpenReport)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = directoryPath + fileName + ".xlsx",
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error generating report: " + ex.Message);
            }
            finally
            {
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

            }            
        }

        private string UniqueTimeCode()
        {
            return DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }
    }
}
