using CallMetrics.Controllers.Generators.WorkSheets;
using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CallMetrics.Controllers.Generators
{
    public class MetricsReport
    {
        private Application? excelApp = null;
        private Workbook? workbook = null;
        private Worksheet? worksheet = null;

        public event EventHandler<int> ReportProgressChanged;

        public void Generate(List<Rep> reps, string directoryPath)
        {
            try
            {
                using var generator = new SupportRepMetrics();
                generator.ProgressChanged += (s, e) => ReportProgressChanged?.Invoke(this, e);

                //CloseExcel();
                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true,
                    DisplayAlerts = false
                };

                //// setup and create file
                string fileName = @"\SupportMetrics_" + UniqueTimeCode();
                workbook = excelApp.Workbooks.Add(Type.Missing);

                //// create support metrics report worksheet
                worksheet = (Worksheet)workbook.ActiveSheet;
                worksheet = generator.Create(reps, worksheet);

                //// save the workbook
                workbook.SaveAs(directoryPath + fileName + ".xlsx");

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
            catch 
            {
                throw;
            }
            finally
            {
                CloseExcel();
            }
        }

        private void CloseExcel()
        {
            try
            {
                if (worksheet != null)
                {
                    Marshal.FinalReleaseComObject(worksheet);
                    worksheet = null;
                }

                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch { }

                    Marshal.FinalReleaseComObject(workbook);
                    workbook = null;
                }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch { }

                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                throw;
            }

        }

        private string UniqueTimeCode()
        {
            return DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }
    }
}
