using CallMetrics.Controllers.Generators.WorkSheets.RepMetrics;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;

namespace CallMetrics.Controllers.Generators.WorkSheets.Helpers
{
    internal class HeaderHelper
    {
        public static Worksheet WriteOneSidedHeader(Worksheet worksheet, string name, int row, Dictionary<string, int> headers, out int newRow, string prefix = "", int colOffset = 0)
        {
            try
            {
                var clSt = SupportMetricsHelper.NumToLetter(colOffset + 1);
                var clEd = SupportMetricsHelper.NumToLetter(colOffset + headers.Count);

                if (!string.IsNullOrEmpty(name))
                {
                    worksheet.Cells[row, colOffset + 1] = name;

                    var titleEd = SupportMetricsHelper.NumToLetter(colOffset + 1 + 10);
                    FormatHelper.HeaderTitle(worksheet, row, clSt, clEd, titleEd);

                    row++;
                }

                // write headers
                foreach (var header in headers)
                {
                    if (header.Value != null)
                    {
                        worksheet.Cells[row, colOffset + header.Value] = header.Key;
                    }
                }

                FormatHelper.HeaderRow(worksheet, row, clSt, clEd);
            }
            catch (Exception ex)
            {
                var msg = Logger.ExceptionLog("Error adding rep header: " + ex.Message);
                throw new Exception(msg);
            }

            newRow = row;
            newRow ++;
            return worksheet;
        }


        public static Worksheet WriteTwoSidedHeader(Worksheet worksheet, string name, int row, Dictionary<string, int> headers, out int newRow, string prefix = "", int colOffset = 0)
        {
            var clSt = SupportMetricsHelper.NumToLetter(colOffset + 1);
            var clEd = SupportMetricsHelper.NumToLetter(colOffset + 2);
            var colLast = SupportMetricsHelper.NumToLetter(headers.Count * 2 + 1);

            if (!string.IsNullOrEmpty(name))
            {
                worksheet.Cells[row, colOffset + 1] = name;

                var titleEd = SupportMetricsHelper.NumToLetter(colOffset + 1 + 10);
                FormatHelper.HeaderTitle(worksheet, row, clSt, colLast, titleEd);

                row++;
            }

            // write headers
            foreach (var header in headers)
            {
                var leftHeader = header.Key.Split(':')[0];
                var rightHeader = header.Key.Split(':').Length > 1 ? header.Key.Split(':')[1] : "";

                if (prefix == "") // used for inbound or outbound while using same header
                {
                    worksheet.Cells[row, colOffset + header.Value] = leftHeader;
                }
                else
                {
                    worksheet.Cells[row, colOffset + header.Value] = prefix + " " + leftHeader;
                }

                if (rightHeader != null)
                {
                    worksheet.Cells[row, colOffset + header.Value + 1] = rightHeader;
                }

                var cell1 = $"{SupportMetricsHelper.NumToLetter(colOffset + header.Value)}{row}";
                var cell2 = $"{SupportMetricsHelper.NumToLetter(colOffset + header.Value + 1)}{row}";

                worksheet = FormatHelper.JoinCells(worksheet, cell1, cell2);
            }

            worksheet = FormatHelper.HeaderRow(worksheet, row, clSt, colLast);

            newRow = row;
            newRow++;
            return worksheet;
        }
    }
}
