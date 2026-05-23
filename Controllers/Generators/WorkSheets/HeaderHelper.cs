using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;

namespace CallMetrics.Controllers.Generators.WorkSheets
{
    internal class HeaderHelper
    {
        public static Worksheet WriteOneSidedHeader(Worksheet worksheet, int row, string title, Dictionary<string, int> headers, string prefix = "", int colOffset = 0)
        {
            try
            {
                if (!string.IsNullOrEmpty(title))
                {
                    // Table Name Cells
                    // going back to full name in one cell
                    worksheet.Cells[row, colOffset + 1] = title;

                    //var wrds = title.Split(' ');
                    //worksheet.Cells[row, colOffset + 1] = wrds[0];
                    //worksheet.Cells[row + 1, colOffset + 1] = wrds.Length > 1 ? wrds[1] : "";

                    worksheet.Cells[row + 1, colOffset + 1].Font.Bold = true;
                    worksheet.Cells[row + 1, colOffset + 1].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    worksheet.Cells[row, colOffset + 1].Font.Bold = true;
                    worksheet.Cells[row, colOffset + 1].Font.Size = 14;
                    worksheet.Cells[row, colOffset + 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;

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

                var clSt = SupportMetricsHelper.numToLetter(colOffset + 1);
                var clEd = SupportMetricsHelper.numToLetter(colOffset + headers.Count);

                // format header row
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Font.Bold = true;
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // Make the background color of the header light blue
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }
            catch (Exception ex)
            {
                var msg = Logger.ExceptionLog("Error adding rep header: " + ex.Message);
                throw new Exception(msg);
            }

            return worksheet;
        }


        public static Worksheet WriteTwoSidedHeader(Worksheet worksheet, string name, int row, Dictionary<string, int> headers, string prefix = "", int colOffset = 0)
        {
            if (!string.IsNullOrEmpty(name))
            {
                // Table Name Cells
                var wrds = name.Split(' ');
                worksheet.Cells[row, colOffset + 1] = wrds[0];
                worksheet.Cells[row + 1, colOffset + 1] = wrds.Length > 1 ? wrds[1] : "";
                worksheet.Cells[row + 1, colOffset + 1].Font.Bold = true;
                worksheet.Cells[row + 1, colOffset + 1].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[row, colOffset + 1].Font.Bold = true;
                worksheet.Cells[row, colOffset + 1].Font.Size = 14;
                worksheet.Cells[row, colOffset + 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;

                row++;
            }

            // write headers
            foreach (var header in headers)
            {
                var leftHeader = header.Key.Split(':')[0];
                var rightHeader = header.Key.Split(':').Length > 1 ? header.Key.Split(':')[1] : "";

                if (prefix == "")
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
            }

            // Center and join rank headers // do border here
            for (int i = 0; i < headers.Count; i++)
            {
                var colNum = headers.ElementAt(i).Value;

                var cl1 = SupportMetricsHelper.numToLetter(colOffset + colNum);
                var cl2 = SupportMetricsHelper.numToLetter(colOffset + colNum + 1);

                worksheet.Range[cl1 + row, cl2 + row].Merge();
                worksheet.Range[cl1 + row, cl2 + row].Font.Bold = true;
                worksheet.Range[cl1 + row, cl2 + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }


            var srtCol = SupportMetricsHelper.numToLetter(colOffset + 1);
            var endCol = SupportMetricsHelper.numToLetter(colOffset + (headers.Count) * 2);

            // format header row
            worksheet.Range[$"{srtCol}{row}", $"{endCol}{row}"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Range[$"{srtCol}{row}", $"{endCol}{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            return worksheet;
        }


        public static Worksheet WriteThreeSidedHeader(Worksheet worksheet, string name, int row, Dictionary<string, int> headers, int colOffset = 0)
        {
            if (!string.IsNullOrEmpty(name))
            {
                // Table Name Cells
                var wrds = name.Split(' ');
                worksheet.Cells[row, colOffset + 1] = wrds[0];
                worksheet.Cells[row + 1, colOffset + 1] = wrds.Length > 1 ? wrds[1] : "";
                worksheet.Cells[row + 1, colOffset + 1].Font.Bold = true;
                worksheet.Cells[row + 1, colOffset + 1].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[row, colOffset + 1].Font.Bold = true;
                worksheet.Cells[row, colOffset + 1].Font.Size = 14;
                worksheet.Cells[row, colOffset + 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;

                row++;
            }

            // write headers
            foreach (var header in headers)
            {
                var hd = header.Key.Split(':');
                var prefix = hd[0];
                var leftHeader = hd.Length > 1 ? hd[1] : "";
                var rightHeader = hd.Length > 2 ? hd[2] : "";

                if (leftHeader != "")
                {
                    worksheet.Cells[row, colOffset + header.Value] = leftHeader;
                }

                if (prefix != "" && colOffset + header.Value - 1 > 0)
                {
                    worksheet.Cells[row, colOffset + header.Value - 1] = prefix;
                }

                if (rightHeader != "")
                {
                    worksheet.Cells[row, colOffset + header.Value + 1] = rightHeader;
                }
            }

            var srtCol = SupportMetricsHelper.numToLetter(colOffset + 1);
            var endCol = SupportMetricsHelper.numToLetter(colOffset + (headers.Count) * 3);

            // format header row
            worksheet.Range[$"{srtCol}{row}", $"{endCol}{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"{srtCol}{row}", $"{endCol}{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range[$"{srtCol}{row}", $"{endCol}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            return worksheet;
        }
    }
}
