using CallMetrics.Controllers.Generators.WorkSheets.RepMetrics;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows;

namespace CallMetrics.Controllers.Generators.WorkSheets.Helpers
{
    public static class FormatHelper
    {
        public static Color FontColor = Color.FromArgb(0, 255, 255, 255);
        public static Color MainColor = Color.FromArgb(0, 68, 114, 196);
        public static Color AccentOne = Color.FromArgb(0, 96, 155, 213);
        public static Color AccentTwo = Color.FromArgb(0, 146, 208, 80);
        public static Color OnRowColor = Color.FromArgb(0, 226, 239, 218);
        public static Color OffRowColor = Color.FromArgb(0, 226, 239, 218);
        //public static XlRgbColor OffRowColor = XlRgbColor.rgbWhiteSmoke;

        public static Worksheet FormatWorksheet(Worksheet worksheet, int row)
        {
            // set whole sheet center
            worksheet.Rows.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Rows.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // set the daterange column aligh left
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // make all columns wider
            worksheet.Range["B1", "AA1"].ColumnWidth = 15;
            worksheet.Cells[1, 1].ColumnWidth = 16;



            // Freeze column A
            worksheet.Application.ActiveWindow.SplitColumn = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;


            return worksheet;
        }

        public static Worksheet HighlighRow(Worksheet worksheet, int row, string colStart, string colEnd)
        {
            var range = worksheet.Range[$"{colStart}{row}", $"{colEnd}{row}"];

            range.Interior.Color = AccentTwo;
            range.RowHeight = 8;

            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;

            return worksheet;
        }

        public static Worksheet SeperateRow(Worksheet worksheet, int row, string colStart, string colEnd)
        {
            var range = worksheet.Range[$"{colStart}{row}", $"{colEnd}{row}"];

            range.Interior.Color = AccentTwo;
            range.RowHeight = 3;

            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;

            return worksheet;
        }

        public static Worksheet MainHeader(Worksheet worksheet, int row, string colStart, string colEnd)
        {
            var range = worksheet.Range[$"{colStart}{row}", $"{colEnd}{row}"];

            range.Font.Size = 12;
            range.Font.Color = FontColor;

            range.Interior.Color = MainColor;

            return worksheet;
        }

        public static Worksheet HeaderTitle(Worksheet worksheet, int row, string colStart, string colEnd, string titEnd)
        {
            var range = worksheet.Range[$"{colStart}{row}", $"{colEnd}{row}"];
            var titRange = worksheet.Range[$"{colStart}{row}", $"{titEnd}{row}"];

            range.Font.Size = 14;
            range.Font.Bold = true;
            range.Font.Color = FontColor;

            titRange.MergeCells = true;

            //range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Interior.Color = MainColor;

            return worksheet;
        }


        public static Worksheet HeaderRow(Worksheet worksheet, int row, string colStart, string colEnd)
        {
            var range = worksheet.Range[$"{colStart}{row}", $"{colEnd}{row}"];

            range.Font.Bold = true;
            range.Font.Color = FontColor;
            range.RowHeight = 30;

            range.WrapText = true;

            //range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Interior.Color = AccentOne;

            return worksheet;
        }

        public static Worksheet LineupTable(Worksheet worksheet, int row, int rankCount, int lenght, int offset = 0)
        {
            var colStart = SupportMetricsHelper.NumToLetter(offset);
            var colEnd = SupportMetricsHelper.NumToLetter(offset + lenght);

            // format the rank table
            worksheet.Range[colStart + row, colEnd + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // iterate through rows starting at row
            for (int i = 0; i < rankCount; i++)
            {
                // Ranking numbers
                worksheet.Cells[(i + row), 1] = i + 1;

                // alternate row colors with smoke
                if (i % 2 == 0)
                {
                    worksheet.Range[colStart + (i + row), colEnd + (i + row)].Interior.Color = OffRowColor;
                }
            }

            return worksheet;
        }

        internal static Worksheet JoinCells(Worksheet worksheet, string cell1, string cell2)
        {
            worksheet.Range[cell1, cell2].Merge();
            return worksheet;
        }
    }
}
