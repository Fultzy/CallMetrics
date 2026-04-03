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
    internal class TableHelper
    {
        public static Worksheet WriteTableRow(
            Worksheet worksheet, 
            int row,
            int columnStart,
            List<object> values, 
            int rank = 0, bool bold = false)
        {
            var clSrt = SupportMetricsHelper.numToLetter(columnStart);
            var clEnd = SupportMetricsHelper.numToLetter(columnStart + values.Count - 1);

            for (int i = 0; i < values.Count; i++)
            {
                if (bold)
                    worksheet.Range[clSrt + (row + rank), clEnd + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, columnStart + i] = values[i];
            }

            return worksheet;
        }


        public static Worksheet WriteTwoSidedTable(
            Worksheet worksheet, 
            int row,
            int columnStart,
            string value1, 
            object value2, 
            int rank = 0, bool bold = false)
        {
            var cl1 = SupportMetricsHelper.numToLetter(columnStart);
            var cl2 = SupportMetricsHelper.numToLetter(columnStart + 1);

            if (bold)
                worksheet.Range[cl1 + (row + rank), cl2 + (row + rank)].Font.Bold = true;

            worksheet.Cells[row + rank, cl1] = value1;
            worksheet.Cells[row + rank, cl2] = value2;

            return worksheet;
        }


        public static Worksheet WriteThreeSidedTable(
            Worksheet worksheet,
            int row,
            int columnStart,
            string value1,
            object value2,
            object value3,
            int rank = 0, bool bold = false)
        {
            var cl1 = SupportMetricsHelper.numToLetter(columnStart);
            var cl2 = SupportMetricsHelper.numToLetter(columnStart + 1);
            var cl3 = SupportMetricsHelper.numToLetter(columnStart + 2);

            if (bold)
                worksheet.Range[cl1 + (row + rank), cl3 + (row + rank)].Font.Bold = true;

            worksheet.Cells[row + rank, cl1] = value1;
            worksheet.Cells[row + rank, cl2] = value2;
            worksheet.Cells[row + rank, cl3] = value3;

            return worksheet;
        }
    }
}
