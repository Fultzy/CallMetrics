using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using CallMetrics.Models;
using CallMetrics.Utilities;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Media;
using CallMetrics.Controllers.Readers.Nextiva;
using System.Reflection.Metadata.Ecma335;
using CallMetrics.Data;

namespace CallMetrics.Controllers.Generators.WorkSheets
{

    public class SupportRepMetrics
    {
        public Rep AverageRep = null;
        public Rep TotalRep = null;

        public Dictionary<string, int> TicketsFormulaSourceRows = new();
        public int rankCount = Settings.RankedRepsCount;

        private double percentageStep = 0;
        private double currentProgress = 0;

        public event EventHandler<int> ProgressChanged;

        public Worksheet Create(List<Rep> importReps, Worksheet worksheet)
        {
            ///////////////////////////////////////////////////////////////
            // This method will create the Support Metric Report Worksheet.
            // This worksheet will list out all metrics for support within
            // the specified time span. it will list out company wide metrics
            // as well as metrics for each individual support agent.

            // This report will be very different from the other reports
            // it will show staff wide metrics and rankings of support reps
            // based on their metrics and showing several categories.
            ///////////////////////////////////////////////////////////////


            /////////////////////////////////////////// WorkSheet Setup
            // setup worksheet
            worksheet.Name = "Support Metric Report";
            Settings.Load();

            // set the row counter
            int row = 1;

            var departments = importReps.Where(r => Settings.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IsDepartment == true)).ToList();

            var reps = importReps.Where(r => Settings.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IncludeInMetrics == true)).ToList();

            if (reps.Count == 0)
                reps = importReps;

            if (rankCount > reps.Count)
                rankCount = reps.Count;

            percentageStep = CalculateProgressSteps(reps.Count + departments.Count, rankCount);

            DateRange dateRange = new DateRange();
            dateRange.StartDate = MetricsData.Calls.Min(c => c.DateTime);
            dateRange.EndDate = MetricsData.Calls.Max(c => c.DateTime);

            /////////////////////////////////////////// Begin MetricsData Entry
            // Setup top row
            worksheet.Range["A1", "T1"].Font.Size = 14;
            worksheet.Range["A1", "T1"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            // add date range to the worksheet
            worksheet.Cells[row, 1] = "Date Range:";
            worksheet.Range["B" + 1, "C" + 1].Merge();
            worksheet.Cells[row, 2] = dateRange.ToString();

            // Format DateRange
            worksheet.Range["B1", "B2"].Font.Bold = true;
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["B2"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // set Total Normal Working Days cell
            worksheet.Cells[row, 4] = "Total Working Days:";
            worksheet.Range["D" + 1, "E" + 1].Merge();

            worksheet.Cells[row, 6] = dateRange.WorkDays();
            worksheet.Cells[row, 6].Interior.Color = XlRgbColor.rgbLightYellow;
            worksheet.Cells[row, 6].Font.Bold = true;

            row += 3;

            /////////// Table 1 - Company Wide Metrics
            if (departments.Count > 0)
            {
                worksheet = AddDepartmentHeader(departments, worksheet, row);
                row++;
                worksheet = AddDepartmentMetrics(departments, worksheet, row);
                
                row += departments.Count;
            }

            row += 3;

            /////////// Table 2 - User Metrics
            worksheet = AddGeneralTableHeader(worksheet, row);
            row++;
            worksheet = AddGeneralTableMetrics(reps, worksheet, row);

            row += reps.Count + 2;
            row += 3;

            /////////// Table 3 - Average Lineup
            worksheet = AddAverageLineupHeader(worksheet, row);
            row++;
            worksheet = AddAverageLineupMetrics(reps, worksheet, row, out int newRow);
            row = row + rankCount + 2;

            row++;
            row++;

            /////////// Individual Teams Tables
            worksheet = AddEachTeamsMetrics(reps, worksheet, row);

            // format the worksheet
            worksheet = FormatWorksheet(worksheet, row, rankCount);

            Console.WriteLine(" Done!");
            return worksheet;
        }
        

        ///////// FIRST TABLE
        private Worksheet AddDepartmentHeader(List<Rep> reps, Worksheet worksheet, int row)
        {
            // First header row on worksheet
            worksheet.Cells[row - 1, 1] = "Departments";
            worksheet.Cells[row - 1, 1].Font.Bold = true;
            worksheet.Cells[row - 1, 1].Font.Size = 14;
            worksheet.Cells[row - 1, 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            worksheet.Cells[row, 1] = "Name";
            worksheet.Cells[row, 2] = "Tickets";
            worksheet.Cells[row, 3] = "Wkd Tickets";
            worksheet.Cells[row, 4] = "Adj Tickets";
            worksheet.Cells[row, 5] = "Calls ↓"; // list is sorted
            worksheet.Cells[row, 6] = "Wkd Calls";
            worksheet.Cells[row, 7] = "Internal Calls";
            worksheet.Cells[row, 8] = "Adj Calls";
            worksheet.Cells[row, 9] = "Calls/Tickets";
            worksheet.Cells[row, 10] = "Avg Call Time";
            worksheet.Cells[row, 11] = "Total Ph Time";
            worksheet.Cells[row, 12] = "Calls > 30m";
            worksheet.Cells[row, 13] = "> 30%";
            worksheet.Cells[row, 14] = "Calls > 1h";
            worksheet.Cells[row, 15] = "> 1h%";
            worksheet.Cells[row, 16] = "Absences";
            worksheet.Cells[row, 17] = "Tickets/Day";
            worksheet.Cells[row, 18] = "Calls/Day";
            worksheet.Cells[row, 20] = "Notes";

            // format header row
            worksheet.Range[$"A{row}", $"T{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"T{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range["A" + row, "T" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddDepartmentMetrics(List<Rep> departments, Worksheet worksheet, int row)
        {
            // sort departments by total calls descending
            departments = departments.OrderByDescending(u => u.TotalCalls).ToList();

            var totalUser = CreateTotalUser(departments);
            departments.Insert(0, totalUser);

            foreach (var dept in departments)
            {
                worksheet = AddUserMetrics(dept, worksheet, row);
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                row++;
            }

            //// alternate row colors with smoke
            //if (row % 2 == 0)
            //{
            //    worksheet.Range["A" + row, "T" + row].Interior.Color = XlRgbColor.rgbWhiteSmoke;
            //}

            return worksheet;
        }




        ///////// SECOND TABLE
        private Worksheet AddGeneralTableHeader(Worksheet worksheet, int row)
        {
            // First header row on worksheet
            worksheet.Cells[row - 1, 1] = "Calls";
            worksheet.Cells[row - 1, 1].Font.Bold = true;
            worksheet.Cells[row - 1, 1].Font.Size = 14;
            worksheet.Cells[row - 1, 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            worksheet.Cells[row, 1] = "Name  ↓"; // list is sorted
            worksheet.Cells[row, 2] = "Tickets";
            worksheet.Cells[row, 3] = "Wkd Tickets";
            worksheet.Cells[row, 4] = "Adj Tickets";
            worksheet.Cells[row, 5] = "Calls"; 
            worksheet.Cells[row, 6] = "Wkd Calls";
            worksheet.Cells[row, 7] = "Internal Calls";
            worksheet.Cells[row, 8] = "Adj Calls";
            worksheet.Cells[row, 9] = "Calls/Tickets";
            worksheet.Cells[row, 10] = "Avg Call Time";
            worksheet.Cells[row, 11] = "Total Ph Time";
            worksheet.Cells[row, 12] = "Calls > 30m";
            worksheet.Cells[row, 13] = "> 30%";
            worksheet.Cells[row, 14] = "Calls > 1h";
            worksheet.Cells[row, 15] = "> 1h%";
            worksheet.Cells[row, 16] = "Absences";
            worksheet.Cells[row, 17] = "Tickets/Day";
            worksheet.Cells[row, 18] = "Calls/Day";
            worksheet.Cells[row, 20] = "Notes";


            // format header row
            worksheet.Range[$"A{row}", $"T{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"T{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            // add bottom border to this header
            worksheet.Range["A" + row, "T" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // left aligh the name cell
            worksheet.Cells[row, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // center align rest of header row
            worksheet.Range["B" + row, "T" + row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddGeneralTableMetrics(List<Rep> reps, Worksheet worksheet, int row)
        {
            // sort then add total then average rep to the front
            reps = reps.OrderBy(u => u.Name).ToList();
            var repInt= reps.Count;

            var totalUser = CreateTotalUser(reps);
            var averageUser = CreateAverageUser(reps);

            reps.Insert(0, totalUser);
            reps.Insert(1, averageUser);

            // setup Total and Average users Ticket Formula

            int stRow = row + 2;
            int edRow = row + repInt + 1; 

            int tRow = row;
            worksheet.Cells[tRow, 2].Formula = "=SUM(B" + (stRow) + ":B" + (edRow) + ")";
            worksheet.Cells[tRow, 3].Formula = "=SUM(C" + (stRow) + ":C" + (edRow) + ")";

            int aRow = tRow + 1;
            worksheet.Cells[aRow, 2].Formula = "=IF(SUM(B" + (stRow) + ":B" + (edRow) + ")<=0,0,ROUND(AVERAGE(B" + (stRow) + ":B" + (edRow) + "),2))";
            worksheet.Cells[aRow, 3].Formula = "=IF(SUM(C" + (stRow) + ":C" + (edRow) + ")<=0,0,ROUND(AVERAGE(C" + (stRow) + ":C" + (edRow) + "),2))";

            int userCtr = 0;
            foreach (var rep in reps)
            {
                worksheet = AddUserMetrics(rep, worksheet, row);

                if (rep.Name == "-- TOTAL --")
                {
                    // =SUM(P10:P28)
                    worksheet.Cells[row, 16].Formula = $"=SUM(P{stRow}:P{edRow})";
                }
                else if (rep.Name == "-- AVERAGE --")
                {
                    // =IF(SUM(P10:P28)<=0,0,ROUND(AVERAGE(P10:P28),2))
                    worksheet.Cells[row, 16].Formula = $"=IF(SUM(P{stRow}:P{edRow})<=0,0,ROUND(AVERAGE(P{stRow}:P{edRow}),2))";
                }
                else
                {
                    worksheet.Cells[row, 16] = 0; // absences

                }


                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                userCtr++;
                row++;
            }

            return worksheet;
        }



        private Worksheet AddUserMetrics(Rep rep, Worksheet worksheet, int row)
        {
            if (rep == null)
                return worksheet;

            var team = Settings.Teams.FirstOrDefault(t => t.Members.Contains(rep.Name));

            // add the record to the worksheet
            if (team.IsDepartment)
                worksheet.Cells[row, 1] = rep.Name;
            else
                worksheet.Cells[row, 1] = rep.LastInitial();

            if (rep.Name != "-- TOTAL --" && rep.Name != "-- AVERAGE --")
            {
                worksheet.Cells[row, 2] = rep.TotalTickets;
                worksheet.Cells[row, 3] = rep.WeekendTickets;
            }

            worksheet.Cells[row, 5] = rep.TotalCalls;
            worksheet.Cells[row, 6] = rep.WeekendCalls;
            worksheet.Cells[row, 7] = rep.InternalCalls;
            worksheet.Cells[row, 8] = rep.AdjustedCalls();
            worksheet.Cells[row, 10] = rep.FormattedDuration(Convert.ToInt32(rep.AverageDuration()));
            worksheet.Cells[row, 11] = rep.FormattedDuration(rep.TotalDuration);
            worksheet.Cells[row, 12] = rep.CallsOver30;
            worksheet.Cells[row, 13] = rep.Over30Percentage();
            worksheet.Cells[row, 14] = rep.CallsOver60;
            worksheet.Cells[row, 15] = rep.Over60Percentage();

            worksheet.Cells[row, 16] = 0; // absences

            // formula for Absences Calculations
            // =IF(H9/E1>1,ROUND(H9/E1,2),0)
            worksheet.Cells[row, 17].Formula = $"=IF(F1>P{row},ROUND(D{row}/(F1-P{row}),0),0)";
            worksheet.Cells[row, 18].Formula = $"=IF(F1>P{row},ROUND(H{row}/(F1-P{row}),0),0)";

            // Add Adj Tickets and Calls/Ticket formulas
            // =IF(B9-C9<0,0,B9-C9)
            worksheet.Cells[row, 4].Formula = "=IF(B" + row + "-C" + row + "<0,0,B" + row + " - C" + row + ")";
            worksheet.Cells[row, 9].Formula = "=IF(D" + row + "<=0,0,ROUND(H" + row + "/D" + row + ", 2))";

            if (!TicketsFormulaSourceRows.ContainsKey(rep.Name))
            {
                TicketsFormulaSourceRows.Add(rep.Name, row);
            }

            // add borders to the row
            worksheet.Range["A" + row, "R" + row].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            worksheet.Range["A" + row, "R" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // left aligh the name cell
            worksheet.Range["A" + row].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // center align rest of row
            worksheet.Range["B" + row, "R" + row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // format the row
            // if the username is average make the row bold
            if (rep.Name == "-- AVERAGE --")
            {
                worksheet.Range["A" + row, "R" + row].Font.Bold = true;
            }

            // alternate row colors with smoke
            if (row % 2 == 0)
            {
                worksheet.Range["A" + row, "T" + row].Interior.Color = XlRgbColor.rgbWhiteSmoke;
            }

            return worksheet;
        }


        ///////// THIRD TABLE
        private Worksheet AddAverageLineupHeader(Worksheet worksheet, int row)
        {
            worksheet.Cells[row - 1, 1] = "Averages";
            worksheet.Cells[row - 1, 1].Font.Bold = true;
            worksheet.Cells[row - 1, 1].Font.Size = 14;
            worksheet.Cells[row - 1, 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;


            worksheet.Cells[row, 1] = "Rankings";
            worksheet.Cells[row, 2] = "Adj Tickets";
            worksheet.Cells[row, 4] = "Adj Calls";
            worksheet.Cells[row, 6] = "Calls/Ticket";
            worksheet.Cells[row, 8] = "Avg Call Time";
            worksheet.Cells[row, 10] = "Total Ph Time";
            worksheet.Cells[row, 12] = " > 30m %";
            worksheet.Cells[row, 14] = " > 1h %";
            worksheet.Cells[row, 16] = "Tickets/Day";
            worksheet.Cells[row, 18] = "Calls/Day";

            // format header row
            worksheet.Range[$"A{row}", $"P{row}"].Font.Bold = true;

            // Center and join rank headers
            worksheet.Range["B" + row, "C" + row].Merge();
            worksheet.Range["D" + row, "E" + row].Merge();
            worksheet.Range["F" + row, "G" + row].Merge();
            worksheet.Range["H" + row, "I" + row].Merge();
            worksheet.Range["J" + row, "K" + row].Merge();
            worksheet.Range["L" + row, "M" + row].Merge();
            worksheet.Range["N" + row, "O" + row].Merge();
            worksheet.Range["P" + row, "Q" + row].Merge();
            worksheet.Range["R" + row, "S" + row].Merge();

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"S{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range["A" + row, "S" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddAverageLineupMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            // this table is going to display data in columns. each column will be a different metric and
            // will be sorted vertically by the metric with that genReps name to the left. 
            // there will be a sudo rep added to the list that will be the average of all the genReps.

            var averageRep = CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            // setup sorted lists for each metric, adjusted calls, average call time, total phone time, calls over 30, calls over 30 %, calls over 60, calls over 60 %
            List<Rep> adjCalls = reps.OrderByDescending(u => u.AdjustedCalls()).ToList();
            List<Rep> avgCallTime = reps.OrderBy(u => u.AverageDuration()).ToList();
            List<Rep> totalPhoneTime = reps.OrderByDescending(u => u.TotalDuration).ToList();

            List<Rep> callsOver30Percent = reps.OrderBy(u => u.Over30PercentFloat()).ToList();
            List<Rep> callsOver60Percent = reps.OrderBy(u => u.Over60PercentFloat()).ToList();

            int rank = 0;
            int startRow = row - reps.Count - 4;
            int endRow = row - 5;


            worksheet.Cells[row + rank, 2].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, D{startRow}:D{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            worksheet.Cells[row + rank, 6].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, I{startRow}:I{endRow}, 1), SEQUENCE({rankCount},1,1,1))";
            foreach (Rep rep in reps.Take(rankCount))
            {

                // Value =LARGE(D9:D23, 10)
                worksheet.Cells[row + rank, 3].Formula =
                    $"=LARGE(D{startRow}:D{endRow}, A{row + rank})";

                // Value =SMALL(I9:I23, 10)
                worksheet.Cells[row + rank, 7].Formula =
                    $"=SMALL(I{startRow}:I{endRow}, A{row + rank})";
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            rank = 0;
            foreach (Rep user in adjCalls.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["D" + (row + rank), "E" + (row + rank)].Font.Bold = true;


                worksheet.Cells[row + rank, 4] = user.LastInitial();
                worksheet.Cells[row + rank, 5] = user.AdjustedCalls();
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            foreach (Rep user in avgCallTime.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["H" + (row + rank), "I" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 8] = user.LastInitial();
                worksheet.Cells[row + rank, 9] = user.FormattedDuration(Convert.ToInt32(user.AverageDuration()));
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            foreach (Rep user in totalPhoneTime.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["J" + (row + rank), "K" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 10] = user.LastInitial();
                worksheet.Cells[row + rank, 11] = user.FormattedDuration(user.TotalDuration);
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            foreach (Rep user in callsOver30Percent.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["L" + (row + rank), "M" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 12] = user.LastInitial();
                worksheet.Cells[row + rank, 13] = user.Over30Percentage();
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            foreach (Rep user in callsOver60Percent.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["N" + (row + rank), "O" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 14] = user.LastInitial();
                worksheet.Cells[row + rank, 15] = user.Over60Percentage();
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            worksheet.Cells[row + rank, 16].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, Q{startRow}:Q{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            foreach (Rep user in reps.Take(rankCount))
            {
                // Tickets/Day
                // NAME =INDEX(SORTBY(A9:A28, Q9:Q28, -1), SEQUENCE(20,1,1,1))

                worksheet.Cells[row + rank, 17].Formula =
                   $"=LARGE(Q{startRow}:Q{endRow}, A{row + rank})";

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            rank = 0;
            worksheet.Cells[row + rank, 18].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, R{startRow}:R{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            foreach (Rep user in reps.Take(rankCount))
            {
                // Calls/Day
                //worksheet.Cells[row + rank, 18].Formula =
                //$"=INDEX(A{9}:A{9 + genReps.Count - 1}, MATCH(LARGE(R{9}:R{9 + genReps.Count - 1}, A{row + rank}), R{9}:R{9 + genReps.Count - 1}, 0))";


                worksheet.Cells[row + rank, 19].Formula =
                   $"=LARGE(R{startRow}:R{endRow}, A{row + rank})";

                //worksheet.Cells[row + rank, 18].Formula = $"=IF(H{9 + rank}/(F1-P{9 + rank})>1,ROUND(H{9 + rank}/(F1-Q{9 + rank}),0),0)";
                //worksheet.Cells[row + rank, 19].Formula = $"=IF(H{9 + rank}/(F1-P{9 + rank})>1,ROUND(H{9 + rank}/(F1-R{9 + rank}),0),0)";

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            worksheet = FormatLineupTable(worksheet, row, rankCount);
            newRow = row + rankCount;
            return worksheet;
        }

        private Worksheet FormatLineupTable(Worksheet worksheet, int row, int rankCount)
        {
            // format the rank table
            worksheet.Range["A" + row, "S" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // iterate through rows starting at row
            for (int i = 0; i < rankCount; i++)
            {
                // Ranking numbers
                worksheet.Cells[(i + row), 1] = i + 1;
                worksheet.Cells[(i + row), 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // alternate row colors with smoke
                if (i % 2 == 0)
                {
                    worksheet.Range["A" + (i + row), "T" + (i + row)].Interior.Color = XlRgbColor.rgbWhiteSmoke;
                }
            }

            return worksheet;
        }


        ////// Teams Tables
        private Worksheet AddTeamHeader(Worksheet worksheet, int row)
        {
            worksheet.Cells[row, 1] = "Name"; 
            worksheet.Cells[row, 2] = "Tickets";
            worksheet.Cells[row, 3] = "Wkd Tickets";
            worksheet.Cells[row, 4] = "Adj Tickets";
            worksheet.Cells[row, 5] = "Calls ↓"; // list is sorted
            worksheet.Cells[row, 6] = "Wkd Calls";
            worksheet.Cells[row, 7] = "Internal Calls";
            worksheet.Cells[row, 8] = "Adj Calls";
            worksheet.Cells[row, 9] = "Calls/Tickets";
            worksheet.Cells[row, 10] = "Avg Call Time";
            worksheet.Cells[row, 11] = "Total Ph Time";
            worksheet.Cells[row, 12] = "Calls > 30m";
            worksheet.Cells[row, 13] = "> 30%";
            worksheet.Cells[row, 14] = "Calls > 1h";
            worksheet.Cells[row, 15] = "> 1h%";
            worksheet.Cells[row, 16] = "Absences";
            worksheet.Cells[row, 17] = "Tickets/Day";
            worksheet.Cells[row, 18] = "Calls/Day";
            worksheet.Cells[row, 20] = "Notes";


            // format header row
            worksheet.Range[$"A{row}", $"T{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"T{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            // add bottom border to this header
            worksheet.Range["A" + row, "T" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // left aligh the name cell
            worksheet.Cells[row, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // center align rest of header row
            worksheet.Range["B" + row, "T" + row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddEachTeamsMetrics(List<Rep> reps, Worksheet worksheet, int row)
        {
            foreach (var team in Settings.Teams)
            {
                if (team.Members.Count == 0)
                    continue;

                // skip team if not enabled for teams table
                if (team.IncludeInMetrics == false)
                    continue;

                // add team name as header
                worksheet.Cells[row, 1] = team.Name;
                worksheet.Cells[row, 1].Font.Bold = true;
                worksheet.Cells[row, 1].Font.Size = 14;
                worksheet.Cells[row, 1].Interior.Color = (int)XlRgbColor.rgbLightSteelBlue;
               
                row++;

                // add the rep header
                worksheet = AddTeamHeader(worksheet, row);
                row++;

                var tRow = row;
                var aRow = tRow + 1;

                // get the genReps for this team
                List<Rep> teamReps = new();
                foreach (var repName in team.Members)
                {
                    var rep = reps.FirstOrDefault(r => r.Name == repName);
                    if (rep != null)
                    {
                        if (rep.TotalCalls > 0)
                            teamReps.Add(rep);
                    }
                }

                // create average and total genReps for this team
                var teamAverageRep = CreateAverageUser(teamReps);
                var teamTotalRep = CreateTotalUser(teamReps);

                // sort then add total then average rep to the front
                teamReps = teamReps.OrderByDescending(u => u.TotalCalls).ToList();
                var teamCount = teamReps.Count;

                teamReps.Insert(0, teamTotalRep);
                teamReps.Insert(1, teamAverageRep);

                foreach (var user in teamReps)
                {
                    worksheet = AddUserMetrics(user, worksheet, row);

                    // reference B{row} & C{row} from second table for ticket counts
                    if (TicketsFormulaSourceRows.ContainsKey(user.Name))
                    {
                        int sourceRow = TicketsFormulaSourceRows[user.Name];
                        worksheet.Cells[row, 2].Formula = "=B" + sourceRow;
                        worksheet.Cells[row, 3].Formula = "=C" + sourceRow;
                        worksheet.Cells[row, 16].Formula = "=P" + sourceRow;
                    }

                    row++;
                }
                row++;

                // override average and total rep formula for this team only.
                // setup Total and Average users Ticket Formula

                var startRow = row - teamCount - 1;
                var endRow = row - 2;
                
                // Total
                worksheet.Cells[tRow, 2].Formula = $"=SUM(B{startRow}:B{endRow})";
                worksheet.Cells[tRow, 3].Formula = $"=SUM(C{startRow}:C{endRow})";
                worksheet.Cells[tRow, 16].Formula = $"=SUM(P{startRow}:P{endRow})";

                // Average
                worksheet.Cells[aRow, 2].Formula = $"=IF(SUM(B{startRow}:B{endRow})<=0,0,ROUND(AVERAGE(B{startRow}:B{endRow}),2))";
                worksheet.Cells[aRow, 3].Formula = $"=IF(SUM(C{startRow}:C{endRow})<=0,0,ROUND(AVERAGE(C{startRow}:C{endRow}),2))";
                worksheet.Cells[aRow, 16].Formula = $"=IF(SUM(P{startRow}:P{endRow})<=0,0,ROUND(AVERAGE(P{startRow}:P{endRow}),2))";


                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
            }
            return worksheet;
        }


        ///////// formatting and helpers
        private Rep CreateTotalUser(List<Rep> reps)
        {
            Rep totalUser = new Rep();

            // add all the call records to the total rep
            totalUser.Name = "-- TOTAL --";

            // get the average metrics
            int totalCalls = 0;
            int inboundCalls = 0;
            int outboundCalls = 0;

            int totalPhoneTime = 0;
            int inboundPhoneTime = 0;
            int outboundPhoneTime = 0;

            int callsOver30 = 0;
            int callsOver60 = 0;

            int WeekendCalls = 0;
            int InternalCalls = 0;

            foreach (var rep in reps)
            {
                if (rep.Name == "-- AVERAGE --") continue; // skip the average rep
                if (rep.Name == "-- TOTAL --") continue; // skip the total rep

                // add total calls
                totalCalls += rep.TotalCalls;
                inboundCalls += rep.InboundCalls;
                outboundCalls += rep.OutboundCalls;

                // add timing
                totalPhoneTime += rep.TotalDuration;
                inboundPhoneTime += rep.InboundDuration;
                outboundPhoneTime += rep.OutboundDuration;

                // add calls over 30 and 60 minutes
                callsOver30 += rep.CallsOver30;
                callsOver60 += rep.CallsOver60;

                InternalCalls += rep.InternalCalls;
                WeekendCalls += rep.WeekendCalls;
            }

            // set the Total rep metrics

            totalUser.TotalCalls = totalCalls;
            totalUser.InboundCalls = inboundCalls;
            totalUser.OutboundCalls = outboundCalls;

            totalUser.TotalDuration = totalPhoneTime;
            totalUser.InboundDuration = inboundPhoneTime;
            totalUser.OutboundDuration = outboundPhoneTime;

            totalUser.WeekendCalls = WeekendCalls;
            totalUser.InternalCalls = InternalCalls;

            totalUser.CallsOver30 = callsOver30;
            totalUser.CallsOver60 = callsOver60;

            return totalUser;
        }

        private Rep CreateAverageUser(List<Rep> reps)
        {
            // create new rep
            Rep avgUser = new Rep();
            avgUser.Name = "-- AVERAGE --";

            // get the average metrics
            int userCount = reps.Count;

            if (userCount == 0)
            {
                Console.WriteLine("No genReps found");
                return avgUser;
            }

            int totalCalls = 0;
            int weekendCalls = 0;
            int internalCalls = 0;
            int adjustedCalls = 0;

            int totalPhoneTime = 0;
            int callsOver30 = 0;
            int callsOver60 = 0;

            // add from all genReps
            foreach (var user in reps)
            {
                if (user.Name == "-- TOTAL --") continue; // skip the total rep
                if (user.Name == "-- AVERAGE --") continue; // skip the average rep

                // add total calls
                totalCalls += user.TotalCalls;
                weekendCalls += user.WeekendCalls;
                internalCalls += user.InternalCalls;
                adjustedCalls += user.AdjustedCalls();

                // add timing
                totalPhoneTime += user.TotalDuration;

                // add calls over 30 and 60 minutes
                callsOver30 += user.CallsOver30;
                callsOver60 += user.CallsOver60;
            }


            avgUser.TotalCalls = totalCalls / userCount;
            avgUser.WeekendCalls = weekendCalls / userCount;
            avgUser.InternalCalls = internalCalls / userCount;
            avgUser.TotalDuration = totalPhoneTime / userCount;
            avgUser.CallsOver30 = callsOver30 / userCount;
            avgUser.CallsOver60 = callsOver60 / userCount;

            return avgUser;
        }

        private Worksheet FormatWorksheet(Worksheet worksheet, int row, int rankCount)
        {
            // format columns to fit
            worksheet.Columns.AutoFit();

            // center align the entire sheet
            worksheet.Range["B1", "T" + (row)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // set the daterange column aligh left
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // make all columns wider
            worksheet.Range["B1", "S1"].ColumnWidth = 14;
            worksheet.Cells[1,1].ColumnWidth = 16;

            // make the notes column wide enough for the rep to make notes
            worksheet.Range["T1", "T" + (row - 1)].ColumnWidth = 70;

            // Freeze column A
            worksheet.Application.ActiveWindow.SplitColumn = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;


            return worksheet;
        }

        public static double CalculateProgressSteps(int repCount, int rankCount)
        {
            // each row for average lineup table
            var total = repCount += rankCount ; // 6 columns per rank

            // each team table
            foreach (var team in Settings.Teams)
            {
                if (team.Members.Count == 0)
                    continue;

                if (!team.IncludeInMetrics) continue; // skip ignored teams
                total += team.Members.Count * 6;
            }

            return total / 120;
        }


        private string numToLetter(int num)
        {
            string letter = "";
            switch (num)
            {
                case 1:
                    letter = "A";
                    break;
                case 2:
                    letter = "B";
                    break;
                case 3:
                    letter = "C";
                    break;
                case 4:
                    letter = "D";
                    break;
                case 5:
                    letter = "E";
                    break;
                case 6:
                    letter = "F";
                    break;
                case 7:
                    letter = "G";
                    break;
                case 8:
                    letter = "H";
                    break;
                case 9:
                    letter = "I";
                    break;
                case 10:
                    letter = "J";
                    break;
                case 11:
                    letter = "K";
                    break;
                case 12:
                    letter = "L";
                    break;
                case 13:
                    letter = "M";
                    break;
                case 14:
                    letter = "N";
                    break;
                case 15:
                    letter = "O";
                    break;
                case 16:
                    letter = "P";
                    break;
                case 17:
                    letter = "Q";
                    break;
                case 18:
                    letter = "R";
                    break;
                case 19:
                    letter = "S";
                    break;
                case 20:
                    letter = "T";
                    break;
                case 21:
                    letter = "U";
                    break;
                case 22:
                    letter = "V";
                    break;
                case 23:
                    letter = "W";
                    break;
                case 24:
                    letter = "X";
                    break;
                case 25:
                    letter = "Y";
                    break;
                case 26:
                    letter = "Z";
                    break;
                default:
                    letter = "A";
                    break;
            }

            return letter;
        }
    }
}
