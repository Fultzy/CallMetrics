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

namespace CallMetrics.Controllers.Generators.WorkSheets
{

    public class SupportRepMetrics
    {

        public RepData AverageRep = null;
        public RepData TotalRep = null;

        public Dictionary<string, int> TicketsFormulaSourceRows = new();

        public int rankCount = 10; // number of reps to show in the average lineup table

        public Worksheet Create(List<RepData> reps, Worksheet worksheet)
        {
            ///////////////////////////////////////////////////////////////
            // This method will create the Support Metric Report Worksheet.
            // This worksheet will list out all metrics for support within
            // the specified time span. it will list out company wide metrics
            // as well as metrics for each individual support agent.

            // This report will be very different from the other reports
            // it will show staff wide metrics and rankings of supprt reps
            // based on their metrics and showing several categories.
            ///////////////////////////////////////////////////////////////


            /////////////////////////////////////////// WorkSheet Setup

            // setup worksheet
            worksheet.Name = "Support Metric Report";

            // set the row counter
            int row = 1;
            
            DateRange dateRange = new DateRange();
            dateRange.StartDate = reps.Min(r => r.Calls.Min(c => c.Time));
            dateRange.EndDate = reps.Max(r => r.Calls.Max(c => c.Time));

            // Add Sheet wide conditioning
            ApplyBoldIfAverage(worksheet);

            // Create average and total Users
            AverageRep = CreateAverageUser(reps);
            TotalRep = CreateTotalUser(reps);

            /////////////////////////////////////////// Begin Data Entry
            // add date range to the worksheet
            worksheet.Cells[row, 1] = "Date Range:";
            worksheet.Cells[row, 2] = dateRange.ToString();

            // Format DateRange
            worksheet.Range["B1", "B2"].Font.Bold = true;
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["B2"].HorizontalAlignment = XlHAlign.xlHAlignLeft;


            row++;
            row++;

            /////////// Table 1 - Company Wide Metrics
            // list the support wide metrics here
            worksheet = AddSupportHeader(worksheet, row);
            row++;
            worksheet = AddSupportWideMetrics(reps, worksheet, row);


            row++;
            row++;
            row++;

            /////////// Table 2 - User Metrics
            // list each support agent and their metrics here
            worksheet = AddUserHeader(worksheet, row);
            row++;

            // sort then add total then average rep to the front
            reps = reps.OrderByDescending(u => u.TotalCalls).ToList();
            reps.Insert(0, TotalRep);
            reps.Insert(1, AverageRep);

            // setup Total and Average users Ticket Formula
            int tRow = row;
            worksheet.Cells[tRow, 2].Formula = "=SUM(B" + (tRow + 2) + ":B" + (tRow + reps.Count - 1) + ")";
            worksheet.Cells[tRow, 3].Formula = "=SUM(C" + (tRow + 2) + ":C" + (tRow + reps.Count - 1) + ")";

            int aRow = tRow + 1; //=IF(SUM(C10:C27) <= 0,0,AVERAGE(C10:C27))
            worksheet.Cells[aRow, 2].Formula = "=IF(SUM(B" + (aRow + 1) + ":B" + (aRow + reps.Count - 2) + ")<=0,0,ROUND(AVERAGE(B" + (aRow + 1) + ":B" + (aRow + reps.Count - 2) + "),2))";
            worksheet.Cells[aRow, 3].Formula = "=IF(SUM(C" + (aRow + 1) + ":C" + (aRow + reps.Count - 2) + ")<=0,0,ROUND(AVERAGE(C" + (aRow + 1) + ":C" + (aRow + reps.Count - 2) + "),2))";

            int userCtr = 0;
            foreach (var user in reps.Take(reps.Count + 2)) // adding two for total and average reps
            {

                worksheet = AddUserMetrics(user, worksheet, row);
                userCtr++;
                row++;
            }

            row++;

            /////////// Table 3 - Average Lineup
            // list the average lineup here
            worksheet = AddAverageLineupHeader(worksheet, row);
            row++;

            worksheet = AddAverageLineupMetrics(reps, worksheet, row, rankCount, out int newRow);

            row = row + rankCount + 2;

            /////////// Individual Teams Tables
            worksheet = AddEachTeamsMetrics(reps, worksheet, row);



            // format the worksheet
            worksheet = FormatWorksheet(worksheet, row, rankCount);

            Console.WriteLine(" Done!");
            return worksheet;
        }


        ///////// FIRST TABLE
        private Worksheet AddSupportHeader(Worksheet worksheet, int row)
        {
            // First header row on worksheet
            worksheet.Cells[row, 1] = "Tech Support";
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
            worksheet.Cells[row, 16] = "Notes";

            // format header row
            worksheet.Range[$"A{row}", $"P{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"P{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range["A" + row, "P" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddSupportWideMetrics(List<RepData> reps, Worksheet worksheet, int row)
        {
            // there will be two rows added, one for total numbers across all support agents,
            // the other is for the averages across all support agents
            AddUserMetrics(TotalRep, worksheet, row);
            row++;
            AddUserMetrics(AverageRep, worksheet, row);

            // alternate row colors with smoke
            if (row % 2 == 0)
            {
                worksheet.Range["A" + row, "P" + row].Interior.Color = XlRgbColor.rgbWhiteSmoke;
            }

            return worksheet;
        }




        ///////// SECOND TABLE
        private Worksheet AddUserHeader(Worksheet worksheet, int row)
        {
            worksheet.Cells[row, 1] = "Tech Support";
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
            worksheet.Cells[row, 16] = "Notes";

            // format header row
            worksheet.Range[$"A{row}", $"P{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"P{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            // add bottom border to this header
            worksheet.Range["A" + row, "P" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddUserMetrics(RepData rep, Worksheet worksheet, int row)
        {
            // add the record to the worksheet
            worksheet.Cells[row, 1] = rep.LastInitial();
            worksheet.Cells[row, 5] = rep.TotalCalls;
            worksheet.Cells[row, 6] = rep.WeekendCalls;
            worksheet.Cells[row, 7] = rep.InternalCalls;
            worksheet.Cells[row, 8] = rep.AdjustedCalls();
            worksheet.Cells[row, 10] = rep.FormatedDuration(Convert.ToInt32(rep.AverageDuration()));
            worksheet.Cells[row, 11] = rep.FormatedDuration(rep.TotalDuration);
            worksheet.Cells[row, 12] = rep.CallsOver30;
            worksheet.Cells[row, 13] = rep.Over30Percentage();
            worksheet.Cells[row, 14] = rep.CallsOver60;
            worksheet.Cells[row, 15] = rep.Over60Percentage();

            // Add Adj Tickets and Calls/Ticket formulas
            //worksheet.Cells[row, 2] = "0";
            //worksheet.Cells[row, 3] = "0"; 
            worksheet.Cells[row, 4].Formula = "=IF(B" + row + "-C" + row + "<0,0,B" + row + " - C" + row + ")";
            worksheet.Cells[row, 9].Formula = "=IF(D" + row + "<=0,0,ROUND(H" + row + "/D" + row + ", 2))";

            if (!TicketsFormulaSourceRows.ContainsKey(rep.Name))
            {
                TicketsFormulaSourceRows.Add(rep.Name, row);
            }

            worksheet.Range["A" + row, "O" + row].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            worksheet.Range["A" + row, "O" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // format the row
            // if the username is average make the row bold
            if (rep.Name == "-- AVERAGE --")
            {
                worksheet.Range["A" + row, "P" + row].Font.Bold = true;
            }

            // alternate row colors with smoke
            if (row % 2 == 0)
            {
                worksheet.Range["A" + row, "P" + row].Interior.Color = XlRgbColor.rgbWhiteSmoke;
            }

            return worksheet;
        }




        ///////// THIRD TABLE
        private Worksheet AddAverageLineupHeader(Worksheet worksheet, int row)
        {
            worksheet.Cells[row, 1] = "Rankings";
            worksheet.Cells[row, 3] = "Adj Tickets";
            worksheet.Cells[row, 5] = "Adj Calls";
            worksheet.Cells[row, 7] = "Calls/Ticket";
            worksheet.Cells[row, 9] = "Avg Call Time";
            worksheet.Cells[row, 11] = "Total Ph Time";
            worksheet.Cells[row, 13] = " > 30m %";
            worksheet.Cells[row, 15] = " > 1h %";

            // format header row
            worksheet.Range[$"A{row}", $"P{row}"].Font.Bold = true;

            // Make the background color of the header light blue
            worksheet.Range[$"A{row}", $"O{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range["A" + row, "O" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // freeze the header row
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            return worksheet;
        }

        private Worksheet AddAverageLineupMetrics(List<RepData> reps, Worksheet worksheet, int row, int rankCount, out int newRow)
        {
            // this table is going to display data in columns. each column will be a different metric and
            // will be sorted vertically by the metric with that reps name to the left. 
            // there will be a sudo rep added to the list that will be the average of all the reps.

            // remove the total rep
            reps.Remove(TotalRep);

            // setup sorted lists for each metric, adjusted calls, average call time, total phone time, calls over 30, calls over 30 %, calls over 60, calls over 60 %
            List<RepData> adjCalls = reps.OrderByDescending(u => u.AdjustedCalls()).ToList();
            List<RepData> avgCallTime = reps.OrderBy(u => u.AverageDuration()).ToList();
            List<RepData> totalPhoneTime = reps.OrderByDescending(u => u.TotalDuration).ToList();

            List<RepData> callsOver30Percent = reps.OrderBy(u => u.Over60PercentFloat()).ToList();
            List<RepData> callsOver60Percent = reps.OrderBy(u => u.Over60PercentFloat()).ToList();

            int rank = 0;
            foreach (RepData rep in reps.Take(rankCount))
            {

                // === TOP 10 Adj Tickets ===   
                // NAME =INDEX(A9:A23, MATCH(LARGE(D9:D23, 10), D9:D23, 0))
                worksheet.Cells[row + rank, 2].Formula =
                    $"=INDEX(A{9}:A{9 + reps.Count - 1}, MATCH(LARGE(D{9}:D{9 + reps.Count - 1}, A{row + rank}), D{9}:D{9 + reps.Count - 1}, 0))";

                // Value =LARGE(D9:D23, 10)
                worksheet.Cells[row + rank, 3].Formula =
                    $"=LARGE(D{9}:D{9 + reps.Count - 1}, A{row + rank})";

                // === TOP 10 Calls/Ticket ===
                // NAME =INDEX(A9:A23, MATCH(SMALL(I9:I23, 10), I9:I23, 0))
                worksheet.Cells[row + rank, 6].Formula =
                    $"=INDEX(A{9}:A{9 + reps.Count - 1}, MATCH(SMALL(I{9}:I{9 + reps.Count - 1}, A{row + rank}), I{9}:I{9 + reps.Count - 1}, 0))";

                // Value =SMALL(I9:I23, 10)
                worksheet.Cells[row + rank, 7].Formula =
                    $"=SMALL(I{9}:I{9 + reps.Count - 1}, A{row + rank})";
                rank++;
            }


            rank = 0;
            foreach (RepData user in adjCalls.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["D" + (row + rank), "E" + (row + rank)].Font.Bold = true;


                worksheet.Cells[row + rank, 4] = user.LastInitial();
                worksheet.Cells[row + rank, 5] = user.AdjustedCalls();
                rank++;
            }

            rank = 0;
            foreach (RepData user in avgCallTime.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["H" + (row + rank), "I" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 8] = user.LastInitial();
                worksheet.Cells[row + rank, 9] = user.FormatedDuration(Convert.ToInt32(user.AverageDuration()));
                rank++;
            }

            rank = 0;
            foreach (RepData user in totalPhoneTime.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["J" + (row + rank), "K" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 10] = user.LastInitial();
                worksheet.Cells[row + rank, 11] = user.FormatedDuration(user.TotalDuration);
                rank++;
            }

            rank = 0;
            foreach (RepData user in callsOver30Percent.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["L" + (row + rank), "M" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 12] = user.LastInitial();
                worksheet.Cells[row + rank, 13] = user.Over30Percentage();
                rank++;
            }

            rank = 0;
            foreach (RepData user in callsOver60Percent.Take(rankCount))
            {
                if (user.Name == "-- AVERAGE --") worksheet.Range["N" + (row + rank), "O" + (row + rank)].Font.Bold = true;

                worksheet.Cells[row + rank, 14] = user.LastInitial();
                worksheet.Cells[row + rank, 15] = user.Over60Percentage();
                rank++;
            }

            worksheet = FormatLineupTable(worksheet, row, rankCount);
            newRow = row + rankCount;
            return worksheet;
        }

        private Worksheet FormatLineupTable(Worksheet worksheet, int row, int rankCount)
        {
            // format the rank table
            worksheet.Range["A" + row, "P" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // make the name columns align right
            worksheet.Range["D" + (row - 1), "D" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["H" + (row - 1), "H" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["J" + (row - 1), "J" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["L" + (row - 1), "L" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["N" + (row - 1), "N" + (row + rankCount)].HorizontalAlignment = XlHAlign.xlHAlignRight;


            // iterate through rows starting at row
            for (int i = 0; i < rankCount; i++)
            {
                // Ranking numbers
                worksheet.Cells[(i + row), 1] = i + 1;

                // alternate row colors with smoke
                if (i % 2 == 0)
                {
                    worksheet.Range["A" + (i + row), "P" + (i + row)].Interior.Color = XlRgbColor.rgbWhiteSmoke;
                }
            }

            return worksheet;
        }


        ////// Teams Tables
        private Worksheet AddEachTeamsMetrics(List<RepData> reps, Worksheet worksheet, int row)
        {

            foreach (var team in Settings.Teams)
            {
                var teamRowStart = row;

                if (team.Value.Count == 0)
                    continue;

                if (Settings.IgnoreTeamMetrics.Contains(team.Key)) continue; // skip ignored teams

                // add team name as header
                worksheet.Cells[row, 1] = team.Key + " Team Metrics";

                // format header row
                worksheet.Range[$"A{row}", $"P{row}"].Font.Bold = true;
                 worksheet.Range[$"A{row}", $"B{row}"].Interior.Color = (int)XlRgbColor.rgbLightSteelBlue;
                worksheet.Range["A" + row, "B" + row].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                row++;

                // add the user header
                worksheet = AddUserHeader(worksheet, row);
                row++;

                // get the reps for this team
                List<RepData> teamReps = new();
                foreach (var repName in team.Value)
                {
                    var rep = reps.FirstOrDefault(r => r.Name == repName);
                    if (rep != null)
                    {
                        teamReps.Add(rep);
                    }
                }

                // create average and total reps for this team
                var teamAverageRep = CreateAverageUser(teamReps);
                var teamTotalRep = CreateTotalUser(teamReps);

                // sort then add total then average rep to the front
                teamReps = teamReps.OrderByDescending(u => u.TotalCalls).ToList();
                teamReps.Insert(0, teamTotalRep);
                teamReps.Insert(1, teamAverageRep);
                List<string> ignoreList = new();

                teamReps.RemoveAll(user => ignoreList.Any(ignored => ignored.Trim() == user.Name));
                int userCtr = 0;
                foreach (var user in teamReps)
                {
                    worksheet = AddUserMetrics(user, worksheet, row);

                    // reference B{row} & C{row} from second table for ticket counts
                    if (TicketsFormulaSourceRows.ContainsKey(user.Name))
                    {
                        int sourceRow = TicketsFormulaSourceRows[user.Name];
                        worksheet.Cells[row, 2].Formula = "=B" + sourceRow;
                        worksheet.Cells[row, 3].Formula = "=C" + sourceRow;
                    }

                    userCtr++;
                    row++;
                }
                row++;

                // override average and total user formula for this team only.
                // setup Total and Average users Ticket Formula
                int tRow = teamRowStart + 2; // move row down past headers
                worksheet.Cells[tRow, 2].Formula = "=SUM(B" + (tRow + 2) + ":B" + (tRow + team.Value.Count + 1) + ")";
                worksheet.Cells[tRow, 3].Formula = "=SUM(C" + (tRow + 2) + ":C" + (tRow + team.Value.Count + 1) + ")";

                int aRow = tRow + 1; //=IF(SUM(C10:C27) <= 0,0,AVERAGE(C10:C27))
                worksheet.Cells[aRow, 2].Formula = "=IF(SUM(B" + (aRow + 1) + ":B" + (aRow + team.Value.Count) + ")<=0,0,ROUND(AVERAGE(B" + (aRow + 1) + ":B" + (aRow + team.Value.Count) + "),2))";
                worksheet.Cells[aRow, 3].Formula = "=IF(SUM(C" + (aRow + 1) + ":C" + (aRow + team.Value.Count) + ")<=0,0,ROUND(AVERAGE(C" + (aRow + 1) + ":C" + (aRow + team.Value.Count) + "),2))";

            }
            return worksheet;
        }



        ///////// formatting and helpers
        private Worksheet ApplyBoldIfAverage(Worksheet sheet)
        {
            // add a condition to the entire sheet to make any cell with -- AVERAGE -- bold aswell the cell to its right

            // idk dude this is hard

            return sheet;
        }

        private RepData CreateTotalUser(List<RepData> reps)
        {
            RepData totalUser = new RepData();

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
                // add total calls
                totalCalls += rep.TotalCalls;
                inboundCalls += rep.InboundCalls;
                outboundCalls += rep.OutboundCalls;

                // add timing
                totalPhoneTime += rep.TotalDuration;
                inboundPhoneTime += rep.InboundDuration;
                outboundPhoneTime += rep.OutboundDuration;

                InternalCalls += rep.InternalCalls;
                WeekendCalls += rep.WeekendCalls;


                foreach (var call in rep.Calls)
                {
                    if (call.Duration > 1800) callsOver30++;
                    if (call.Duration > 3600) callsOver60++;
                }
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

        private RepData CreateAverageUser(List<RepData> reps)
        {
            List<string> ignoreList = new();
            //List<string> ignoreList = ConfigurationManager.AppSettings["ignore_users"].Split(',').ToList();

            // create new rep
            RepData avgUser = new RepData();
            avgUser.Name = "-- AVERAGE --";

            // get the average metrics
            int userCount = reps.Count;

            if (userCount == 0)
            {
                Console.WriteLine("No reps found");
                return avgUser;
            }

            int totalCalls = 0;
            int weekendCalls = 0;
            int internalCalls = 0;
            int adjustedCalls = 0;

            int totalPhoneTime = 0;
            int callsOver30 = 0;
            int callsOver60 = 0;

            // add from all reps
            foreach (var user in reps)
            {
                if (user.Name == "-- TOTAL --") continue; // skip the total rep
                if (ignoreList.Any(ignored => ignored.Trim() == user.Name)) continue; // skip the ignored reps

                // add total calls
                totalCalls += user.Calls.Count;
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
            worksheet.Range["A1", "P" + (row - 1)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // set the daterange column aligh left
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            // make all columns wider
            worksheet.Range["A1", "P1"].ColumnWidth = 14;

            // make the notes column wide enough for the rep to make notes
            worksheet.Range["P1", "P" + (row - 1)].ColumnWidth = 70;

            // left alight the whole sheet
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["A"]).HorizontalAlignment = XlHAlign.xlHAlignLeft;

            return worksheet;
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
