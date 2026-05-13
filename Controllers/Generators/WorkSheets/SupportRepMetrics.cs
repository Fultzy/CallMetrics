using CallMetrics.Controllers.Readers.Nextiva;
using CallMetrics.Data;
using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System;
using System.Collections.Generic;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Media;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CallMetrics.Controllers.Generators.WorkSheets
{

    public class SupportRepMetrics : IDisposable
    {
        public Rep AverageRep = null;
        public Rep TotalRep = null;

        public Dictionary<string, int> TicketsFormulaSourceRows = new();
        public int rankCount = Settings.RankedRepsCount;

        private double percentageStep = 0;
        private double currentProgress = 0;

        public event EventHandler<int> ProgressChanged;

        // as repName, generalTableRow
        private Dictionary<string, int> GeneralTableReference;
        private int GeneralTableStart = 0;
        private int GeneralTableEnd = 0;
        private int WorkdaysRow = 0;
        private int WorkdaysCol = 2;

        // as teamName, teamTableRow
        private Dictionary<string, int> TeamsTableReference;

        // headers as ColumnName, ColumnNumber
        public static Dictionary<string, int> UsrHeader = new()
        {
            { "Name", 1 },
            { "Tickets", 2 },
            { "Wkd Tickets", 3 },
            { "Adj Tickets", 4 },
            { "Tickets/Day", 5 },
            { "Calls", 6 },
            { "Wkd Calls", 7 },
            { "Inbound Calls", 8 },
            { "Outbound Calls", 9 },
            { "Adj Calls", 10 },
            { "Calls/Day", 11 },
            { "Calls/Tickets", 12 },
            { "Avg Call Time", 13 },
            { "Total Ph Time", 14 },
            { "Calls > 30m", 15 },
            { "> 30%", 16 },
            { "Calls > 1h", 17 },
            { "> 60%", 18 },
            { "Absences", 19 },
            { "Notes", 20 }, // skip column 18 for formatting
        };

        public static Dictionary<string, int> AvgHeader = new()
        {
            { "Adj Tickets", 2 },
            { "Tickets/Day", 4 },
            { "Adj Calls", 6 },
            { "Calls/Day", 8 },
            { "Calls/Tickets", 10 },
            { "Avg Call Time", 12 },
            { "Total Ph Time", 14 },
            { "> 30%", 16 },
            { "> 60%", 18 },
        };

        public static Dictionary<string, int> AvgTypeHeader = new()
        {
            { "Calls", 2 },
            { "Calls/Day", 4 },
            { "Avg Call Time", 6 },
            { "Total Ph Time", 8 },
            { "Calls > 30m", 10 },
            { "> 30%", 12 },
            { "Calls > 1h", 14 },
            { "> 60%", 16 },  
        };

        public static Dictionary<string, int> TeamRefTable = new()
        {
            { "Name", 1 },
            { "Adj Tickets", 2 },
            { "Tickets/Day", 3 },
            { "Inbound Calls", 4 },
            { "Inbound Calls/Day", 5 },
            { "Outbound Calls", 6 },
            { "Outbound Calls/Day", 7 },
            { "Calls", 8 },
            { "Calls/Day", 9 },
            { "Avg Call Time", 10 },
            { "Absences", 12 },
        };

        private List<Rep> Departments;
        private List<Team> Teams;
        private List<Rep> Reps;
        private List<Rep> ExcludeReps;
        private List<Rep> IncludedReps;
        private List<Rep> WeekendReps;

        public void Dispose()
        {
        }

        public Worksheet Create(List<Rep> importReps, Worksheet worksheet)
        {
            /////////////////////////////////////////// WorkSheet Setup
            // setup worksheet
            worksheet.Name = "Support Metric Report";
            Settings.Load();

            // set the row counter
            int row = 1;
            GeneralTableReference = new();
            TeamsTableReference = new();

            ButcketReps(importReps);

            if (Reps.Count == 0)
                Reps = importReps ;

            var repCount = Reps.Count - ExcludeReps.Count;
            if (rankCount > repCount + 1) // add 1 to include the average rep
                rankCount = repCount + 1;


            // calculate progress step
            currentProgress = 0;
            var teamTables = Settings.Teams.Count(t => t.IncludeInMetrics && t.Members.Count > 0);
            percentageStep = SupportMetricsHelper.CalculateProgressSteps(Departments.Count, Reps.Count, rankCount, teamTables);

            // set Date Range
            DateRange dateRange = new DateRange();
            dateRange.StartDate = MetricsData.Calls.Min(c => c.DateTime);
            dateRange.EndDate = MetricsData.Calls.Max(c => c.DateTime);


            /////////////////////////////////////////// Begin MetricsData Entry0.
            /////////// Top References
            worksheet = AddTopReferences(worksheet, dateRange, row, out row);

            /////////// Table 1 - Company Wide Metrics
            worksheet = AddDepartmentMetrics(Departments, worksheet, row, out row);

            /////////// Table 2 - User Metrics
            worksheet = AddGeneralTableMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 3 - Average Lineup
            worksheet = AddAverageLineupMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 4 - Inbound Average Lineup
            worksheet = AddInboundAverageMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 5 - Outbound Average Lineup
            worksheet = AddOutboundAverageMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 6+ - Individual Teams Tables
            worksheet = AddEachTeamsMetrics(Reps, worksheet, row, out row);

            // go back and write top ref team table
            worksheet = AddTeamsOverview(Teams, worksheet, 2, out int noRow, 3);

            // format the worksheet
            worksheet = FormatWorksheet(worksheet, row, rankCount);

            Console.WriteLine(" Done!");
            return worksheet;
        }

        private void ButcketReps(List<Rep> importReps)
        {
            Departments = importReps.Where(r => Settings.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IsDepartment == true)).ToList();

            Teams = Settings.Teams.Where(t => t.IncludeInMetrics && t.Members.Count > 0).ToList();

            Reps = importReps.Where(r => Settings.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IncludeInMetrics == true)).ToList();

            ExcludeReps = Reps.Where(r => Settings.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IsExcluded == true)).ToList();

            IncludedReps = importReps.Where(r => Settings.Teams.Any(t => t.Members.Contains(r.Name) && t.IncludeInMetrics && !t.IsExcluded)).ToList();

            WeekendReps = Reps.Where(r => r.WeekendCalls > 0 || r.WeekendTickets > 0).ToList();
        }



        private Worksheet AddTopReferences(Worksheet worksheet, DateRange dateRange, int row, out int newRow)
        {
            // Setup top row
            worksheet.Range["A1", "T1"].Font.Size = 12;
            worksheet.Range["A1", "T1"].Interior.Color = XlRgbColor.rgbLightSteelBlue;

            // add date range to the worksheet
            worksheet.Cells[row, 1] = "Date Range:";
            worksheet.Range["B" + 1, "C" + 1].Merge();
            worksheet.Cells[row, 2] = dateRange.ToString();

            // Format DateRange
            worksheet.Range["B1", "B2"].Font.Bold = true;
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["B2"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            row += 1;

            // set Rep Counts
            worksheet.Cells[row, 1] = "Rep Counts";
            worksheet.Cells[row, 1].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Cells[row, 1].Font.Bold = true;
            foreach (var team in Teams)
            {
                row += 1;
                worksheet.Cells[row, 1] = team.Name;
                worksheet.Cells[row, 2] = team.Members.Count;
            }

            row += 1;
            worksheet.Cells[row, 1] = "Total Reps";
            worksheet.Cells[row, 2] = Reps.Count;

            // add weekend rep count
            row += 2;
            worksheet.Cells[row, 1] = "Weekend Reps";
            worksheet.Cells[row, 2] = WeekendReps.Count;

            // add Work Days
            row += 1;
            worksheet.Cells[row, 1] = "Work Days";
            worksheet.Cells[row, 2] = dateRange.WorkDays();
            WorkdaysRow = row;


            newRow = row > Teams.Count * 2 ? row + 2 : Teams.Count * 2 + 2;
            return worksheet;
        }


        ///////// FIRST TABLE
        private Worksheet AddDepartmentMetrics(List<Rep> departments, Worksheet worksheet, int row, out int newRow)
        {
            worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, row, "Departments", UsrHeader);
            row += 2;

            // sort departments by total calls descending
            departments = departments.OrderByDescending(u => u.TotalCalls).ToList();

            // total user
            var totalUser = SupportMetricsHelper.CreateTotalUser(Reps);
            departments.Insert(0, totalUser);

            // add support department
            var supportUser = SupportMetricsHelper.CreateTotalUser(IncludedReps);
            supportUser.Name = "Support";
            departments.Insert(1, supportUser);

            foreach (var dept in departments)
            {
                worksheet = WriteUserMetrics(dept, worksheet, row);
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                row++;
            }

            newRow = row + Departments.Count + 1;
            return worksheet;
        }


        ///////// SECOND TABLE
        private Worksheet AddGeneralTableMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            // add absences note above header
            worksheet.Cells[row, UsrHeader["Absences"]] = "Edit Absences Here";
            worksheet.Cells[row, UsrHeader["Absences"]].Font.Italic = true;
            worksheet.Cells[row, UsrHeader["Absences"]].Font.Size = 10;

            worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, row, "Metrics", UsrHeader);
            
            row += 2;

            worksheet = WriteAverageUser(worksheet, row, reps);
            GeneralTableStart = row;
            row++;


            int userCtr = 0;
            foreach (var rep in reps)
            {

                worksheet = WriteUserMetrics(rep, worksheet, row);
                GeneralTableReference.Add(rep.Name, row);

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                userCtr++;
                row++;
            }

            GeneralTableEnd = row - 1;
            
            newRow = row + 1;
            return worksheet;
        }

        private Worksheet WriteUserMetrics(Rep rep, Worksheet worksheet, int row)
        {
            if (rep == null)
                return worksheet;

            var team = Settings.Teams.FirstOrDefault(t => t.Members.Contains(rep.Name));

            try
            {
                // add the record to the worksheet
                if (team.IsDepartment)
                    worksheet.Cells[row, UsrHeader["Name"]] = rep.Name;
                else
                    worksheet.Cells[row, UsrHeader["Name"]] = Formatter.LastInitial(rep.Name);

                worksheet.Cells[row, UsrHeader["Tickets"]] = rep.TotalTickets;
                worksheet.Cells[row, UsrHeader["Wkd Tickets"]] = rep.WeekendTickets;
                worksheet.Cells[row, UsrHeader["Adj Tickets"]] = rep.AdjustedTickets();

                worksheet.Cells[row, UsrHeader["Calls"]] = rep.TotalCalls;
                worksheet.Cells[row, UsrHeader["Wkd Calls"]] = rep.WeekendCalls;
                worksheet.Cells[row, UsrHeader["Inbound Calls"]] = rep.InboundCalls;
                worksheet.Cells[row, UsrHeader["Outbound Calls"]] = rep.OutboundCalls;
                worksheet.Cells[row, UsrHeader["Adj Calls"]] = rep.AdjustedCalls();

                worksheet.Cells[row, UsrHeader["Calls/Tickets"]] = rep.CallsToTicketsRatio();
                worksheet.Cells[row, UsrHeader["Avg Call Time"]] = Formatter.Duration(Convert.ToInt32(rep.AverageDuration()));
                worksheet.Cells[row, UsrHeader["Total Ph Time"]] = Formatter.Duration(rep.TotalDuration);

                worksheet.Cells[row, UsrHeader["Calls > 30m"]] = rep.CallsOver30;
                worksheet.Cells[row, UsrHeader["> 30%"]] = rep.Over30Percentage() + "%";
                worksheet.Cells[row, UsrHeader["Calls > 1h"]] = rep.CallsOver60;
                worksheet.Cells[row, UsrHeader["> 60%"]] = rep.Over60Percentage() + "%";

                worksheet.Cells[row, UsrHeader["Absences"]] = 0;

                // formula for Absences with ratio Calculations
                // =IF(H9/E1>1,ROUND(H9/E1,2),0)
                // something is overwriting the tickets/days cell - DEBUG
                var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
                var tkCol = SupportMetricsHelper.numToLetter(UsrHeader["Adj Tickets"]);
                var clCol = SupportMetricsHelper.numToLetter(UsrHeader["Adj Calls"]);

                var wdRow = WorkdaysRow;
                var wdCol = SupportMetricsHelper.numToLetter(2);

                worksheet.Cells[row, UsrHeader["Tickets/Day"]].Formula = $"=IF({wdCol}{wdRow}>{abCol}{row},ROUND({tkCol}{row}/({wdCol}{wdRow}-{abCol}{row}),0),0)"; // tickets
                worksheet.Cells[row, UsrHeader["Calls/Day"]].Formula = $"=IF({wdCol}{wdRow}>{abCol}{row},ROUND({clCol}{row}/({wdCol}{wdRow}-{abCol}{row}),0),0)"; // calls

                if (!TicketsFormulaSourceRows.ContainsKey(rep.Name))
                {
                    TicketsFormulaSourceRows.Add(rep.Name, row);
                }

                var clSt = SupportMetricsHelper.numToLetter(1);
                var clEd = SupportMetricsHelper.numToLetter(UsrHeader.Count);

                // add borders to the row
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                // left align the name cell
                worksheet.Range["A" + row].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                // center align rest of row
                worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // merge notes columns
                worksheet.Range[$"{clEd}{row}", $"{SupportMetricsHelper.numToLetter(UsrHeader.Count + 2)}{row}"].Merge();

                // left align notes column
                worksheet.Range[$"{clEd}{row}", $"{SupportMetricsHelper.numToLetter(UsrHeader.Count + 2)}{row}"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                // alternate row colors with smoke
                if (row % 2 == 0)
                {
                    worksheet.Range["A" + row, "T" + row].Interior.Color = XlRgbColor.rgbWhiteSmoke;
                }
            }
            catch (Exception ex)
            {
                var msg = Logger.ExceptionLog("Error adding rep metrics for " + rep.Name + ": " + ex.Message);
                throw new Exception(msg);
            }

            return worksheet;
        }


        ///////// AVERAGES TABLES
        private Worksheet AddAverageLineupMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            // this table is going to display data in columns. each column will be a different metric and
            // will be sorted vertically by the metric with that genReps name to the left. 
            // there will be a sudo rep added to the list that will be the average of all the genReps.

            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "General Averages", row, AvgHeader);
            row += 2;

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            // setup sorted lists
            List<Rep> adjTickets = reps.OrderByDescending(u => u.AdjustedTickets()).ToList();
            List<Rep> callsToTickets = reps.OrderByDescending(u => u.CallsToTicketsRatio()).ToList();

            List<Rep> adjCalls = reps.OrderByDescending(u => u.AdjustedCalls()).ToList();
            List<Rep> avgCallTime = reps.OrderBy(u => u.AverageDuration()).ToList();
            List<Rep> totalPhoneTime = reps.OrderByDescending(u => u.TotalDuration).ToList();

            List<Rep> callsOver30Percent = reps.OrderBy(u => u.Over30Percentage()).ToList();
            List<Rep> callsOver60Percent = reps.OrderBy(u => u.Over60Percentage()).ToList();

            int startRow = GeneralTableStart;
            int endRow = GeneralTableEnd;
            var ajTCol = SupportMetricsHelper.numToLetter(UsrHeader["Adj Tickets"]);
            var ajCCol = SupportMetricsHelper.numToLetter(UsrHeader["Adj Calls"]);
            var TpDCol = SupportMetricsHelper.numToLetter(UsrHeader["Tickets/Day"]);
            var CpDCol = SupportMetricsHelper.numToLetter(UsrHeader["Calls/Day"]);
            int rank = 0;


            foreach (Rep rep in reps.Take(rankCount))
            {
                //// Value =LARGE(D9:D23, 10)
                //worksheet.Cells[row + rank, 3].Formula =
                //    $"=LARGE({ajTCol}{startRow}:{ajTCol}{endRow}, A{row + rank})";

                //// Value =SMALL(I9:I23, 10)
                //worksheet.Cells[row + rank, 7].Formula =
                //    $"=SMALL({ajCCol}{startRow}:{ajCCol}{endRow}, A{row + rank})";

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Adjusted Tickets
            rank = 0;
            foreach (Rep user in adjTickets.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["Adj Tickets"],
                    Formatter.LastInitial(user.Name),
                    user.AdjustedTickets(),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Tickets/Day
            rank = 0;
            worksheet.Cells[row + rank, AvgHeader["Tickets/Day"]].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, {TpDCol}{startRow}:{TpDCol}{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            foreach (Rep user in reps.Take(rankCount))
            {
                // Tickets/Day
                // NAME =INDEX(SORTBY(A9:A28, Q9:Q28, -1), SEQUENCE(20,1,1,1))

                worksheet.Cells[row + rank, AvgHeader["Tickets/Day"] + 1].Formula =
                   $"=LARGE({TpDCol}{startRow}:{TpDCol}{endRow}, A{row + rank})";

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Adjusted Calls
            rank = 0;
            foreach (Rep user in adjCalls.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["Adj Calls"],
                    Formatter.LastInitial(user.Name),
                    user.AdjustedCalls(),
                    rank
                );


                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Calls/Day
            rank = 0;
            worksheet.Cells[row + rank, AvgHeader["Calls/Day"]].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, {CpDCol}{startRow}:{CpDCol}{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            foreach (Rep user in reps.Take(rankCount))
            {
                // Calls/Day
                //worksheet.Cells[row + rank, 18].Formula =
                //$"=INDEX(A{9}:A{9 + genReps.Count - 1}, MATCH(LARGE(R{9}:R{9 + genReps.Count - 1}, A{row + rank}), R{9}:R{9 + genReps.Count - 1}, 0))";


                worksheet.Cells[row + rank, AvgHeader["Calls/Day"] + 1].Formula =
                   $"=LARGE({CpDCol}{startRow}:{CpDCol}{endRow}, A{row + rank})";

                //worksheet.Cells[row + rank, 18].Formula = $"=IF(H{9 + rank}/(F1-P{9 + rank})>1,ROUND(H{9 + rank}/(F1-Q{9 + rank}),0),0)";
                //worksheet.Cells[row + rank, 19].Formula = $"=IF(H{9 + rank}/(F1-P{9 + rank})>1,ROUND(H{9 + rank}/(F1-R{9 + rank}),0),0)";

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Calls/Tickets
            rank = 0;
            foreach (Rep user in callsToTickets.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["Calls/Tickets"],
                    Formatter.LastInitial(user.Name),
                    user.CallsToTicketsRatio(),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Avg Call Time
            rank = 0;
            foreach (Rep user in avgCallTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["Avg Call Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(Convert.ToInt32(user.AverageDuration())),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Total Phone Time
            rank = 0;
            foreach (Rep user in totalPhoneTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["Total Ph Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(user.TotalDuration),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Calls > 30m %
            rank = 0;
            foreach (Rep user in callsOver30Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["> 30%"],
                    Formatter.LastInitial(user.Name),
                    user.Over30Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            // Calls > 60m %
            rank = 0;
            foreach (Rep user in callsOver60Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgHeader["> 60%"],
                    Formatter.LastInitial(user.Name),
                    user.Over60Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            worksheet = FormatLineupTable(worksheet, row, rankCount);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;
        }

        private Worksheet AddInboundAverageMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "Inbound Averages", row, AvgTypeHeader, "Inbound");
            row += 2;

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            // setup sorted lists
            List<Rep> calls = reps.OrderByDescending(u => u.InboundCalls).ToList();
            List<Rep> avgCallTime = reps.OrderBy(u => u.AverageInboundDuration()).ToList();
            List<Rep> totalPhoneTime = reps.OrderByDescending(u => u.InboundDuration).ToList();

            List<Rep> inboundOver30 = reps.OrderBy(u => u.InboundCallsOver30).ToList();
            List<Rep> inboundOver60 = reps.OrderBy(u => u.InboundCallsOver60).ToList();

            List<Rep> callsOver30Percent = reps.OrderBy(u => u.InboundOver30Percentage()).ToList();
            List<Rep> callsOver60Percent = reps.OrderBy(u => u.InboundOver60Percentage()).ToList();

            int startRow = GeneralTableStart;
            int endRow = GeneralTableEnd;
            var iCCol = SupportMetricsHelper.numToLetter(UsrHeader["Inbound Calls"]);
            var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
            //var CpDCol = SupportMetricsHelper.numToLetter(UHeader["Calls/Day"]);
            var wdCol = SupportMetricsHelper.numToLetter(2);
            int rank = 0;

            // write
            // Total Inbound Calls
            rank = 0;
            foreach (Rep user in calls.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls"],
                    Formatter.LastInitial(user.Name),
                    user.InboundCalls,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            worksheet.Cells[row, AvgTypeHeader["Calls/Day"]].Formula2 =
                $"=INDEX(SORTBY(CHOOSE(" + "{1,2}" + $",A{startRow}:A{endRow},FLOOR(IFERROR({iCCol}{startRow}:{iCCol}{endRow}/({wdCol}{WorkdaysRow}-{abCol}{startRow}:{abCol}{endRow}),0),1)),FLOOR(IFERROR({iCCol}{startRow}:{iCCol}{endRow}/({wdCol}{WorkdaysRow}-{abCol}{startRow}:{abCol}{endRow}),0),1),-1),SEQUENCE({rankCount})," + "{1,2})";

            //worksheet.Cells[row, AvgTypeHeader["Calls/Day"]].Formula = $"=LET(names,A{startRow}:A{endRow},calls,{clCol}{startRow}:{clCol}{endRow},absences,{abCol}{startRow}:{abCol}{endRow},workdays,$F$1,metric,IFERROR(calls/(workdays-absences),0),sorted,SORTBY(HSTACK(names,metric),metric,-1),TAKE(sorted,100))";

            rank = 0;
            //worksheet.Cells[row + rank, AvgTypeHeader["Calls/Day"]].Formula2 = $"=INDEX(SORTBY(A{startRow}:A{endRow}, {CpDCol}{startRow}:{CpDCol}{endRow}, -1), SEQUENCE({rankCount},1,1,1))";
            //foreach (Rep user in reps.Take(rankCount))
            //{
            //    worksheet.Cells[row + rank, AvgTypeHeader["Calls/Day"] + 1].Formula =
            //       $"=LARGE({CpDCol}{startRow}:{CpDCol}{endRow}, A{row + rank})";

            //    ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
            //    rank++;
            //}

            // Avg Call Time
            rank = 0;
            foreach (Rep user in avgCallTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Avg Call Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(Convert.ToInt32(user.AverageInboundDuration())),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Total Phone Time
            rank = 0;
            foreach (Rep user in totalPhoneTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Total Ph Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(user.InboundDuration),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 30m
            rank = 0;
            foreach (Rep user in inboundOver30.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls > 30m"],
                    Formatter.LastInitial(user.Name),
                    user.InboundCallsOver30,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 30m %
            rank = 0;
            foreach (Rep user in callsOver30Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["> 30%"],
                    Formatter.LastInitial(user.Name),
                    user.InboundOver30Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 60m
            rank = 0;
            foreach (Rep user in inboundOver60.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls > 1h"],
                    Formatter.LastInitial(user.Name),
                    user.InboundCallsOver60,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 60m %
            rank = 0;
            foreach (Rep user in callsOver60Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["> 60%"],
                    Formatter.LastInitial(user.Name),
                    user.InboundOver60Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            worksheet = FormatLineupTable(worksheet, row, rankCount);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;

        }

        private Worksheet AddOutboundAverageMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "Outbound Averages", row, AvgTypeHeader, "Outbound");
            row += 2;

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            // setup sorted lists
            List<Rep> calls = reps.OrderByDescending(u => u.OutboundCalls).ToList();
            List<Rep> avgCallTime = reps.OrderBy(u => u.AverageOutboundDuration()).ToList();
            List<Rep> totalPhoneTime = reps.OrderByDescending(u => u.OutboundDuration).ToList();

            List<Rep> outboundOver30 = reps.OrderBy(u => u.OutboundCallsOver30).ToList();
            List<Rep> outboundOver60 = reps.OrderBy(u => u.OutboundCallsOver60).ToList();

            List<Rep> callsOver30Percent = reps.OrderBy(u => u.OutboundOver30Percentage()).ToList();
            List<Rep> callsOver60Percent = reps.OrderBy(u => u.OutboundOver60Percentage()).ToList();

            int startRow = GeneralTableStart; 
            int endRow = GeneralTableEnd;
            var iCCol = SupportMetricsHelper.numToLetter(UsrHeader["Outbound Calls"]);
            var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
            var wdCol = SupportMetricsHelper.numToLetter(2);
            int rank = 0;

            // write
            // Calls/Day
            rank = 0;
            worksheet.Cells[row, AvgTypeHeader["Calls/Day"]].Formula2 =
                $"=INDEX(SORTBY(CHOOSE(" + "{1,2}" + $",A{startRow}:A{endRow},FLOOR(IFERROR({iCCol}{startRow}:{iCCol}{endRow}/({wdCol}{WorkdaysRow}-{abCol}{startRow}:{abCol}{endRow}),0),1)),FLOOR(IFERROR({iCCol}{startRow}:{iCCol}{endRow}/({wdCol}{WorkdaysRow}-{abCol}{startRow}:{abCol}{endRow}),0),1),-1),SEQUENCE({rankCount})," + "{1,2})";

            // Total Outbound Calls
            rank = 0;
            foreach (Rep user in calls.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls"],
                    Formatter.LastInitial(user.Name),
                    user.OutboundCalls,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Avg Call Time
            rank = 0;
            foreach (Rep user in avgCallTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Avg Call Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(Convert.ToInt32(user.AverageOutboundDuration())),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Total Phone Time
            rank = 0;
            foreach (Rep user in totalPhoneTime.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Total Ph Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(user.OutboundDuration),
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 30m
            rank = 0;
            foreach (Rep user in outboundOver30.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls > 30m"],
                    Formatter.LastInitial(user.Name),
                    user.OutboundCallsOver30,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 30m %
            rank = 0;
            foreach (Rep user in callsOver30Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["> 30%"],
                    Formatter.LastInitial(user.Name),
                    user.OutboundOver30Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 60m
            rank = 0;
            foreach (Rep user in outboundOver60.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Calls > 1h"],
                    Formatter.LastInitial(user.Name),
                    user.OutboundCallsOver60,
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }

            // Calls > 60m %
            rank = 0;
            foreach (Rep user in callsOver60Percent.Take(rankCount))
            {
                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["> 60%"],
                    Formatter.LastInitial(user.Name),
                    user.OutboundOver60Percentage() + "%",
                    rank
                );

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                rank++;
            }


            worksheet = FormatLineupTable(worksheet, row, rankCount);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;
        }


        ////// Teams Tables
        private Worksheet AddTeamsOverview(List<Team> teams, Worksheet worksheet, int row, out int newRow, int offset = 0)
        {
            // determin Columns
            var stCol = SupportMetricsHelper.numToLetter(offset + 1);
            var edCol = SupportMetricsHelper.numToLetter(offset + TeamRefTable.Count + 1);
            var wdCol = SupportMetricsHelper.numToLetter(WorkdaysCol);
            var iCol = SupportMetricsHelper.numToLetter(offset + TeamRefTable["Inbound Calls"]);
            var oCol = SupportMetricsHelper.numToLetter(offset + TeamRefTable["Outbound Calls"]);
            var aCol = SupportMetricsHelper.numToLetter(offset + TeamRefTable["Absences"]);


            // Add Header
            worksheet.Cells[row, offset + 1] = "Team";

            worksheet.Cells[row, offset + 2] = "Adj Tickets";
            worksheet.Cells[row, offset + 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[row, offset + 3] = "/Day";
            worksheet.Cells[row, offset + 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            worksheet.Cells[row, offset + 4] = "Inbound Calls";
            worksheet.Cells[row, offset + 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[row, offset + 5] = "/Day";
            worksheet.Cells[row, offset + 5].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            worksheet.Cells[row, offset + 6] = "Outbound Calls";
            worksheet.Cells[row, offset + 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[row, offset + 7] = "/Day";
            worksheet.Cells[row, offset + 7].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            worksheet.Cells[row, offset + 8] = "Adj Calls";
            worksheet.Cells[row, offset + 8].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[row, offset + 9] = "/Day";
            worksheet.Cells[row, offset + 9].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            worksheet.Cells[row, offset + 10] = "Avg Call Time";
            worksheet.Cells[row, offset + 10].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[row, offset + 12] = "Absences";
            worksheet.Cells[row, offset + 12].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Range[$"{stCol}{row}", $"{edCol}{row}"].Font.Bold = true;
            worksheet.Range[$"{stCol}{row}", $"{edCol}{row}"].Interior.Color = XlRgbColor.rgbLightSteelBlue;
            worksheet.Range[$"{stCol}{row}", $"{edCol}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[row, offset + 1].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;

            row++;

            // write content
            foreach (var teamRef in TeamsTableReference)
            {
                worksheet.Cells[row, offset + 1] = teamRef.Key;

                var ttCol = SupportMetricsHelper.numToLetter(UsrHeader["Adj Tickets"]);
                worksheet.Cells[row, offset + TeamRefTable["Adj Tickets"]].Formula = $"={ttCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Adj Tickets"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tpdCol = SupportMetricsHelper.numToLetter(UsrHeader["Tickets/Day"]);
                worksheet.Cells[row, offset + TeamRefTable["Tickets/Day"]].Formula = $"={tpdCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Tickets/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                var icCol = SupportMetricsHelper.numToLetter(UsrHeader["Inbound Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls"]].Formula = $"={icCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var ocCol = SupportMetricsHelper.numToLetter(UsrHeader["Outbound Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls"]].Formula = $"={ocCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tcCol = SupportMetricsHelper.numToLetter(UsrHeader["Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Calls"]].Formula = $"={tcCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tcpdCol = SupportMetricsHelper.numToLetter(UsrHeader["Calls/Day"]);
                worksheet.Cells[row, offset + TeamRefTable["Calls/Day"]].Formula = $"={tcpdCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Calls/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                var actCol = SupportMetricsHelper.numToLetter(UsrHeader["Avg Call Time"]);
                worksheet.Cells[row, offset + TeamRefTable["Avg Call Time"]].Formula = $"={actCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Avg Call Time"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
                worksheet.Cells[row, offset + TeamRefTable["Absences"]].Formula = $"={abCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Absences"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;


                // =FLOOR(IFERROR(G3/(B9-N3),0),1)
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls/Day"]].Formula = 
                    $"=FLOOR(IFERROR({iCol}{row}/({wdCol}{WorkdaysRow}-{aCol}{row}),0),1)";
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls/Day"]].Formula = 
                    $"=FLOOR(IFERROR({oCol}{row}/({wdCol}{WorkdaysRow}-{aCol}{row}),0),1)";
                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;


                // add bottom border if avg in name
                if (teamRef.Key.Contains("Avg"))
                {
                    
                    worksheet.Range[$"{stCol}{row}", $"{edCol}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }

                // always border left
                worksheet.Cells[row, offset + 1].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;

                // alternate row colors with smoke
                if (row % 2 == 0)
                {
                    worksheet.Range[stCol + (row), edCol + (row)].Interior.Color = XlRgbColor.rgbWhiteSmoke;
                }

                row ++;
            }

            newRow = row;
            return worksheet;
        }

        private Worksheet AddEachTeamsMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            foreach (var team in Settings.Teams)
            {
                if (team.Members.Count == 0)
                    continue;

                // skip team if not enabled for teams table
                if (team.IncludeInMetrics == false)
                    continue;

                // note not included if not included
                worksheet.Cells[row, 2] = team.IsExcluded ? "Is Excluded From Metrics" : "";
                worksheet.Cells[row, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[row, 2].Font.Italic = true;
                worksheet.Cells[row, 2].Font.Size = 11;
               
                // add the rep header
                worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, row, team.Name, UsrHeader);
                row+=2;

                // get the genReps for this team
                List<Rep> teamReps = new();
                foreach (var repName in team.Members)
                {
                    var rep = reps.FirstOrDefault(r => r.Name == repName);
                    if (rep != null)
                    {
                        teamReps.Add(rep);
                    }
                    else
                    {
                        var newRep = new Rep();
                        newRep.Name = repName;
                        teamReps.Add(newRep);
                    }
                }

                worksheet = WriteTotalUser(worksheet, row, teamReps);
                if (!team.IsExcluded) TeamsTableReference.Add("Total " + team.Name, row);
                row++;

                worksheet = WriteAverageUser(worksheet, row, teamReps);
                if (!team.IsExcluded) TeamsTableReference.Add("Avg " + team.Name, row);
                row++;

                foreach (var rep in teamReps)
                {
                    worksheet = WriteUserMetrics(rep, worksheet, row);

                    // add absence cell reference back to general table
                    if (GeneralTableReference.ContainsKey(rep.Name))
                    {
                        var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
                        worksheet.Cells[row, UsrHeader["Absences"]].Formula = $"={abCol}{GeneralTableReference[rep.Name]}";
                    }

                    row++;
                }
                row++;

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
            }

            newRow = row;
            return worksheet;
        }


        ///////// formatting
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

        private Worksheet FormatWorksheet(Worksheet worksheet, int row, int rankCount)
        {
            // format columns to fit
            worksheet.Columns.AutoFit();

            // center align the entire sheet
            //worksheet.Range["B1", "T" + (row)].HorizontalAlignment = XlHAlign.xlHAlignCenter;

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

        private Worksheet WriteTotalUser(Worksheet worksheet, int row, List<Rep> reps)
        {
            var totalUser = SupportMetricsHelper.CreateTotalUser(reps);

            worksheet = WriteUserMetrics(totalUser, worksheet, row);

            int stRow = row + 2;
            int edRow = row + reps.Count + 1;
            var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
            var stCol = SupportMetricsHelper.numToLetter(1);
            var edCol = SupportMetricsHelper.numToLetter(UsrHeader.Count);

            // =SUM(P10:P28)
            worksheet.Cells[row, UsrHeader["Absences"]].Formula = $"=SUM({abCol}{stRow}:{abCol}{edRow})";

            // format row
            worksheet.Range[stCol + row, edCol + row].Font.Bold = true;

            return worksheet;
        }

        private Worksheet WriteAverageUser(Worksheet worksheet, int row, List<Rep> reps)
        {
            var avgUser = SupportMetricsHelper.CreateAverageUser(reps);

            worksheet = WriteUserMetrics(avgUser, worksheet, row);

            int stRow = row + 1;
            int edRow = row + reps.Count;
            var abCol = SupportMetricsHelper.numToLetter(UsrHeader["Absences"]);
            var stCol = SupportMetricsHelper.numToLetter(1);
            var edCol = SupportMetricsHelper.numToLetter(UsrHeader.Count);

            // =IF(SUM(P10:P28)<=0,0,ROUND(AVERAGE(P10:P28),2))
            worksheet.Cells[row, UsrHeader["Absences"]].Formula = $"=IF(SUM({abCol}{stRow}:{abCol}{edRow})<=0,0,ROUND(AVERAGE({abCol}{stRow}:{abCol}{edRow}),2))";

            // format row
            worksheet.Range[stCol + row, edCol + row].Font.Bold = true;

            return worksheet;
        }
    }
}
