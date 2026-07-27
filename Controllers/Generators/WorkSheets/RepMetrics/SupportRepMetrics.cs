using CallMetrics.Controllers.Generators.WorkSheets.Helpers;
using CallMetrics.Data;
using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;


namespace CallMetrics.Controllers.Generators.WorkSheets.RepMetrics
{

    public class SupportRepMetrics : IDisposable
    {
        public Rep AverageRep = new();
        public Rep TotalRep = new();

        public Dictionary<string, int> TicketsFormulaSourceRows = new();
        public int rankCount = Settings.Data.RankedRepsCount;

        private double percentageStep = 0;
        private double currentProgress = 0;

        public event EventHandler<int> ProgressChanged;

        // as repName, generalTableRow
        private Dictionary<string, int> GeneralTableReference = new();
        private int GeneralTableStart = 0;
        private int GeneralTableEnd = 0;
        private int WorkdaysRow = 0;
        private int WorkdaysCol = 2;

        // as teamName, teamTableRow
        private Dictionary<string, int> TeamsTableReference = new();

        // headers as ColumnName, ColumnNumber
        public static Dictionary<string, int> RepHeader = new()
        {
            { "Name", 1 },
            { "Tickets", 2 },
            { "Open Tickets", 3 },
            { "Wkd Tickets", 4 },
            { "Adj Tickets", 5 },
            { "Tickets/Day", 6 },
            { "Calls", 7 },
            { "Wkd Calls", 8 },
            { "Inbound Calls", 9 },
            { "Outbound Calls", 10 },
            { "Adj Calls", 11 },
            { "Calls/Day", 12 },
            { "Calls/Tickets", 13 },
            { "Avg Call Time", 14 },
            { "Total Ph Time", 15 },
            { ">30min Calls", 16 },
            { ">1h Calls", 17 },
            { "Average Ticket Resolution", 18 },
            { "Same Day Resolution", 19 },
            { "Absences", 20 },
            { "Notes", 21 }, // skip column for formatting
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
            ProgressChanged = null;
            GC.SuppressFinalize(this); // ????
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

            BucketReps(importReps);

            if (Reps.Count == 0)
                Reps = importReps ;

            var repCount = Reps.Count - ExcludeReps.Count;
            if (rankCount > repCount + 1) // add 1 to include the average rep
                rankCount = repCount + 1;


            // calculate progress step
            currentProgress = 0;
            var teamTables = Settings.Data.Teams.Count(t => t.IncludeInMetrics && t.Members.Count > 0);
            percentageStep = SupportMetricsHelper.CalculateProgressSteps(Departments.Count, Reps.Count, rankCount, teamTables);

            // set Date Range
            DateRange dateRange = new DateRange();
            dateRange.StartDate = MetricsData.Calls.Min(c => c.DateTime);
            dateRange.EndDate = MetricsData.Calls.Max(c => c.DateTime);

            // format the worksheet
            worksheet = FormatHelper.FormatWorksheet(worksheet, row);

            /////////////////////////////////////////// Begin MetricsData Entry0.
            /////////// Top References
            worksheet = AddTopReferences(worksheet, dateRange, row, out row);

            /////////// Table 1 - Company Wide Metrics
            worksheet = AddDepartmentMetrics(Departments, worksheet, row, out row);

            /////////// Table 2 - User Metrics
            worksheet = AddGeneralTableMetrics(IncludedReps, worksheet, row, out row);
            row++;

            // Predict teams overview row
            var overviewRow = row;
            row = row + Teams.Count * 3 + 3;
            
            /////////// Table 6+ - Individual Teams Tables
            worksheet = AddEachTeamsMetrics(Reps, worksheet, row, out row);
            row++;

            // go back and write top ref team table
            worksheet = AddTeamsOverview(Teams, worksheet, overviewRow, out int norow);

            /////////// Table 3 - Average Lineup
            worksheet = AddAverageLineupMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 4 - Inbound Average Lineup
            worksheet = AddInboundAverageMetrics(IncludedReps, worksheet, row, out row);

            /////////// Table 5 - Outbound Average Lineup
            worksheet = AddOutboundAverageMetrics(IncludedReps, worksheet, row, out row);
            row++;

            Console.WriteLine(" Done!");
            return worksheet;
        }

        private void BucketReps(List<Rep> importReps)
        {
            Departments = importReps.Where(r => Settings.Data.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IsDepartment == true)).ToList();

            Teams = Settings.Data.Teams.Where(t => t.IncludeInMetrics && t.Members.Count > 0).ToList();

            Reps = importReps.Where(r => Settings.Data.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IncludeInMetrics == true)).ToList();

            ExcludeReps = Reps.Where(r => Settings.Data.Teams.Any(t =>
                t.Members.Contains(r.Name) &&
                t.IsExcluded == true)).ToList();

            IncludedReps = importReps.Where(r => Settings.Data.Teams.Any(t => t.Members.Contains(r.Name) && t.IncludeInMetrics && !t.IsExcluded)).ToList();

            WeekendReps = Reps.Where(r => r.WeekendCalls > 0 || r.WeekendTickets > 0).ToList();
        }

        private Worksheet AddTopReferences(Worksheet worksheet, DateRange dateRange, int row, out int newRow)
        {
            // Setup top row
            worksheet = FormatHelper.MainHeader(worksheet, row, "A", "T");

            // add date range to the worksheet
            worksheet.Cells[row, 1] = "Date Range:";
            worksheet.Range["B" + 1, "C" + 1].Merge();
            worksheet.Cells[row, 2] = dateRange.ToString();

            // Format DateRange
            worksheet.Range["A1", "B1"].Font.Bold = true;
            worksheet.Range["A1", "B1"].Font.Size = 14;
            worksheet.Range["B1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //worksheet.Range["B2"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            row += 1;

            // set Rep Counts
            FormatHelper.JoinCells(worksheet, "A2", "B2");
            worksheet.Cells[row, 1] = "Rep Counts";
            worksheet.Cells[row, 1].Interior.Color = FormatHelper.AccentOne;
            worksheet.Cells[row, 1].Font.Color = FormatHelper.FontColor;
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
            worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, "Departments", row, RepHeader, out row);

            // sort departments by total calls descending
            departments = departments.OrderByDescending(u => u.TotalCalls).ToList();
            
            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count));
            row++;

            foreach (var dept in departments)
            {
                worksheet = WriteRepMetrics(dept, worksheet, row);
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                row++;
            }

            worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count - 1));
            row++;

            worksheet = WriteTotalUser(worksheet, row, departments);
            worksheet.Rows[row].Font.Bold = true;
            row++;

            worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count - 1));
            row++;

            newRow = row+2;
            return worksheet;
        }


        ///////// SECOND TABLE
        private Worksheet AddGeneralTableMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {

            worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, "Metrics", row, RepHeader, out row);

            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count));
            row++;

            // add absences note above header
            worksheet.Cells[row, RepHeader["Absences"]] = "Edit Absences Here";
            worksheet.Cells[row, RepHeader["Absences"]].Font.Italic = true;
            worksheet.Cells[row, RepHeader["Absences"]].Font.Size = 10;

            GeneralTableStart = row;

            worksheet = WriteAverageUser(worksheet, row, reps);
            row++;

            int userCtr = 0;
            foreach (var rep in reps)
            {

                worksheet = WriteRepMetrics(rep, worksheet, row);
                GeneralTableReference.Add(rep.Name, row);

                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
                userCtr++;
                row++;
            }

            GeneralTableEnd = row - 1;

            worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count-1));
            row++;

            worksheet = WriteTotalUser(worksheet, row, reps);
            worksheet.Rows[row].Font.Bold = true;
            row++;

            worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count-1));
            row++;

            
            newRow = row + 1;
            return worksheet;
        }

        private Worksheet WriteRepMetrics(Rep rep, Worksheet worksheet, int row)
        {
            if (rep == null)
                return worksheet;

            var team = Settings.Data.Teams.FirstOrDefault(t => t.Members.Contains(rep.Name)) ?? new Team();

            try
            {
                // add the record to the worksheet
                if (team.IsDepartment)
                    worksheet.Cells[row, RepHeader["Name"]] = rep.Name;
                else
                    worksheet.Cells[row, RepHeader["Name"]] = rep.Name; //Formatter.LastInitial(rep.Name);

                worksheet.Cells[row, RepHeader["Tickets"]] = rep.TotalTickets;
                worksheet.Cells[row, RepHeader["Open Tickets"]] = rep.OpenTickets;
                worksheet.Cells[row, RepHeader["Wkd Tickets"]] = rep.WeekendTickets;
                worksheet.Cells[row, RepHeader["Adj Tickets"]] = rep.AdjustedTickets();
                worksheet.Cells[row, RepHeader["Calls"]] = rep.TotalCalls;
                worksheet.Cells[row, RepHeader["Wkd Calls"]] = rep.WeekendCalls;
                worksheet.Cells[row, RepHeader["Inbound Calls"]] = rep.InboundCalls;
                worksheet.Cells[row, RepHeader["Outbound Calls"]] = rep.OutboundCalls;
                worksheet.Cells[row, RepHeader["Adj Calls"]] = rep.AdjustedCalls();
                worksheet.Cells[row, RepHeader["Calls/Tickets"]] = rep.CallsToTicketsRatio();
                worksheet.Cells[row, RepHeader["Avg Call Time"]] = Formatter.Duration(Convert.ToInt32(rep.AverageDuration()));
                worksheet.Cells[row, RepHeader["Total Ph Time"]] = Formatter.Duration(rep.TotalDuration);
                worksheet.Cells[row, RepHeader[">30min Calls"]] = $"{rep.CallsOver30} ({rep.Over30Percentage()}%)";
                worksheet.Cells[row, RepHeader[">1h Calls"]] = $"{rep.CallsOver60} ({rep.Over60Percentage()}%)";
                worksheet.Cells[row, RepHeader["Average Ticket Resolution"]] = rep.AverageTicketResolutionTimeFormatted();
                worksheet.Cells[row, RepHeader["Same Day Resolution"]] = $"{rep.SameDayResolvedTickets} ({Mather.DoublePercentage(rep.SameDayResolvedTickets, rep.TotalTickets)}%)";
                worksheet.Cells[row, RepHeader["Absences"]] = 0;

                // formula for Absences with ratio Calculations
                // =IF(H9/E1>1,ROUND(H9/E1,2),0)
                // something is overwriting the tickets/days cell - DEBUG
                var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
                var tkCol = SupportMetricsHelper.NumToLetter(RepHeader["Adj Tickets"]);
                var clCol = SupportMetricsHelper.NumToLetter(RepHeader["Adj Calls"]);

                var wdRow = WorkdaysRow;
                var wdCol = SupportMetricsHelper.NumToLetter(2);

                worksheet.Cells[row, RepHeader["Tickets/Day"]].Formula = 
                    $"=IF({wdCol}{wdRow}>{abCol}{row},ROUND({tkCol}{row}/({wdCol}{wdRow}-{abCol}{row}),0),0)"; // tickets

                worksheet.Cells[row, RepHeader["Calls/Day"]].Formula = 
                    $"=IF({wdCol}{wdRow}>{abCol}{row},ROUND({clCol}{row}/({wdCol}{wdRow}-{abCol}{row}),0),0)"; // calls

                if (!TicketsFormulaSourceRows.ContainsKey(rep.Name))
                {
                    TicketsFormulaSourceRows.Add(rep.Name, row);
                }

                var clSt = SupportMetricsHelper.NumToLetter(1);
                var clEd = SupportMetricsHelper.NumToLetter(RepHeader.Count);

                //// add borders to the row
                //worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                //worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                //// left align the name cell
                //worksheet.Range["A" + row].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //// center align rest of row
                //worksheet.Range[$"{clSt}{row}", $"{clEd}{row}"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                //// merge notes columns
                //worksheet.Range[$"{clEd}{row}", $"{SupportMetricsHelper.NumToLetter(RepHeader.Count + 2)}{row}"].Merge();

                //// left align notes column
                //worksheet.Range[$"{clEd}{row}", $"{SupportMetricsHelper.NumToLetter(RepHeader.Count + 2)}{row}"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

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
            // will be sorted vertically by the metric with that Reps name to the left. 

            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "General Averages", row, AvgHeader, out row);

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(AvgHeader.Count * 2 + 1));
            row++;

            worksheet = FormatHelper.LineupTable(worksheet, row, rankCount, AvgHeader.Count * 2, 1);

            // setup sorted lists
            List<Rep> adjTickets = reps.OrderByDescending(u => u.AdjustedTickets()).ToList();
            List<Rep> callsToTickets = reps.OrderBy(u => u.CallsToTicketsRatio()).ToList();
            List<Rep> adjCalls = reps.OrderByDescending(u => u.AdjustedCalls()).ToList();
            List<Rep> avgCallTime = reps.OrderBy(u => u.AverageDuration()).ToList();
            List<Rep> totalPhoneTime = reps.OrderByDescending(u => u.TotalDuration).ToList();
            List<Rep> callsOver30Percent = reps.OrderBy(u => u.Over30Percentage()).ToList();
            List<Rep> callsOver60Percent = reps.OrderBy(u => u.Over60Percentage()).ToList();

            int startRow = GeneralTableStart;
            int endRow = GeneralTableEnd;
            var ajTCol = SupportMetricsHelper.NumToLetter(RepHeader["Adj Tickets"]);
            var ajCCol = SupportMetricsHelper.NumToLetter(RepHeader["Adj Calls"]);
            var TpDCol = SupportMetricsHelper.NumToLetter(RepHeader["Tickets/Day"]);
            var CpDCol = SupportMetricsHelper.NumToLetter(RepHeader["Calls/Day"]);
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


            //worksheet = FormatHelper.LineupTable(worksheet, row, rankCount, AvgHeader.Count * 2, 1);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;
        }

        private Worksheet AddInboundAverageMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "Inbound Averages", row, AvgTypeHeader, out row, "Inbound");

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(AvgTypeHeader.Count * 2 + 1));
            row++;

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
            var iCCol = SupportMetricsHelper.NumToLetter(RepHeader["Inbound Calls"]);
            var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
            //var CpDCol = SupportMetricsHelper.NumToLetter(UHeader["Calls/Day"]);
            var wdCol = SupportMetricsHelper.NumToLetter(2);
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
                var aid = user.AverageInboundDuration();

                TableHelper.WriteTwoSidedTable(worksheet, row,
                    AvgTypeHeader["Avg Call Time"],
                    Formatter.LastInitial(user.Name),
                    Formatter.Duration(Convert.ToInt32(aid)),
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

            worksheet = FormatHelper.LineupTable(worksheet, row, rankCount, AvgTypeHeader.Count * 2, 1);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;

        }

        private Worksheet AddOutboundAverageMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            worksheet = HeaderHelper.WriteTwoSidedHeader(worksheet, "Outbound Averages", row, AvgTypeHeader, out row, "Outbound");

            var averageRep = SupportMetricsHelper.CreateAverageUser(reps);
            reps.Insert(0, averageRep);
            var genCount = reps.Count;

            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(AvgTypeHeader.Count * 2 + 1));
            row++;

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
            var iCCol = SupportMetricsHelper.NumToLetter(RepHeader["Outbound Calls"]);
            var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
            var wdCol = SupportMetricsHelper.NumToLetter(2);
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


            worksheet = FormatHelper.LineupTable(worksheet, row, rankCount, AvgTypeHeader.Count * 2, 1);
            newRow = row + rankCount + 1;
            reps.Remove(averageRep);
            return worksheet;
        }


        ////// Teams Tables
        private Worksheet AddTeamsOverview(List<Team> teams, Worksheet worksheet, int row, out int newRow, int offset = 0)
        {
            // determin Columns
            var startCol = SupportMetricsHelper.NumToLetter(offset + 1);
            var endCol = SupportMetricsHelper.NumToLetter(offset + TeamRefTable.Count + 1);
            var wdCol = SupportMetricsHelper.NumToLetter(WorkdaysCol);
            var iCol = SupportMetricsHelper.NumToLetter(offset + TeamRefTable["Inbound Calls"]);
            var oCol = SupportMetricsHelper.NumToLetter(offset + TeamRefTable["Outbound Calls"]);
            var aCol = SupportMetricsHelper.NumToLetter(offset + TeamRefTable["Absences"]);


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

            worksheet = FormatHelper.HeaderRow(worksheet, row, startCol, endCol);

            row++;

            worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(TeamRefTable.Count + 1));
            row++;

            // write content
            foreach (var teamRef in TeamsTableReference)
            {
                worksheet.Cells[row, offset + 1] = teamRef.Key;

                var ttCol = SupportMetricsHelper.NumToLetter(RepHeader["Adj Tickets"]);
                worksheet.Cells[row, offset + TeamRefTable["Adj Tickets"]].Formula = $"={ttCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Adj Tickets"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tpdCol = SupportMetricsHelper.NumToLetter(RepHeader["Tickets/Day"]);
                worksheet.Cells[row, offset + TeamRefTable["Tickets/Day"]].Formula = $"={tpdCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Tickets/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                var icCol = SupportMetricsHelper.NumToLetter(RepHeader["Inbound Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls"]].Formula = $"={icCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Inbound Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var ocCol = SupportMetricsHelper.NumToLetter(RepHeader["Outbound Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls"]].Formula = $"={ocCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Outbound Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tcCol = SupportMetricsHelper.NumToLetter(RepHeader["Calls"]);
                worksheet.Cells[row, offset + TeamRefTable["Calls"]].Formula = $"={tcCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Calls"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var tcpdCol = SupportMetricsHelper.NumToLetter(RepHeader["Calls/Day"]);
                worksheet.Cells[row, offset + TeamRefTable["Calls/Day"]].Formula = $"={tcpdCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Calls/Day"]].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                var actCol = SupportMetricsHelper.NumToLetter(RepHeader["Avg Call Time"]);
                worksheet.Cells[row, offset + TeamRefTable["Avg Call Time"]].Formula = $"={actCol}{teamRef.Value}";
                worksheet.Cells[row, offset + TeamRefTable["Avg Call Time"]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
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
                    worksheet.Range[startCol + (row), endCol + (row)].Interior.Color = XlRgbColor.rgbWhiteSmoke;
                    row++;
                    worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(TeamRefTable.Count + 1));
                }

                // alternate row colors with smoke
                //if (row % 2 == 0)
                //{
                //}

                row ++;
            }

            newRow = row + 1;
            return worksheet;
        }

        private Worksheet AddEachTeamsMetrics(List<Rep> reps, Worksheet worksheet, int row, out int newRow)
        {
            foreach (var team in Settings.Data.Teams)
            {
                if (team.Members.Count == 0)
                    continue;

                // skip team if not enabled for teams table
                if (team.IncludeInMetrics == false)
                    continue;

                // note not included if not included
                worksheet.Cells[row, 3] = team.IsExcluded ? "Is Excluded From Metrics" : "";
                worksheet.Cells[row, 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[row, 3].Font.Italic = true;
                worksheet.Cells[row, 3].Font.Size = 11;
               
                // add the rep header
                worksheet = HeaderHelper.WriteOneSidedHeader(worksheet, team.Name, row, RepHeader, out row);

                worksheet = FormatHelper.HighlighRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count));
                row++;

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

                

                worksheet = WriteAverageUser(worksheet, row, teamReps);
                var avgRepRow = row;
                row++;

                foreach (var rep in teamReps)
                {
                    worksheet = WriteRepMetrics(rep, worksheet, row);

                    // add absence cell reference back to general table
                    if (GeneralTableReference.ContainsKey(rep.Name))
                    {
                        var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
                        worksheet.Cells[row, RepHeader["Absences"]].Formula = $"={abCol}{GeneralTableReference[rep.Name]}";
                    }

                    row++;
                }

                worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count - 1));
                row++;

                worksheet = WriteTotalUser(worksheet, row, teamReps);
                worksheet.Rows[row].Font.Bold = true;
                var totalRepRow = row;
                row++;

                worksheet = FormatHelper.SeperateRow(worksheet, row, "A", SupportMetricsHelper.NumToLetter(RepHeader.Count - 1));
                row = row + 2; // add one space between teams

                if (!team.IsExcluded) TeamsTableReference.Add("Total " + team.Name, totalRepRow);
                if (!team.IsExcluded) TeamsTableReference.Add("Avg " + team.Name, avgRepRow);
                ProgressChanged?.Invoke(this, (int)(currentProgress += percentageStep));
            }

            newRow = row;
            newRow++;
            return worksheet;
        }


        ///////// 
        private Worksheet WriteTotalUser(Worksheet worksheet, int row, List<Rep> reps)
        {
            var totalUser = SupportMetricsHelper.CreateTotalUser(reps);

            worksheet = WriteRepMetrics(totalUser, worksheet, row);

            int stRow = row - reps.Count - 1;
            int edRow = row - 2;
            var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
            var stCol = SupportMetricsHelper.NumToLetter(1);
            var edCol = SupportMetricsHelper.NumToLetter(RepHeader.Count);

            // =SUM(P10:P28)
            worksheet.Cells[row, RepHeader["Absences"]].Formula = $"=SUM({abCol}{stRow}:{abCol}{edRow})";

            // format row
            worksheet.Range[stCol + row, edCol + row].Font.Bold = true;
            worksheet.Range[stCol + row, edCol + row].Font.Size = 12;

            return worksheet;
        }

        private Worksheet WriteAverageUser(Worksheet worksheet, int row, List<Rep> reps)
        {
            var avgUser = SupportMetricsHelper.CreateAverageUser(reps);

            worksheet = WriteRepMetrics(avgUser, worksheet, row);

            int stRow = row + 1;
            int edRow = row + reps.Count;
            var abCol = SupportMetricsHelper.NumToLetter(RepHeader["Absences"]);
            var stCol = SupportMetricsHelper.NumToLetter(1);
            var edCol = SupportMetricsHelper.NumToLetter(RepHeader.Count);

            // =IF(SUM(P10:P28)<=0,0,ROUND(AVERAGE(P10:P28),2))
            worksheet.Cells[row, RepHeader["Absences"]].Formula = $"=IF(SUM({abCol}{stRow}:{abCol}{edRow})<=0,0,ROUND(AVERAGE({abCol}{stRow}:{abCol}{edRow}),2))";

            // format row
            worksheet.Range[stCol + row, edCol + row].Font.Bold = true;

            return worksheet;
        }
    }
}