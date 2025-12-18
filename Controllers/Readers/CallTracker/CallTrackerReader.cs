using CallMetrics.Models;
using CallMetrics.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CallMetrics.Controllers.Readers.CallTracker
{
    internal class CallTrackerReader
    {
        internal async Task<List<Ticket>> Start()
        {
            // open explorer to select file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.DefaultDirectory = Settings.DefaultReportPath;
            openFileDialog.DefaultExt = ".xlsx";
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string filePath = openFileDialog.FileName;
                return await Task.Run(() => Read(filePath));
            }

            return new List<Ticket>();
        }

        internal async Task<List<Ticket>> Read(string filePath)
        {
            try
            {
                return await Task.Run(() => ProcessFile(filePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading CallTracker report: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<Ticket>();
            }
        }


        private bool DebugMode = true; // Set to true for debugging
        private async Task<List<Ticket>> ProcessFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                Debug.WriteLine("No file selected.");
                return new List<Ticket>();
            }

            Microsoft.Office.Interop.Excel.Application excelApp = new();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];

            try
            {
                int rowCount = worksheet.UsedRange.Rows.Count;
                int maxThreads = Environment.ProcessorCount - 1;

                ////// Debug mode is single threaded
                if (DebugMode) maxThreads = 1;

                SemaphoreSlim semaphore = new(maxThreads);
                var tasks = new List<Task<List<Ticket>>>();
                int chunkSize = rowCount / maxThreads;
                int currentRow = 1;

                for (int i = 0; i < maxThreads; i++)
                {
                    int startRow = currentRow;
                    int endRow = FindNextBlankRow(worksheet, startRow + chunkSize, rowCount);
                    int thread = i + 1;

                    await semaphore.WaitAsync();
                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            var data = ExtractData(worksheet, startRow, endRow);

                            return data;
                        }
                        catch (Exception e)
                        {
                            throw new Exception($"Error in thread {thread} processing rows {startRow} to {endRow}: {e.Message}", e);
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    }));

                    currentRow = endRow + 1;
                    if (currentRow > rowCount)
                        break;
                }

                var ticks = await Task.WhenAll(tasks);
                return ticks.SelectMany(t => t).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing CallTracker report: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<Ticket>();
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        

        private List<Ticket> ExtractData(Worksheet worksheet, int startRow, int endRow)
        {
            List<Ticket> results = new List<Ticket>();

            Microsoft.Office.Interop.Excel.Range range = worksheet.Range[$"A{startRow}:F{endRow}"];
            object[,] values = range.Value2;

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                // Skip blank rows  
                if (string.IsNullOrWhiteSpace(values[row, 1]?.ToString())) continue;

                // Parse ticket  
                if (ParseTicket(row, startRow, values) is Ticket newTicket)
                {
                    row = row + 2;

                    if (ParseTicketEntries(row, startRow, values) is (List<TicketEntry> ticketEntries, int newRow))
                    {
                        newTicket.TicketEntries = ticketEntries;
                        results.Add(newTicket);
                        row = newRow;
                    }
                }
            }

            return results;
        }

        private Ticket ParseTicket(int row, int startRow, object[,] values)
        {
            // Check if row is blank
            if (values[row, 1] == null || string.IsNullOrWhiteSpace(values[row, 1]?.ToString()))
            {
                row++;
            }

            // Read ClientName MetricsData
            string clientName = values[row, 1]?.ToString() ?? string.Empty;
            string dateTime = values[row, 2]?.ToString() ?? string.Empty;
            string callRec = values[row, 3]?.ToString() ?? string.Empty;

            // Check if ClientName Header
            if (clientName == "Client" || dateTime == "Date DateTime" || callRec == "Call Rec #")
            {
                row++; // move to ticket data  
                clientName = values[row, 1]?.ToString() ?? string.Empty;
                dateTime = values[row, 2]?.ToString() ?? string.Empty;
                callRec = values[row, 3]?.ToString() ?? string.Empty;
            }

            // Validate DateTime
            if (!DateTime.TryParse(dateTime, out DateTime parsedDateTime))
            {
                return null;
            }

            return new Ticket
            {
                ClientName = clientName,
                DateTime = parsedDateTime,
                CallRecNumber = callRec
            };
        }

        private (List<TicketEntry>, int) ParseTicketEntries(int row, int startRow, object[,] values)
        {
            List<TicketEntry> ticketEntries = new List<TicketEntry>();

            // Check ticket headers  
            string ticketHeader1 = values[row, 1]?.ToString() ?? string.Empty;
            string ticketHeader2 = values[row, 2]?.ToString() ?? string.Empty;
            if (ticketHeader1 == "Call Rec #" || ticketHeader2 == "Date & DateTime")
            {
                row++; // move to ticket entries  
            }
            else
            {
                return (ticketEntries, row);
            }

            // Read Ticket TicketEntries  
            while (row <= values.GetLength(0) && !string.IsNullOrWhiteSpace(values[row, 1]?.ToString()))
            {
                // Skip if invalid datetime  
                string ticketDateTime = values[row, 2]?.ToString() ?? string.Empty;
                if (!DateTime.TryParse(ticketDateTime, out DateTime parsedTicketDateTime))
                {
                    row++;
                    continue;
                }

                // assigned Calls
                var toRepName = values[row, 3]?.ToString() ?? string.Empty;
                var byRepName = values[row, 4]?.ToString() ?? string.Empty;

                // Create new TicketEntry
                TicketEntry entry = new TicketEntry
                {
                    CallRecNumber = int.TryParse(values[row, 1]?.ToString(), out int callRecNum) ? callRecNum : 0,
                    DateTime = parsedTicketDateTime,
                    Status = values[row, 5]?.ToString() ?? string.Empty,
                    Description = values[row, 6]?.ToString() ?? string.Empty,

                    AssignedByName = toRepName,
                    AssignedToName = byRepName,
                };

                ticketEntries.Add(entry);
                row++;
            }

            return (ticketEntries, row);
        }


        // report helper methods
        private int FindNextBlankRow(Worksheet worksheet, int startRow, int rowCount)
        {
            Microsoft.Office.Interop.Excel.Range range = worksheet.Range[$"A{startRow}:F{rowCount}"];
            object[,] values = range.Value2;

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                bool isBlank = true;
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    if (values[row, col] != null && !string.IsNullOrWhiteSpace(values[row, col].ToString()))
                    {
                        isBlank = false;
                        break;
                    }
                }
                if (isBlank)
                {
                    return startRow + row - 1;
                }
            }
            return rowCount;
        }

        private bool IsRowBlank(Worksheet worksheet, int row)
        {
            // optimized to only check first column. why it was checking 6... idk. 
            if (!string.IsNullOrWhiteSpace(GetCellValue(worksheet, row, 1)))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private string GetCellValue(Worksheet worksheet, int row, int col)
        {
            var cellValue = worksheet.Cells[row, col].Value2;
            return cellValue?.ToString().Trim() ?? string.Empty;
        }
    }
}
