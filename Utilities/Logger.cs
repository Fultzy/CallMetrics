using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace CallMetrics.Utilities
{
    public static class Logger
    {
        private static readonly ConcurrentQueue<(string message, string logName)> logQueue = new ConcurrentQueue<(string message, string logName)>();
        private static readonly Task logTask;
        private static readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();

        static Logger()
        {
            logTask = Task.Run(ProcessLogQueue, cancellationTokenSource.Token);
        }

        private static async Task ProcessLogQueue()
        {
            while (!cancellationTokenSource.Token.IsCancellationRequested)
            {
                while (logQueue.TryDequeue(out var logEntry))
                {
                    WriteToFile(logEntry.message, logEntry.logName);
                }
                await Task.Delay(100); // Adjust delay as needed
            }
        }

        private static string EnqueueLog(string message, string logName, bool timeStamp)
        {
            if (timeStamp)
            {
                message = DateTime.Now + " - " + message;
            }
            else
            {
                message = "\n" + message;
            }

            logQueue.Enqueue((message, logName));
            return message;
        }

        private static string WriteToFile(string message, string logName)
        {
            var logfolder = System.IO.Path.Combine(Environment.CurrentDirectory, "Logs");
            var logPath = System.IO.Path.Combine(logfolder, $"{logName}.log");

            if (!Directory.Exists(logfolder))
            {
                Directory.CreateDirectory(logfolder);
            }

            CheckLogFile(logPath); // Check if the log file is older than 7 days

            File.AppendAllText(logPath, message + Environment.NewLine);

            return message;
        }

        private static void CheckLogFile(string filepath)
        {
            if (!File.Exists(filepath))
            {
                File.Create(filepath).Dispose();
            }

            // If the log file size is larger than 10 MB, archive it
            var fileInfo = new FileInfo(filepath);
            if (fileInfo.Length > 10 * 1024 * 1024)
            {
                ArchiveLogFile(filepath);
            }
        }

        private static void ArchiveLogFile(string filepath)
        {
            var archivePath = System.IO.Path.Combine(Environment.CurrentDirectory, "Logs/LogArchive");
            if (!Directory.Exists(archivePath))
            {
                Directory.CreateDirectory(archivePath);
            }

            var fileName = System.IO.Path.GetFileName(filepath);
            var archiveFilePath = System.IO.Path.Combine(archivePath, fileName);

            // Move the log file to the archive folder and overwrite if exists
            File.Move(filepath, archiveFilePath);
        }

        // Log methods
        public static string MainLog(string message, bool timeStamp = true)
        {
            return EnqueueLog(message, "MainLog", timeStamp);
        }

        public static string ExceptionLog(string message, bool timeStamp = true)
        {
            return EnqueueLog(message, "ExceptionLog", timeStamp);
        }
    }
}
