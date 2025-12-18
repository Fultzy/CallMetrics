using System.Security.Policy;

namespace CallMetrics.Utilities
{
    public static class Notifications
    {
        public static Notification GenerateComplete = new Notification
        {
            Title = "Report Generation Complete",
            Message = $"Your metrics report has been successfully generated!\n Saved to {Settings.DefaultReportPath}"
        };

        public static Notification ImportCallsComplete = new Notification
        {
            Title = "Import Complete",
            Message = "The Call report has been successfully imported!"
        };

        public static Notification ImportTicketsComplete = new Notification
        {
            Title = "Import Complete",
            Message = "The Ticket report has been successfully imported!"
        };

        public static Notification ImportFail = new Notification
        {
            Title = "Import Failed",
            Message = "Failed to import.\nEnsure the file is a valid and not open in another program."
        };

        public static Notification NoTeams = new Notification
        {
            Title = "No Teams Created",
            Message = "Create and Populate atleast one Team before Generating."
        };

        public static Notification NoRepsAssignedToTeams = new Notification
        {
            Title = "No Reps Assigned to Teams",
            Message = "Import a report and Assign Reps to Teams before Generating."
        };

        public static Notification NoReps = new Notification
        {
            Title = "No Reps Assigned to Teams",
            Message = "Import a report and Assign Reps to Teams before Generating."
        };

        public static Notification NoCalls = new Notification
        {
            Title = "No Calls Imported",
            Message = "Import a report before generating metrics."
        };

        public static Notification NoTickets = new Notification
        {
            Title = "No Tickets Imported",
            Message = "Import a report before generating metrics."
        };

        public static Notification NoData = new Notification
        {
            Title = "No MetricsData Imported",
            Message = "Import a report before generating metrics."
        };

        public static Notification NoTeamInMetricsOrDepartments = new Notification
        {
            Title = "No Teams Included in Metrics",
            Message = "Enable Include in Metrics or set a team to Department"
        };
    }

    public struct Notification
    {
        public string Title;
        public string Message;
    }
}