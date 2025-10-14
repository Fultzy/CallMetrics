using System.Security.Policy;

namespace CallMetrics.Utilities
{
    public static class Notifications
    {
        public static Notification GenerateComplete = new Notification
        {
            Title = "Report Generation Complete",
            Message = "Your metrics report has been successfully generated!\nFind it on your Desktop!"
        };

        public static Notification ImportComplete = new Notification
        {
            Title = "Import Complete",
            Message = "The Nextiva report has been successfully imported!"
        };

        public static Notification ImportFail = new Notification
        {
            Title = "Import Failed",
            Message = "Failed to import the Nextiva report.\nEnsure the file is a valid CSV and not open in another program."
        };

        public static Notification NoTeams = new Notification
        {
            Title = "No Teams Created",
            Message = "Report will be generated without Teams"
        };
        public static Notification NoData = new Notification
        {
            Title = "No Data Imported",
            Message = "Import a Nextiva report before generating metrics."
        };
    }

    public struct Notification
    {
        public string Title;
        public string Message;
    }
}