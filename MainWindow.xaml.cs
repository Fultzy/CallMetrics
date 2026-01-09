using CallMetrics.Controllers.Generators;
using CallMetrics.Controllers.Readers.CallTracker;
using CallMetrics.Controllers.Readers.Dynamics;
using CallMetrics.Controllers.Readers.Five9;
using CallMetrics.Controllers.Readers.Nextiva;
using CallMetrics.Data;
using CallMetrics.Menus;
using CallMetrics.Models;
using CallMetrics.Utilities;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static MaterialDesignThemes.Wpf.Theme;

namespace CallMetrics
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MetricsReport ReportGenerator = new();

        public MainWindow()
        {
            InitializeComponent();
            
            Settings.Load();

            this.SourceInitialized += MainWindow_SourceInitialized;
            ReportGenerator.ReportProgressChanged += (s, e) => UpdateProgressBar(e);
            ProgressBar.Value = 0;

            // set toggles based on settings
            TicketImportTypeToggle.IsChecked = Settings.TicketImportType == ImportType.Dynamics;
            CallImportTypeToggle.IsChecked = Settings.CallImportType == ImportType.Five9;
            CheckTicketImportType();
            CheckCallImportType();
        }

        public void UpdateProgressBar(int value)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = value;
            });
        }

        private async void ImportCallReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ClearButton.IsEnabled = false;
                ImportTicketsButton.IsEnabled = false;
                ImportCallsButton.IsEnabled = false;
                GenerateButton.IsEnabled = false;

                var calls = new List<Call>();
                if (Settings.CallImportType == ImportType.Nextiva)
                {
                    calls = await new NextivaReader().Start();
                }
                else if (Settings.CallImportType == ImportType.Five9)
                {
                    calls = await new Five9Reader().Start();
                }

                MetricsData.AddCalls(calls);
                if (MetricsData.Calls.Count == 0)
                {
                    Notify(Notifications.ImportFail);
                    return;
                }

                

                RepsDataGrid.ItemsSource = MetricsData.Reps;
                RepsDataGrid.Items.Refresh();

                Notify(Notifications.ImportCallsComplete);
            }
            finally
            {
                ClearButton.IsEnabled = true;
                ImportTicketsButton.IsEnabled = true;
                ImportCallsButton.IsEnabled = true;
                GenerateButton.IsEnabled = true;
            }
        }

        private async void ImportTicketReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ClearButton.IsEnabled = false;
                ImportTicketsButton.IsEnabled = false;
                ImportCallsButton.IsEnabled = false;
                GenerateButton.IsEnabled = false;

                var tickets = new List<Ticket>();
                if (Settings.TicketImportType == ImportType.Dynamics)
                {
                    tickets = await new DynamicsReader().Start();
                }
                else if (Settings.TicketImportType == ImportType.CallTracker)
                {
                    tickets = await new CallTrackerReader().Start();
                }

                MetricsData.AddTickets(tickets);
                if (MetricsData.Tickets.Count == 0)
                {
                    Notify(Notifications.ImportFail);
                    return;
                }


                RepsDataGrid.ItemsSource = MetricsData.Reps;
                RepsDataGrid.Items.Refresh();

                Notify(Notifications.ImportTicketsComplete);
            }
            finally
            {
                ClearButton.IsEnabled = true;
                ImportTicketsButton.IsEnabled = true;
                ImportCallsButton.IsEnabled = true;
                GenerateButton.IsEnabled = true;

            }
        }


        private async void GenerateReport_Click(object sender, RoutedEventArgs e)
        {
            if (MetricsData.Calls.Count == 0)
            {
                Notify(Notifications.NoCalls);
                return;
            }

            // temp removed to allow use of only call data
            //if (MetricsData.Tickets.Count == 0)
            //{
            //    Notify(Notifications.NoTickets);
            //    return;
            //}

            if (MetricsData.Reps.Count == 0)
            {
                Notify(Notifications.NoReps);
                return;
            }

            GenerateButton.IsEnabled = false;
            ImportTicketsButton.IsEnabled = false;
            ImportCallsButton.IsEnabled = false;
            ClearButton.IsEnabled = false;
            Task task = Task.Run(() =>
            {
                ReportGenerator.Generate(MetricsData.Reps, Settings.DefaultReportPath);
            });

            await task;
            ClearButton.IsEnabled = true;
            ImportTicketsButton.IsEnabled = true;
            ImportCallsButton.IsEnabled = true;
            GenerateButton.IsEnabled = true;
            Notify(Notifications.GenerateComplete);
            ProgressBar.Value = 0;
        }

        private void Notify(Notification noti)
        {
            var notification = new Controls.NotificationCard(noti);
            NotificationStack.Children.Add(notification);
        }

        private void SetRepsToTeams_Click(object sender, RoutedEventArgs e)
        {
            TeamsWindow teamsWindow = new TeamsWindow(MetricsData.Reps);
            teamsWindow.ShowDialog();
        }

        private void SetRepAliases_Click(object sender, RoutedEventArgs e)
        {
            AliasWindow aliasWindow = new AliasWindow(MetricsData.Reps);
            aliasWindow.ShowDialog();
        }

        private void ClearData_Click(object sender, RoutedEventArgs e)
        {
            MetricsData.Clear();
            RepsDataGrid.ItemsSource = null;
            RepsDataGrid.Items.Refresh();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settingsWindow = new SettingsWindow();
            settingsWindow.ShowDialog();
        }

        private void TicketImportTypeToggle_Checked(object sender, RoutedEventArgs e)
        {
            CheckTicketImportType();
            Settings.Save();
        }

        private void TicketImportTypeToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckTicketImportType();
            Settings.Save();
        }

        private void CheckTicketImportType()
        {
            if (TicketImportTypeToggle.IsChecked == false)
            {
                Settings.TicketImportType = ImportType.CallTracker;
                ImportTicketsButton.Content = "Import CallTracker Tickets";
                ImportTicketsButton.Background = new SolidColorBrush(Color.FromRgb(33, 150, 243));
            }
            else
            {
                Settings.TicketImportType = ImportType.Dynamics;
                ImportTicketsButton.Content = "Import Dynamics Tickets";
                ImportTicketsButton.Background = new SolidColorBrush(Color.FromRgb(10, 109, 187));
            }
        }

        private void CallImportTypeToggle_Checked(object sender, RoutedEventArgs e)
        {
            CheckCallImportType();
            Settings.Save();
        }

        private void CallImportTypeToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckCallImportType();
            Settings.Save();
        }

        public void CheckCallImportType()
        {
            if (CallImportTypeToggle.IsChecked == false)
            {
                Settings.CallImportType = ImportType.Nextiva;
                ImportCallsButton.Content = "Import Nextiva Calls";
                ImportCallsButton.Background = new SolidColorBrush(Color.FromRgb(33, 150, 243));
            }
            else
            {
                Settings.CallImportType = ImportType.Five9;
                ImportCallsButton.Content = "Import Five9 Calls";
                ImportCallsButton.Background = new SolidColorBrush(Color.FromRgb(10, 109, 187));
            }
        }

        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            IntPtr handle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(handle)?.AddHook(new HwndSourceHook(WindowProc));
        }

        private IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_NCHITTEST = 0x0084;
            const int HTCLIENT = 1;
            const int HTLEFT = 10;
            const int HTRIGHT = 11;
            const int HTTOP = 12;
            const int HTTOPLEFT = 13;
            const int HTTOPRIGHT = 14;
            const int HTBOTTOM = 15;
            const int HTBOTTOMLEFT = 16;
            const int HTBOTTOMRIGHT = 17;

            if (msg == WM_NCHITTEST)
            {
                handled = true;
                var mousePoint = new Point((lParam.ToInt32() & 0xFFFF), (lParam.ToInt32() >> 16));

                // Convert screen to window position
                var windowPos = this.PointFromScreen(mousePoint);
                double resizeBorder = 8; // thickness of resize area

                bool left = windowPos.X <= resizeBorder;
                bool right = windowPos.X >= this.ActualWidth - resizeBorder;
                bool top = windowPos.Y <= resizeBorder;
                bool bottom = windowPos.Y >= this.ActualHeight - resizeBorder;

                if (left && top) return new IntPtr(HTTOPLEFT);
                if (right && top) return new IntPtr(HTTOPRIGHT);
                if (left && bottom) return new IntPtr(HTBOTTOMLEFT);
                if (right && bottom) return new IntPtr(HTBOTTOMRIGHT);
                if (left) return new IntPtr(HTLEFT);
                if (right) return new IntPtr(HTRIGHT);
                if (top) return new IntPtr(HTTOP);
                if (bottom) return new IntPtr(HTBOTTOM);

                return new IntPtr(HTCLIENT);
            }

            return IntPtr.Zero;
        }
    }
}