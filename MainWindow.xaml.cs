using CallMetrics.Controllers.Generators;
using CallMetrics.Controllers.Readers.Nextiva;
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

namespace CallMetrics
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<RepData> ImportResults = new();
        public NextivaReportReader ReportReader = new();
        public MetricsReport ReportGenerator = new();

        public MainWindow()
        {
            InitializeComponent();
            Settings.Load();
            this.SourceInitialized += MainWindow_SourceInitialized;
            ReportGenerator.ReportProgressChanged += (s, e) => UpdateProgressBar(e);
            ProgressBar.Value = 0;
        }

        public void UpdateProgressBar(int value)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = value;
            });
        }

        private void ImportNextivaReport_Click(object sender, RoutedEventArgs e)
        {
            // open explorer to select file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.DefaultDirectory = Settings.DefaultReportPath;
            openFileDialog.DefaultExt = ".csv";
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                ClearButton.IsEnabled = false;
                ImportButton.IsEnabled = false;
                GenerateButton.IsEnabled = false;

                var filePath = openFileDialog.FileName;
                ImportResults = ReportReader.Read(filePath);
                if (ImportResults.Count == 0)
                {
                    Notify(Notifications.ImportFail);
                    return;
                }

                RepsDataGrid.ItemsSource = ImportResults;
                RepsDataGrid.Items.Refresh();

                ClearButton.IsEnabled = true;
                ImportButton.IsEnabled = true;
                GenerateButton.IsEnabled = true;

                Notify(Notifications.ImportComplete);
            }
        }

        private async void GenerateReport_Click(object sender, RoutedEventArgs e)
        {
            if (ImportResults.Count == 0)
            {
                Notify(Notifications.NoData);
                return;
            }
            
            //if (Settings.Teams.Count == 0)
            //{
            //    Notify(Notifications.NoTeams);
            //    return;
            //}

            //if (!Settings.Teams.Any(t => t.Members.Count > 0))
            //{
            //    Notify(Notifications.NoReps);
            //    return;
            //}

            //if (!Settings.Teams.Any(t => t.IncludeInMetrics || t.IsDepartment))
            //{
            //    Notify(Notifications.NoTeamInMetricsOrDepartments);
            //    return;
            //}

            GenerateButton.IsEnabled = false;
            ImportButton.IsEnabled = false;
            ClearButton.IsEnabled = false;
            Task task = Task.Run(() =>
            {
                ReportGenerator.Generate(ImportResults, Settings.DefaultReportPath);
            });

            await task;
            ClearButton.IsEnabled = true;
            ImportButton.IsEnabled = true;
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
            TeamsWindow teamsWindow = new TeamsWindow(ImportResults);
            teamsWindow.ShowDialog();
        }

        private void ClearData_Click(object sender, RoutedEventArgs e)
        {
            ImportResults.Clear();
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