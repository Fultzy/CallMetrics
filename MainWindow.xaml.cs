using CallMetrics.Controllers;
using CallMetrics.Controllers.Generators;
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
        public string OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public MainWindow()
        {
            InitializeComponent();
            Settings.Load();
            this.SourceInitialized += MainWindow_SourceInitialized;
        }

        private void ImportNextivaReport_Click(object sender, RoutedEventArgs e)
        {
            // open explorer to select file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.DefaultExt = ".csv";
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                var filePath = openFileDialog.FileName;
                ImportResults = ReportReader.Read(filePath);
                RepsDataGrid.ItemsSource = ImportResults;
                RepsDataGrid.Items.Refresh();
            }
        }

        private void GenerateReport_Click(object sender, RoutedEventArgs e)
        {
            if (ImportResults.Count == 0)
            {
                MessageBox.Show("No data to generate report. Please import a Nextiva report first. Loser", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ReportGenerator.Generate(ImportResults, OutputDirectory);
        }

        private void SetRepsToTeams_Click(object sender, RoutedEventArgs e)
        {
            // open the TeamsWindow.xaml
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