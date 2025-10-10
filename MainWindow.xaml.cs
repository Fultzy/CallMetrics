using CallMetrics.Controllers;
using CallMetrics.Models;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
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
        public MetricsReportGenerator ReportGenerator = new();

        public MainWindow()
        {
            InitializeComponent();
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
            // generate report
            ReportGenerator.Generate(ImportResults);
        }

        private void SetRepsToTeams_Click(object sender, RoutedEventArgs e)
        {
            // open another window that allows for drag and drop of reps to teams along side creation of each team. 

        }

        private void ClearData_Click(object sender, RoutedEventArgs e)
        {
            ImportResults.Clear();
            RepsDataGrid.ItemsSource = null;
            RepsDataGrid.Items.Refresh();
        }
    }
}