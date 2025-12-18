using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace CallMetrics.Menus
{
    /// <summary>
    /// Interaction logic for ImportWindow.xaml
    /// </summary>
    public partial class ImportWindow : Window
    {
        public ImportWindow()
        {
            InitializeComponent();
        }

        private void ImportCallRecordsButton_Click(object sender, RoutedEventArgs e)
        {
            //ImportCallRecordsWindow importCallRecordsWindow = new ImportCallRecordsWindow();
            //importCallRecordsWindow.Owner = this;
            //importCallRecordsWindow.ShowDialog();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CustomMappingButton_Click(object sender, RoutedEventArgs e)
        {
            //CustomMappingWindow mappingWindow = new CustomMappingWindow();
            //mappingWindow.Owner = this;
            //mappingWindow.ShowDialog();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
    }
}
