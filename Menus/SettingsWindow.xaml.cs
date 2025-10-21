using CallMetrics.Utilities;
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
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            
            this.Loaded += ApplySettings;
        }

        public void ApplySettings(object sender, RoutedEventArgs e)
        {
            if (Settings.RankedRepsCount < 1 || Settings.RankedRepsCount > 999)
                Settings.RankedRepsCount = 10;

            NumberOfRowsTextBox.Text = Settings.RankedRepsCount.ToString();
            AutoOpenReportCheckBox.IsChecked = Settings.AutoOpenReport;
            DefaultSaveLocationTextBlock.Text = Settings.DefaultReportPath;
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // open explorer to select file
            var openFolderDialog = new Microsoft.Win32.OpenFolderDialog();
            openFolderDialog.DefaultDirectory = Settings.DefaultReportPath;
            var result = openFolderDialog.ShowDialog();
            if (result == true)
            {
                DefaultSaveLocationTextBlock.Text = openFolderDialog.FolderName;
                Settings.DefaultReportPath = openFolderDialog.FolderName;
                Settings.Save();
            }
        }

        private void OnRowsTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(NumberOfRowsTextBox.Text, out int value))
            {
                if (value >= 1 && value <= 999)
                {
                    Settings.RankedRepsCount = value;
                    Settings.Save();
                }
            }
        }

        private void DecreaseRowsButton_Click(object sender, RoutedEventArgs e)
        {
            if (Settings.RankedRepsCount > 1 && Settings.RankedRepsCount < 999)
            {
                Settings.RankedRepsCount--;
                NumberOfRowsTextBox.Text = Settings.RankedRepsCount.ToString();
                Settings.Save();
            }
        }

        private void IncreaseRowsButton_Click(object sender, RoutedEventArgs e)
        {
            if (Settings.RankedRepsCount >= 1 && Settings.RankedRepsCount < 999)
            {
                Settings.RankedRepsCount++;
                NumberOfRowsTextBox.Text = Settings.RankedRepsCount.ToString();
                Settings.Save();
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void AutoOpenReportCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Settings.AutoOpenReport = true;
            Settings.Save();
        }

        private void AutoOpenReportCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Settings.AutoOpenReport = false;
            Settings.Save();
        }
    }
}
