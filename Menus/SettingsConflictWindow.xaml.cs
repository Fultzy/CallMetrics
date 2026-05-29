using CallMetrics.Models;
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
    /// Interaction logic for SettingsConflictWindow.xaml
    /// </summary>
    public partial class SettingsConflictWindow : Window
    {
        public bool ApplyNewSettings;

        public SettingsConflictWindow()
        {
            InitializeComponent();
        }

        public SettingsConflictWindow(SettingsData newSettings, SettingsData oldSettings)
        {
            InitializeComponent();

            SetOldSettings(oldSettings);
            SetNewSettings(newSettings);

        }

        private void SetOldSettings(SettingsData settings)
        {
            CurrentVersionText.Text = "Version: " + settings.Version;
            CurrentLastSavedText.Text = "Last Saved:\n" + settings.LastSave.ToString();
            CurrentRepCountText.Text = $"Aliases: {settings.Aliases.Count}";
            CurrentTeamCountText.Text = $"Teams: {settings.Teams.Count}";
        }

        private void SetNewSettings(SettingsData settings)
        {
            NewVersionText.Text = "Version: " + settings.Version;
            NewLastSavedText.Text = "Last Saved:\n" + settings.LastSave.ToString();
            NewRepCountText.Text = $"Aliases: {settings.Aliases.Count}";
            NewTeamCountText.Text = $"Teams: {settings.Teams.Count}";
        }

        private void ApplyNewButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyNewSettings = true;
            this.Close();
        }

        private void KeepCurrentButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyNewSettings = false;
            this.Close();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
    }
}
