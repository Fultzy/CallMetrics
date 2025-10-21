using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CallMetrics.Controls
{
    /// <summary>
    /// Interaction logic for TeamControl.xaml
    /// </summary>
    public partial class TeamControl : UserControl
    {
        public event EventHandler DeleteTeamClicked;
        public string TeamName;
        internal Action<object, object> RefreshRequest;

        public TeamControl()
        {
            InitializeComponent();
        }

        public TeamControl(string teamName)
        {
            InitializeComponent();
            TeamName = teamName;
            TeamNameLabel.Content = teamName;

            Team team = Settings.Teams.Where(t => t.Name == TeamName).FirstOrDefault();
            if (team.Equals(default(SettingsData)))
                return;

            IncludeInMetricsCheckBox.IsChecked = team.IncludeInMetrics;

            IsDepartmentCheckBox.IsChecked = team.IsDepartment;
            if (team.IsDepartment)
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA9A6A6"));
                IncludeLabel.Foreground = brush;

                IncludeInMetricsCheckBox.IsChecked = false;
                IncludeInMetricsCheckBox.IsEnabled = false;
            }

            HideTeamCheckBox.IsChecked = team.HideTeam;
            if (team.HideTeam)
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA9A6A6"));

                IncludeLabel.Foreground = brush;
                IncludeInMetricsCheckBox.IsChecked = false;
                IncludeInMetricsCheckBox.IsEnabled = false;

                DepartmentLabel.Foreground = brush;
                IsDepartmentCheckBox.IsChecked = false;
                IsDepartmentCheckBox.IsEnabled = false;
            }


            RefreshTeamMembers();

        }

        private void DeleteTeamButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteTeamClicked?.Invoke(this, EventArgs.Empty);
        }

        public void RefreshTeamMembers()
        {
            Team team = Settings.Teams.Where(t => t.Name == TeamName).FirstOrDefault();
            if (team.Equals(default(SettingsData)))
                return;
            
            RepsWrapPanel.Children.Clear();
            foreach (var member in team.Members) 
            {
                var repItem = new RepItem(member);
                RepsWrapPanel.Children.Add(repItem);
            }
        }

        // Checkboxes
        public void IncludeInMetricsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.IncludeInMetrics = true;
                Settings.Teams[teamIndex] = updatedTeam;

                Settings.Save();
            }
        }

        public void IncludeInMetricsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.IncludeInMetrics = false;
                Settings.Teams[teamIndex] = updatedTeam;

                Settings.Save();
            }
        }

        public void IsDepartmentCheckBox_Checked(object sender, RoutedEventArgs e)
        {

            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.IsDepartment = true;
                updatedTeam.IncludeInMetrics = false;
                Settings.Teams[teamIndex] = updatedTeam;

                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA9A6A6"));
                IncludeLabel.Foreground = brush;

                IncludeInMetricsCheckBox.IsChecked = false;
                IncludeInMetricsCheckBox.IsEnabled = false;

                Settings.Save();
            }

        }

        public void IsDepartmentCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.IsDepartment = false;
                Settings.Teams[teamIndex] = updatedTeam;

                // #FFFFFFFF
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));
                IncludeLabel.Foreground = brush;
                IncludeInMetricsCheckBox.IsEnabled = true;

                Settings.Save();
            }
        }

        public void HideTeamCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.HideTeam = true;
                updatedTeam.IncludeInMetrics = false;
                updatedTeam.IsDepartment = false;

                Settings.Teams[teamIndex] = updatedTeam;
                Settings.Save();

                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFA9A6A6"));

                IncludeLabel.Foreground = brush;
                IncludeInMetricsCheckBox.IsChecked = false;
                IncludeInMetricsCheckBox.IsEnabled = false;

                DepartmentLabel.Foreground = brush;
                IsDepartmentCheckBox.IsChecked = false;
                IsDepartmentCheckBox.IsEnabled = false;
            }

            RefreshRequest?.Invoke(this, e);
        }

        public void HideTeamCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            int teamIndex = Settings.Teams.FindIndex(t => t.Name == TeamName);
            if (teamIndex >= 0)
            {
                var updatedTeam = Settings.Teams[teamIndex];
                updatedTeam.HideTeam = false;
                updatedTeam.IncludeInMetrics = false;
                updatedTeam.IsDepartment = false;

                Settings.Teams[teamIndex] = updatedTeam;
                Settings.Save();

                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFFFFFF"));

                IncludeInMetricsCheckBox.IsEnabled = true;

                IsDepartmentCheckBox.IsEnabled = true;
            }

            RefreshRequest?.Invoke(this, e);
        }



        private void RepsWrapPanel_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(RepItem)))
                e.Effects = DragDropEffects.Move;   // <--- REQUIRED
            else
                e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void RepsWrapPanel_Drop(object sender, DragEventArgs e)
        {
            Team team = Settings.Teams.Where(t => t.Name == TeamName).FirstOrDefault();
            if (team.IsNull())
                return;

            if (e.Data.GetData(typeof(RepItem)) is RepItem item)
            {
                if (item.Parent is ListBox parentBox)
                {
                    parentBox.Items.Remove(item);
                    this.RemoveLogicalChild(item);
                }

                if (item.Parent is Panel parentParent)
                {
                    parentParent.Children.Remove(item);
                    this.RemoveLogicalChild(item);
                }

                foreach (var tm in Settings.Teams.Where(t => t.Members.Contains(item.RepName)))
                {
                    tm.Members.Remove(item.RepName);
                }

                team.Members.Add(item.RepName);
                Settings.Save();

                var newItem = new RepItem(TeamName);
                (sender as ListBox)?.Items.Add(newItem);
            }

            RefreshTeamMembers();
        }

        private void RootDragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            e.Handled = false; // <--- Important! Do NOT swallow it
        }
    }
}
