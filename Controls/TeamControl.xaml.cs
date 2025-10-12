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

        public TeamControl()
        {
            InitializeComponent();
        }

        public TeamControl(string teamName)
        {
            InitializeComponent();
            TeamName = teamName;
            TeamNameLabel.Content = teamName;

            if (Settings.IgnoreTeamMetrics.Contains(TeamName))
            {
                IncludeInMetricsCheckBox.IsChecked = false;
            }
            else
            {
                IncludeInMetricsCheckBox.IsChecked = true;
            }
            
            RefreshTeamMembers();

        }

        private void DeleteTeamButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteTeamClicked?.Invoke(this, EventArgs.Empty);
        }

        public void RefreshTeamMembers()
        {
            if (Settings.Teams[TeamName] != null)
            {
                RepsWrapPanel.Children.Clear();
                foreach (var member in Settings.Teams[TeamName]) 
                {
                    var repItem = new RepItem(member);
                    RepsWrapPanel.Children.Add(repItem);
                }
            }
        }

        public void IncludeInMetricsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (Settings.IgnoreTeamMetrics.Contains(TeamName))
            {
                Settings.IgnoreTeamMetrics.Remove(TeamName);
                Settings.Save();
            }
        }

        public void IncludeInMetricsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!Settings.IgnoreTeamMetrics.Contains(TeamName))
            {
                Settings.IgnoreTeamMetrics.Add(TeamName);
                Settings.Save();
            }
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

                foreach (var kvp in Settings.Teams.Where(k => k.Value.Contains(item.RepName)))
                {
                    kvp.Value.Remove(item.RepName);
                }

                Settings.Teams[TeamName].Add(item.RepName);
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
