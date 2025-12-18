using CallMetrics.Controls;
using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace CallMetrics.Menus
{
    /// <summary>  
    /// Interaction logic for TeamsWindow.xaml  
    /// </summary>  
    public partial class TeamsWindow : Window
    {
        private List<Rep> _reps = new();
        private string _newTeamName = string.Empty;
        private bool _showHiddenTeams = false;

        public TeamsWindow(List<Rep> importResults)
        {
            InitializeComponent();
            this.SourceInitialized += TeamsWindow_SourceInitialized;
            this.Closed += (s,e) => Settings.Load();

            _reps = importResults;
            RefreshTeamsList();
            RefreshRepsList();
        }

        private void AddNewTeam_Click(object sender, RoutedEventArgs e)
        {

            if (Settings.Teams.Any(t => t.Name.Equals(_newTeamName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("A team with this name already exists. Please choose a different name.", "Duplicate Team Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_newTeamName))
            {
                MessageBox.Show("Team name cannot be empty. Please enter a valid name.", "Invalid Team Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Capitalize first letter of each word
            _newTeamName = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_newTeamName.ToLower());

            Settings.Teams.Add(new Team
            {
                Name = _newTeamName,
                Members = new List<string>(),
                IncludeInMetrics = true,
                IsDepartment = false,
                HideTeam = false
            });

            Settings.Save();

            NewTeamNameTextBox.Text = "";
            HintLabel.Visibility = Visibility.Visible;

            RefreshTeamsList();
        }

        private void ShowHiddenButton_Click(object sender, RoutedEventArgs e)
        {
            _showHiddenTeams = !_showHiddenTeams;
            if (_showHiddenTeams)
            {
                // #FF42C7FF
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF42C7FF"));
                HideButton.BorderBrush = brush;
            }
            else
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF828282"));
                HideButton.BorderBrush = brush;
            }
            RefreshTeamsList();
        }

        private void NewTeamNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _newTeamName = NewTeamNameTextBox.Text.Trim();

            if (_newTeamName.Length > 0)
            {
                HintLabel.Visibility = Visibility.Collapsed;
            }
            else
            {
                HintLabel.Visibility = Visibility.Visible;
            }
        }

        private void DeleteTeam(string teamName)
        {
            var Teams = Settings.Teams;

            var teamToDelete = Settings.Teams.FirstOrDefault(t => t.Name == teamName);
            if (!teamToDelete.Equals(default(SettingsData)))
            {
                if (teamToDelete.Members.Count > 0)
                {
                    var result = MessageBox.Show($"Are you sure you want to delete the team '{teamName}'?", "Confirm Deletion", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        Settings.Teams.Remove(teamToDelete);
                        Settings.Save();
                        RefreshTeamsList();
                    }
                }
                else
                {
                    Settings.Teams.Remove(teamToDelete);
                    Settings.Save();
                    RefreshTeamsList();
                }
            }
        }

        public void RefreshTeamsList()
        {
            TeamsListBox.Children.Clear();
            Settings.Load();

            var teamsCopy = Settings.Teams.ToList();

            foreach (var team in teamsCopy)
            {
                if (team.HideTeam && !_showHiddenTeams)
                    continue;

                var teamControl = new Controls.TeamControl(team.Name);

                teamControl.DeleteTeamClicked += (sender, e) => DeleteTeam(teamControl.TeamName);
                teamControl.RefreshRequest += (s, e) => RefreshTeamsList();
                TeamsListBox.Children.Add(teamControl);
            }
        }

        public void RefreshRepsList()
        {
            LooseRepsPanel.Children.Clear();
            if(_reps == null || _reps.Count == 0)
            {
                // show warning card
                var warningCard = new RepWarningCard();
                LooseRepsPanel.Children.Add(warningCard);
            }

            foreach (var rep in _reps)
            {
                if (!Settings.Teams.Any(t => t.Members.Contains(rep.Name)))
                {
                    // fixed a bug that included average user after report generation
                    if (rep.Name == "-- AVERAGE --") continue;

                    var repItem = new Controls.RepItem(rep.Name);
                    LooseRepsPanel.Children.Add(repItem);
                }
            }
        }

        private void FreeRepsWrapPanel_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(RepItem)))
                e.Effects = DragDropEffects.Move;   // <--- REQUIRED
            else
                e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void FreeRepsWrapPanel_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(RepItem)) is RepItem item)
            {
                if (item.Parent is ListBox parentBox)
                {
                    parentBox.Items.Remove(item);
                    this.RemoveLogicalChild(item);
                }

                if(item.Parent is Panel parentPanel)
                {
                    parentPanel.Children.Remove(item);
                    this.RemoveLogicalChild(item);
                }

                foreach (var kvp in Settings.Teams.Where(k => k.Members.Contains(item.RepName)))
                {
                    kvp.Members.Remove(item.RepName);
                }

                var newItem = new Controls.RepItem(item.RepName);

                LooseRepsPanel.Children.Add(newItem);
                RefreshRepsList();
            }
        }

        private void NewTeamNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddNewTeam_Click(sender, e);
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TeamsWindow_SourceInitialized(object sender, EventArgs e)
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
