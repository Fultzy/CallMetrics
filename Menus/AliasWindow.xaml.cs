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
using static MaterialDesignThemes.Wpf.Theme;

namespace CallMetrics.Menus
{
    /// <summary>
    /// Interaction logic for AliasWindow.xaml
    /// </summary>
    public partial class AliasWindow : Window
    {
        private List<Rep> _reps;
        private bool _showHiddenAlias;
        private string _newAliasName;

        public AliasWindow(List<Rep> reps)
        {
            InitializeComponent();
            this.SourceInitialized += AliasWindow_SourceInitialized;
            this.Closed += (s, e) => Settings.Load();

            _reps = reps;
            RefreshAliasList();
            RefreshRepsList();
        }

        private void RefreshAliasList()
        {
            AliasesPanel.Children.Clear();
            foreach (var alias in Settings.Aliases.OrderBy(a => a.Name))
            {
                var aliasItem = new AliasItem(alias);
                AliasesPanel.Children.Add(aliasItem);
            }
        }

        private void RefreshRepsList()
        {
            LooseRepsPanel.Children.Clear();
            var aliasedReps = Settings.Aliases.SelectMany(a => a.AliasedTo).Distinct().ToList();
            var freeReps = _reps.Where(r => !aliasedReps.Contains(r.Name)).OrderBy(r => r.Name).ToList();
            foreach (var rep in freeReps)
            {
                var repItem = new RepItem(rep.Name);
                LooseRepsPanel.Children.Add(repItem);
            }
        }


        private void NewAliasButton_Click(object sender, RoutedEventArgs e)
        {
            if (Settings.Aliases.Any(t => t.Name.Equals(_newAliasName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("An Alias with this name already exists. Please choose a different name.", "Duplicate Alias Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_newAliasName))
            {
                MessageBox.Show("Alias name cannot be empty. Please enter a valid name.", "Invalid Alias Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Capitalize first letter of each word
            _newAliasName = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_newAliasName.ToLower());

            var newAlias = new Alias
            {
                Name = _newAliasName,
                AliasedTo = new List<string>()
            };

            Settings.Aliases.Add(newAlias);
            Settings.Save();

            var aliasItem = new AliasItem(newAlias);
            AliasesPanel.Children.Add(aliasItem);

            RefreshAliasList();
        }


        private void NewAliasTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _newAliasName = NewAliasTextBox.Text.Trim();

            if (_newAliasName.Length > 0)
            {
                HintLabel.Visibility = Visibility.Collapsed;
            }
            else
            {
                HintLabel.Visibility = Visibility.Visible;
            }
        }

        private void NewAliasTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                NewAliasButton_Click(sender, e);
            }
        }


        private void ShowHiddenButton_Click(object sender, RoutedEventArgs e)
        {
            _showHiddenAlias = !_showHiddenAlias;
            if (_showHiddenAlias)
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
            RefreshAliasList();
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

                if (item.Parent is Panel parentPanel)
                {
                    parentPanel.Children.Remove(item);
                    this.RemoveLogicalChild(item);
                }

                foreach (var kvp in Settings.Aliases.Where(k => k.AliasedTo.Contains(item.RepName)))
                {
                    kvp.AliasedTo.Remove(item.RepName);
                }

                var newItem = new Controls.RepItem(item.RepName);

                LooseRepsPanel.Children.Add(newItem);
                RefreshRepsList();
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




        private void AliasWindow_SourceInitialized(object sender, EventArgs e)
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
