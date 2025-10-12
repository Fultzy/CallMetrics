using CallMetrics.Models;
using System;
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
    /// Interaction logic for RepItem.xaml
    /// </summary>
    public partial class RepItem : UserControl
    {
        public string RepName;

        public RepItem()
        {
            InitializeComponent();
        }

        public RepItem(string repName)
        {
            InitializeComponent();

            if (repName != "None")
            {
                RepName = repName;
                RepNameLabel.Content = repName;
            }

            this.PreviewMouseMove += RepItem_PreviewMouseMove;
        }

        private void RepItem_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                //this.IsHitTestVisible = false; // <--- LET dragging pass through

                DragDrop.DoDragDrop(this, this, DragDropEffects.Move);

                //this.IsHitTestVisible = true; // <--- restore after
            }
        }
    }
}
