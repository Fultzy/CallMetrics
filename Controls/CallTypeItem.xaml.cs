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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CallMetrics.Controls
{
    /// <summary>
    /// Interaction logic for CallTypeItem.xaml
    /// </summary>
    public partial class CallTypeItem : UserControl
    {
        public string CallType;
        public bool IsInbound;

        public CallTypeItem()
        {
            InitializeComponent();
        }

        public CallTypeItem(string callType, bool isInbound)
        {
            InitializeComponent();

            CallType = callType;
            IsInbound = isInbound;

            // capitalize first letter of each word in callType
            var capitalizedCallType = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(callType.ToLower());

            CallTypeTextBlock.Text = capitalizedCallType;
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var parent = this.Parent as Panel;
            if (parent != null)
            {
                if (IsInbound)
                {
                    Settings.InboundCallTypes.Remove(CallType);
                    Settings.Save();
                }
                else
                {
                    Settings.OutboundCallTypes.Remove(CallType);
                    Settings.Save();
                }
                
                parent.Children.Remove(this);
            }
        }
    }
}
