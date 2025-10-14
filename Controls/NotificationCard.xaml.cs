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
    /// Interaction logic for NotificationCard.xaml
    /// </summary>
    public partial class NotificationCard : UserControl
    {

        public NotificationCard()
        {
            InitializeComponent();
        }

        public NotificationCard(Notification noti)
        {
            InitializeComponent();

            TitleTextBlock.Text = noti.Title;
            MessageTextBlock.Text = noti.Message;

            FadeIn();
            BeginDeath();
        }

        private void BeginDeath()
        {
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(5);
            timer.Tick += (s, e) =>
            {
                FadeOut();
                timer.Stop();
            };
            timer.Start();
        }

        private void FadeIn()
        {
            var fadeInAnimation = new System.Windows.Media.Animation.DoubleAnimation(0, 1, new Duration(TimeSpan.FromSeconds(0.5)));
            this.BeginAnimation(UserControl.OpacityProperty, fadeInAnimation);
        }

        private void FadeOut()
        {
            var fadeOutAnimation = new System.Windows.Media.Animation.DoubleAnimation(1, 0, new Duration(TimeSpan.FromSeconds(0.5)));
            fadeOutAnimation.Completed += OnFadeOutFinished;
            this.BeginAnimation(UserControl.OpacityProperty, fadeOutAnimation);
        }

        private void OnFadeOutFinished(object sender, EventArgs e)
        {
            var parent = this.Parent as Panel;
            if (parent != null)
            {
                parent.Children.Remove(this);
            }
        }
    }
}
