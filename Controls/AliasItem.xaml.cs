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
using System.Xml.Serialization;

namespace CallMetrics.Controls
{
    /// <summary>
    /// Interaction logic for AliasItem.xaml
    /// </summary>
    public partial class AliasItem : UserControl
    {
        public string AliasName;

        public AliasItem()
        {
            InitializeComponent();
        }

        public AliasItem(Alias alias)
        {
            InitializeComponent();
            AliasName = alias.Name;

            AliasTextBox.Text = AliasName;
            AliasTextBox.TextChanged += AliasTextBox_TextChanged;
            RefreshAliasList();
        }

        private void RefreshAliasList()
        {
            Alias alias = Settings.Aliases.Where(t => t.Name == AliasName).FirstOrDefault();
            if (alias.IsNull())
                return;

            RepsPanel.Children.Clear();
            foreach (var repName in alias.AliasedTo)
            {
                var repItem = new RepItem(repName);
                RepsPanel.Children.Add(repItem);
            }
        }

        private void AliasTextBox_TextChanged(object sender, RoutedEventArgs e)
        {
            var existingAlias = Settings.Aliases.FirstOrDefault(a => a.Name == AliasName);
            if (!existingAlias.IsNull())
            {
                // Update the alias name
                existingAlias.Name = AliasTextBox.Text.Trim();
                AliasName = existingAlias.Name;
                Settings.Save();
            }
        }

        private void DeleteAliasButton_Click(object sender, RoutedEventArgs e)
        {
            Settings.Aliases.RemoveAll(a => a.Name == AliasName);
            Settings.Save();
            
            // just remove this item from its parent
            if (this.Parent is Panel parentPanel)
            {
                parentPanel.Children.Remove(this);
                this.RemoveLogicalChild(this);
            }

            if (this.Parent is ItemsControl parentItemsControl)
            {
                parentItemsControl.Items.Remove(this);
                this.RemoveLogicalChild(this);
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
            Alias alias = Settings.Aliases.Where(t => t.Name == AliasName).FirstOrDefault();
            if (alias.IsNull())
                return;

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

                alias.AliasedTo.Add(item.RepName);
                Settings.Save();

                var newItem = new RepItem(item.RepName);
                (sender as ListBox)?.Items.Add(newItem);
            }

            RefreshAliasList();
        }

        private void RootDragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            e.Handled = false; // <--- Important! Do NOT swallow it
        }
    }
}
