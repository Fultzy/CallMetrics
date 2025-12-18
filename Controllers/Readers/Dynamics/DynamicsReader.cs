using CallMetrics.Models;
using CallMetrics.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CallMetrics.Controllers.Readers.Dynamics
{
    internal class DynamicsReader
    {
        internal async Task<List<Ticket>> Start()
        {
            // open explorer to select file
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.DefaultDirectory = Settings.DefaultReportPath;
            openFileDialog.DefaultExt = ".csv";
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string filePath = openFileDialog.FileName;
                return await Task.Run(() => Read(filePath));
            }

            return new List<Ticket>();
        }

        internal List<Ticket> Read(string filePath)
        {
            try
            {
                throw new NotImplementedException();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading Dynamics report: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<Ticket>();
            }
        }
    }
}
