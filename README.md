# CallMetrics

CallMetrics is a WPF desktop application built with C# (.NET 8) for importing, analyzing, and reporting call data from Nextiva CSV reports. The application provides a user-friendly interface for call center managers and teams to visualize, organize, and export key metrics about representative performance.

## Features

- **Import Nextiva Reports:** Easily select and import CSV files containing call data.
- **Data Visualization:** View detailed metrics for each representative, including total calls, durations, inbound/outbound breakdowns, weekend/internal calls, and more.
- **Team Assignment:** Organize representatives into teams for more granular reporting.
- **Report Generation:** Export analyzed metrics to Excel files for further review or sharing.
- **Customizable Output Directory:** Save generated reports to your preferred location. These reports include Formulas for manual inputting additional data.
- **Clear Data:** Reset imported data to start fresh.
- **Modern WPF UI:** Intuitive, drag-movable window with custom resizing and controls.

## Getting Started

### Prerequisites

- **Windows 10/11**
- **Microsoft Excel Locally Installed on your Machine**
- **.NET 8 SDK**
- **Visual Studio 2022** (recommended for development)
- **Nextiva CSV Reports** (for importing call data)

### Building and Running

1. **Clone the repository:**
 `git clone https://github.com/yourusername/CallMetrics.git`
2. 2. **Open the solution in Visual Studio 2022.**
3. **Restore NuGet packages** (if prompted).
4. **Build the solution.**
5. **Run the application (F5 or Ctrl+F5).**

### Usage

1. **Import a Report:** Click the "Import Nextiva Report" button and select a CSV file.
2. **View Data:** Imported data will be displayed in a grid, showing metrics for each representative.
3. **Assign Teams:** Use the "Set Reps to Teams" button to organize reps into teams.
4. **Generate Report:** Click "Generate Report" to export the metrics to an Excel file.
5. **Clear Data:** Use "Clear Data" to remove all imported information and start over.

## Project Structure

- **MainWindow.xaml / MainWindow.xaml.cs:** Main application window and logic.
- **Controllers:** Handles data import (`NextivaReportReader`), report generation (`MetricsReport`), and other core operations.
- **Models:** Data structures for representatives (`RepData`), calls (`CallData`), and related metrics.
- **Menus:** UI components for team assignment and settings.
- **Utilities:** Helper classes for settings, file operations, and other utilities.

## Key Classes

- **NextivaReportReader:** Parses CSV files and extracts call data.
- **RepData:** Stores metrics and call details for each representative.
- **MetricsReport:** Generates Excel reports from analyzed data.
- **CallData:** Represents individual call records.

## Customization

- **Team Management:** Assign and manage teams for more detailed reporting.

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes, enhancements, or new features.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Support

For questions or support, please open an issue on the GitHub repository.
   