using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Microsoft.Win32;

namespace PipeExtractionTool
{
    public partial class PipeExtractionWindow : Window
    {
        private List<DrawingSheetInfo> _drawingSheets;
        private Document _document;
        private List<string> _disciplines;
        private List<DrawingSheetInfo> _currentDisplayedSheets;

        public PipeExtractionWindow(List<DrawingSheetInfo> drawingSheets, Document document)
        {
            InitializeComponent();
            _drawingSheets = drawingSheets;
            _document = document;
            _currentDisplayedSheets = drawingSheets; // Initially show all sheets

            InitializeUI();
            LoadDisciplines();
        }

        private void InitializeUI()
        {
            SheetsListView.ItemsSource = _currentDisplayedSheets;

            // Pre-select all sheets
            foreach (var sheet in _currentDisplayedSheets)
            {
                sheet.IsSelected = true;
            }

            SheetsListView.Items.Refresh();
        }

        private void LoadDisciplines()
        {
            _disciplines = new List<string>();

            try
            {
                // Get all unique disciplines
                _disciplines = _drawingSheets
                    .Select(s => s.Discipline)
                    .Distinct()
                    .OrderBy(d => d)
                    .ToList();

                // Add "All Disciplines" option
                _disciplines.Insert(0, "All Disciplines");

                // Populate the combo box
                DisciplinesComboBox.ItemsSource = _disciplines;
                DisciplinesComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading disciplines: {ex.Message}");
            }
        }

        private void DisciplinesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Optional: Auto-filter when a discipline is selected
        }

        private void SelectDisciplineButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedDiscipline = DisciplinesComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(selectedDiscipline))
            {
                MessageBox.Show("Please select a discipline first.", "No Selection",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                if (selectedDiscipline == "All Disciplines")
                {
                    // Show all sheets
                    _currentDisplayedSheets = _drawingSheets;
                }
                else
                {
                    // Filter sheets by discipline
                    _currentDisplayedSheets = _drawingSheets.Where(s => s.Discipline == selectedDiscipline).ToList();
                }

                SheetsListView.ItemsSource = _currentDisplayedSheets;
                SheetsListView.Items.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error filtering by discipline: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearFilterButton_Click(object sender, RoutedEventArgs e)
        {
            // Show all sheets
            _currentDisplayedSheets = _drawingSheets;
            SheetsListView.ItemsSource = _currentDisplayedSheets;
            SheetsListView.Items.Refresh();
            DisciplinesComboBox.SelectedIndex = 0;
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            // Select only the sheets currently displayed in the list view
            foreach (var sheet in _currentDisplayedSheets)
            {
                sheet.IsSelected = true;
            }
            SheetsListView.Items.Refresh();
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            // Deselect only the sheets currently displayed in the list view
            foreach (var sheet in _currentDisplayedSheets)
            {
                sheet.IsSelected = false;
            }
            SheetsListView.Items.Refresh();
        }

        private async void ExtractButton_Click(object sender, RoutedEventArgs e)
        {
            // Get selected sheets from the original list, not just the displayed ones
            var selectedSheets = _drawingSheets.Where(s => s.IsSelected).ToList();

            if (!selectedSheets.Any())
            {
                MessageBox.Show("Please select at least one drawing sheet.", "No Selection",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Show save file dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Save Pipe Extraction Report",
                FileName = $"Pipe_Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                await StartExtractionAsync(selectedSheets, saveFileDialog.FileName);
            }
        }

        private async Task StartExtractionAsync(List<DrawingSheetInfo> selectedSheets, string filePath)
        {
            // Disable UI
            SetUIEnabled(false);
            ProgressBar.Visibility = System.Windows.Visibility.Visible;
            ProgressText.Text = "Starting extraction...";

            try
            {
                // Create a simple progress reporter that updates the UI
                var progressReporter = new Progress<(int percentage, string message)>(report =>
                {
                    ProgressBar.Value = report.percentage;
                    ProgressText.Text = report.message;
                });

                // Run extraction on UI thread with progress updates
                var result = await Task.Run(() => ExtractDataWithProgress(selectedSheets, filePath, progressReporter));

                if (result.Success)
                {
                    string message = $"Extraction completed successfully!\n\nFile saved to:\n{result.FilePath}\n\nTotal sheets processed: {result.ExtractedData?.Count ?? 0}";
                    MessageBox.Show(message, "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Ask if user wants to open the file
                    var openResult = MessageBox.Show("Would you like to open the Excel file now?",
                                                    "Open File", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (openResult == MessageBoxResult.Yes)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(result.FilePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Could not open file: {ex.Message}", "Error",
                                          MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }

                    DialogResult = true;
                }
                else
                {
                    MessageBox.Show($"Extraction failed: {result.ErrorMessage}", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Extraction failed: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Re-enable UI
                SetUIEnabled(true);
                ProgressBar.Visibility = System.Windows.Visibility.Collapsed;
                ProgressText.Text = "";
            }
        }

        private ExtractionResult ExtractDataWithProgress(List<DrawingSheetInfo> selectedSheets, string filePath, IProgress<(int, string)> progress)
        {
            try
            {
                // IMPORTANT: Use Dispatcher.Invoke to run extraction on UI thread
                return Dispatcher.Invoke(() =>
                {
                    progress?.Report((10, "Initializing extraction..."));

                    var extractor = new PipeDataExtractor();

                    // Create a simple progress wrapper for the extractor
                    var extractorProgress = new SimpleProgressReporter(progress);

                    var extractedData = extractor.ExtractPipeData(_document, selectedSheets, extractorProgress);

                    progress?.Report((90, "Generating Excel file..."));

                    var excelExporter = new ExcelExporter();
                    excelExporter.ExportToExcel(extractedData, filePath);

                    progress?.Report((100, "Extraction completed successfully!"));

                    return new ExtractionResult
                    {
                        Success = true,
                        FilePath = filePath,
                        ExtractedData = extractedData
                    };
                });
            }
            catch (Exception ex)
            {
                return new ExtractionResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private void SetUIEnabled(bool enabled)
        {
            ExtractButton.IsEnabled = enabled;
            SelectAllButton.IsEnabled = enabled;
            DeselectAllButton.IsEnabled = enabled;
            SheetsListView.IsEnabled = enabled;
            DisciplinesComboBox.IsEnabled = enabled;
            SelectDisciplineButton.IsEnabled = enabled;
            ClearFilterButton.IsEnabled = enabled;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }

    // Simple progress reporter that bridges between IProgress<T> and BackgroundWorker style reporting
    public class SimpleProgressReporter
    {
        private readonly IProgress<(int, string)> _progress;
        private int _totalSheets;
        private int _processedSheets;

        public SimpleProgressReporter(IProgress<(int, string)> progress)
        {
            _progress = progress;
        }

        public bool CancellationPending => false; // No cancellation support for now

        public void SetTotalSheets(int total)
        {
            _totalSheets = total;
            _processedSheets = 0;
        }

        public void ReportProgress(int percentage, string message)
        {
            _progress?.Report((percentage, message));
        }

        public void IncrementSheet()
        {
            _processedSheets++;
            if (_totalSheets > 0)
            {
                int percentage = (int)((double)_processedSheets / _totalSheets * 80); // Leave 20% for Excel export
                _progress?.Report((percentage, $"Processed {_processedSheets}/{_totalSheets} sheets"));
            }
        }
    }

    public class ExtractionArgs
    {
        public List<DrawingSheetInfo> SelectedSheets { get; set; }
        public string FilePath { get; set; }
        public Document Document { get; set; }
    }

    public class ExtractionResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string FilePath { get; set; }
        public List<SheetPipeData> ExtractedData { get; set; }
    }
}