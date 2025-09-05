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
        private BackgroundWorker _backgroundWorker;

        public PipeExtractionWindow(List<DrawingSheetInfo> drawingSheets, Document document)
        {
            InitializeComponent();
            _drawingSheets = drawingSheets;
            _document = document;

            InitializeUI();
            SetupBackgroundWorker();
        }

        private void InitializeUI()
        {
            SheetsListView.ItemsSource = _drawingSheets;

            // Pre-select all sheets
            foreach (var sheet in _drawingSheets)
            {
                sheet.IsSelected = true;
            }

            SheetsListView.Items.Refresh();
        }

        private void SetupBackgroundWorker()
        {
            _backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            _backgroundWorker.DoWork += BackgroundWorker_DoWork;
            _backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in _drawingSheets)
            {
                sheet.IsSelected = true;
            }
            SheetsListView.Items.Refresh();
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in _drawingSheets)
            {
                sheet.IsSelected = false;
            }
            SheetsListView.Items.Refresh();
        }

        private void ExtractButton_Click(object sender, RoutedEventArgs e)
        {
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
                FileName = $"DEAXO_Pipe_Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                StartExtraction(selectedSheets, saveFileDialog.FileName);
            }
        }

        private void StartExtraction(List<DrawingSheetInfo> selectedSheets, string filePath)
        {
            // Disable UI
            ExtractButton.IsEnabled = false;
            SelectAllButton.IsEnabled = false;
            DeselectAllButton.IsEnabled = false;
            SheetsListView.IsEnabled = false;

            ProgressBar.Visibility = Visibility.Visible;
            ProgressText.Text = "Starting extraction...";

            // Start background work
            var args = new ExtractionArgs
            {
                SelectedSheets = selectedSheets,
                FilePath = filePath,
                Document = _document
            };

            _backgroundWorker.RunWorkerAsync(args);
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var args = (ExtractionArgs)e.Argument;
            var worker = (BackgroundWorker)sender;

            try
            {
                var extractor = new PipeDataExtractor();
                var result = extractor.ExtractPipeData(args.Document, args.SelectedSheets, worker);

                worker.ReportProgress(90, "Generating Excel file...");

                var excelExporter = new ExcelExporter();
                excelExporter.ExportToExcel(result, args.FilePath);

                worker.ReportProgress(100, "Extraction completed successfully!");

                e.Result = new ExtractionResult
                {
                    Success = true,
                    FilePath = args.FilePath,
                    ExtractedData = result
                };
            }
            catch (Exception ex)
            {
                e.Result = new ExtractionResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;

            if (e.UserState != null)
            {
                ProgressText.Text = e.UserState.ToString();
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Re-enable UI
            ExtractButton.IsEnabled = true;
            SelectAllButton.IsEnabled = true;
            DeselectAllButton.IsEnabled = true;
            SheetsListView.IsEnabled = true;

            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressText.Text = "";

            if (e.Result is ExtractionResult result)
            {
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
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_backgroundWorker?.IsBusy == true)
            {
                _backgroundWorker.CancelAsync();
            }

            DialogResult = false;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_backgroundWorker?.IsBusy == true)
            {
                _backgroundWorker.CancelAsync();
            }

            base.OnClosing(e);
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