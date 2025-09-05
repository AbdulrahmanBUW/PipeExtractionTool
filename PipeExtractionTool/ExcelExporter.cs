using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PipeExtractionTool
{
    public class ExcelExporter
    {
        public void ExportToExcel(List<SheetPipeData> data, string filePath)
        {

            try
            {
                using (var package = new ExcelPackage())
                {
                    // Create worksheet
                    var worksheet = package.Workbook.Worksheets.Add("DEAXO Pipe Report");

                    // Set up headers
                    worksheet.Cells[1, 1].Value = "Drawing Name";
                    worksheet.Cells[1, 2].Value = "Pipes";

                    // Style headers
                    using (var headerRange = worksheet.Cells[1, 1, 1, 2])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Font.Size = 12;
                        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                        headerRange.Style.Font.Color.SetColor(Color.White);
                        headerRange.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // Add data
                    int row = 2;
                    foreach (var sheetData in data)
                    {
                        worksheet.Cells[row, 1].Value = sheetData.SheetName;
                        worksheet.Cells[row, 2].Value = sheetData.PipeSpecPositionsString;

                        // Style data rows
                        using (var dataRange = worksheet.Cells[row, 1, row, 2])
                        {
                            dataRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                            // Alternate row coloring
                            if (row % 2 == 0)
                            {
                                dataRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                dataRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                            }
                        }

                        // Wrap text for pipe column
                        worksheet.Cells[row, 2].Style.WrapText = true;

                        row++;
                    }

                    // Auto-fit columns
                    worksheet.Column(1).AutoFit();

                    // Set specific width for pipes column to handle long lists better
                    worksheet.Column(2).Width = 50;

                    // Set minimum height for data rows to accommodate wrapped text
                    for (int r = 2; r < row; r++)
                    {
                        worksheet.Row(r).Height = Math.Max(15, worksheet.Row(r).Height);
                    }

                    // Add summary information
                    int summaryStartRow = row + 2;

                    worksheet.Cells[summaryStartRow, 1].Value = "Summary";
                    worksheet.Cells[summaryStartRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[summaryStartRow, 1].Style.Font.Size = 14;

                    worksheet.Cells[summaryStartRow + 1, 1].Value = "Total Drawings Processed:";
                    worksheet.Cells[summaryStartRow + 1, 2].Value = data.Count;

                    // Count unique pipes across all drawings
                    var allPipes = new HashSet<string>();
                    foreach (var sheetData in data)
                    {
                        foreach (var pipe in sheetData.PipeSpecPositions)
                        {
                            allPipes.Add(pipe);
                        }
                    }

                    worksheet.Cells[summaryStartRow + 2, 1].Value = "Total Unique Pipes:";
                    worksheet.Cells[summaryStartRow + 2, 2].Value = allPipes.Count;

                    worksheet.Cells[summaryStartRow + 3, 1].Value = "Report Generated:";
                    worksheet.Cells[summaryStartRow + 3, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // Style summary section
                    using (var summaryRange = worksheet.Cells[summaryStartRow + 1, 1, summaryStartRow + 3, 2])
                    {
                        summaryRange.Style.Font.Size = 10;
                        summaryRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Add freeze panes to keep headers visible
                    worksheet.View.FreezePanes(2, 1);

                    // Add auto filter to headers
                    worksheet.Cells[1, 1, row - 1, 2].AutoFilter = true;

                    // Save the file
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to Excel: {ex.Message}", ex);
            }
        }

        public void ExportDetailedReport(List<SheetPipeData> data, string filePath)
        {
            // EPPlus 4.5.3.3 doesn't require license context setting
            try
            {
                using (var package = new ExcelPackage())
                {
                    // Create main summary worksheet
                    CreateSummaryWorksheet(package, data);

                    // Create detailed worksheet with one row per pipe per drawing
                    CreateDetailedWorksheet(package, data);

                    // Create unique pipes list
                    CreateUniquePipesWorksheet(package, data);

                    // Save the file
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export detailed report to Excel: {ex.Message}", ex);
            }
        }

        private void CreateSummaryWorksheet(ExcelPackage package, List<SheetPipeData> data)
        {
            var worksheet = package.Workbook.Worksheets.Add("Summary");

            // This is the same as the main export method
            // Headers
            worksheet.Cells[1, 1].Value = "Drawing Name";
            worksheet.Cells[1, 2].Value = "Pipes";

            // Style headers
            StyleHeaders(worksheet, 1, 1, 1, 2);

            // Add data
            int row = 2;
            foreach (var sheetData in data)
            {
                worksheet.Cells[row, 1].Value = sheetData.SheetName;
                worksheet.Cells[row, 2].Value = sheetData.PipeSpecPositionsString;
                StyleDataRow(worksheet, row, 2);
                row++;
            }

            // Auto-fit and formatting
            worksheet.Column(1).AutoFit();
            worksheet.Column(2).Width = 50;
            worksheet.View.FreezePanes(2, 1);
            worksheet.Cells[1, 1, row - 1, 2].AutoFilter = true;
        }

        private void CreateDetailedWorksheet(ExcelPackage package, List<SheetPipeData> data)
        {
            var worksheet = package.Workbook.Worksheets.Add("Detailed");

            // Headers
            worksheet.Cells[1, 1].Value = "Drawing Name";
            worksheet.Cells[1, 2].Value = "Pipe Spec Position";

            StyleHeaders(worksheet, 1, 1, 1, 2);

            // Add detailed data (one row per pipe per drawing)
            int row = 2;
            foreach (var sheetData in data)
            {
                if (sheetData.PipeSpecPositions.Count == 0)
                {
                    // If no pipes, still show the drawing
                    worksheet.Cells[row, 1].Value = sheetData.SheetName;
                    worksheet.Cells[row, 2].Value = "(No pipes found)";
                    StyleDataRow(worksheet, row, 2);
                    row++;
                }
                else
                {
                    foreach (var pipe in sheetData.PipeSpecPositions)
                    {
                        worksheet.Cells[row, 1].Value = sheetData.SheetName;
                        worksheet.Cells[row, 2].Value = pipe;
                        StyleDataRow(worksheet, row, 2);
                        row++;
                    }
                }
            }

            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.View.FreezePanes(2, 1);
            worksheet.Cells[1, 1, row - 1, 2].AutoFilter = true;
        }

        private void CreateUniquePipesWorksheet(ExcelPackage package, List<SheetPipeData> data)
        {
            var worksheet = package.Workbook.Worksheets.Add("Unique Pipes");

            // Get all unique pipes with their occurrence count
            var pipeOccurrences = new Dictionary<string, int>();
            var pipeDrawings = new Dictionary<string, List<string>>();

            foreach (var sheetData in data)
            {
                foreach (var pipe in sheetData.PipeSpecPositions)
                {
                    if (!pipeOccurrences.ContainsKey(pipe))
                    {
                        pipeOccurrences[pipe] = 0;
                        pipeDrawings[pipe] = new List<string>();
                    }
                    pipeOccurrences[pipe]++;
                    pipeDrawings[pipe].Add(sheetData.SheetName);
                }
            }

            // Headers
            worksheet.Cells[1, 1].Value = "Pipe Spec Position";
            worksheet.Cells[1, 2].Value = "Occurrence Count";
            worksheet.Cells[1, 3].Value = "Found in Drawings";

            StyleHeaders(worksheet, 1, 1, 1, 3);

            // Add data
            int row = 2;
            foreach (var kvp in pipeOccurrences.OrderBy(p => p.Key))
            {
                worksheet.Cells[row, 1].Value = kvp.Key;
                worksheet.Cells[row, 2].Value = kvp.Value;
                worksheet.Cells[row, 3].Value = string.Join(", ", pipeDrawings[kvp.Key]);

                StyleDataRow(worksheet, row, 3);
                worksheet.Cells[row, 3].Style.WrapText = true;
                row++;
            }

            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).Width = 60;
            worksheet.View.FreezePanes(2, 1);
            worksheet.Cells[1, 1, row - 1, 3].AutoFilter = true;
        }

        private void StyleHeaders(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            using (var headerRange = worksheet.Cells[fromRow, fromCol, toRow, toCol])
            {
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.Size = 12;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                headerRange.Style.Font.Color.SetColor(Color.White);
                headerRange.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
        }

        private void StyleDataRow(ExcelWorksheet worksheet, int row, int columnCount)
        {
            using (var dataRange = worksheet.Cells[row, 1, row, columnCount])
            {
                dataRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                if (row % 2 == 0)
                {
                    dataRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    dataRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                }
            }
        }
    }
}