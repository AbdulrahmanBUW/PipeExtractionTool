using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PipeExtractionTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class PipeExtractionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                if (doc == null)
                {
                    TaskDialog.Show("Error", "No active document found.");
                    return Result.Failed;
                }

                // Get all sheets/drawings in the document
                List<DrawingSheetInfo> drawingSheets = GetDrawingSheets(doc);

                if (!drawingSheets.Any())
                {
                    TaskDialog.Show("Information", "No drawing sheets found in the current document.");
                    return Result.Succeeded;
                }

                // Filter sheets that belong to DEAXO_My (assuming this is in the sheet name or parameter)
                List<DrawingSheetInfo> deaxoSheets = FilterDeaxoSheets(drawingSheets);

                if (!deaxoSheets.Any())
                {
                    TaskDialog.Show("Information", "No DEAXO_My drawing sheets found in the current document.");
                    return Result.Succeeded;
                }

                // Show UI for sheet selection
                PipeExtractionWindow window = new PipeExtractionWindow(deaxoSheets, doc);
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                bool? result = window.ShowDialog();

                if (result == true)
                {
                    return Result.Succeeded;
                }

                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
                return Result.Failed;
            }
        }

        private List<DrawingSheetInfo> GetDrawingSheets(Document doc)
        {
            List<DrawingSheetInfo> sheets = new List<DrawingSheetInfo>();

            // Get all ViewSheet elements
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet));

            foreach (ViewSheet sheet in collector.Cast<ViewSheet>())
            {
                // Skip placeholder sheets if possible (some versions don't have this property)
                try
                {
                    if (sheet.IsPlaceholderSheet)
                        continue;
                }
                catch
                {
                    // Property not available in this version, continue anyway
                }

                DrawingSheetInfo sheetInfo = new DrawingSheetInfo
                {
                    Id = sheet.Id,
                    Name = sheet.Name,
                    Number = sheet.SheetNumber,
                    Title = sheet.Title,
                    ViewSheet = sheet
                };

                sheets.Add(sheetInfo);
            }

            return sheets;
        }

        private List<DrawingSheetInfo> FilterDeaxoSheets(List<DrawingSheetInfo> allSheets)
        {
            // Filter sheets that contain "DEAXO_My" in their name, number, or title
            return allSheets.Where(sheet =>
                (sheet.Name?.Contains("DEAXO_My") == true) ||
                (sheet.Number?.Contains("DEAXO_My") == true) ||
                (sheet.Title?.Contains("DEAXO_My") == true) ||
                (sheet.Number?.Contains("DEAXO") == true) // More flexible matching
            ).ToList();
        }
    }

    public class DrawingSheetInfo
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
        public string Title { get; set; }
        public ViewSheet ViewSheet { get; set; }
        public bool IsSelected { get; set; } = false;

        public string DisplayName => $"{Number} - {Name}";

        public override string ToString()
        {
            return DisplayName;
        }
    }
}