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

                // Show UI for sheet selection
                PipeExtractionWindow window = new PipeExtractionWindow(drawingSheets, doc);
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
                // Skip placeholder sheets if possible using reflection
                try
                {
                    var isPlaceholderProp = sheet.GetType().GetProperty("IsPlaceholderSheet");
                    if (isPlaceholderProp != null && (bool)isPlaceholderProp.GetValue(sheet))
                        continue;
                }
                catch
                {
                    // Property not available or accessible, continue anyway
                }

                // Get discipline parameter
                string discipline = GetSheetDiscipline(sheet);

                DrawingSheetInfo sheetInfo = new DrawingSheetInfo
                {
                    Id = sheet.Id,
                    Name = sheet.Name,
                    Number = sheet.SheetNumber,
                    Title = sheet.Title,
                    Discipline = discipline,
                    ViewSheet = sheet
                };

                sheets.Add(sheetInfo);
            }

            return sheets;
        }

        private void DebugSheetParameters(ViewSheet sheet)
        {
            try
            {
                TaskDialog.Show("Debug", $"Listing parameters for sheet: {sheet.SheetNumber}");

                string parameterInfo = "Parameters:\n";
                foreach (Parameter param in sheet.Parameters)
                {
                    string value = "No value";
                    if (param.HasValue)
                    {
                        try
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    value = param.AsString();
                                    break;
                                case StorageType.Integer:
                                    value = param.AsInteger().ToString();
                                    break;
                                case StorageType.Double:
                                    value = param.AsDouble().ToString();
                                    break;
                                case StorageType.ElementId:
                                    ElementId id = param.AsElementId();
                                    if (id != ElementId.InvalidElementId)
                                    {
                                        Element elem = sheet.Document.GetElement(id);
                                        value = elem?.Name ?? "Unknown Element";
                                    }
                                    else
                                    {
                                        value = "Invalid ElementId";
                                    }
                                    break;
                                default:
                                    value = "Unknown type";
                                    break;
                            }
                        }
                        catch
                        {
                            value = "Error reading value";
                        }
                    }

                    parameterInfo += $"{param.Definition.Name}: {value}\n";
                }

                TaskDialog.Show("Sheet Parameters", parameterInfo);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to get parameters: {ex.Message}");
            }
        }

        private string GetSheetDiscipline(ViewSheet sheet)
        {
            try
            {
                // Try to get the discipline parameter by name (shared parameter)
                string disciplineParameterName = "000_000_150_Discipline";
                Parameter disciplineParam = sheet.LookupParameter(disciplineParameterName);

                if (disciplineParam != null && disciplineParam.HasValue)
                {
                    if (disciplineParam.StorageType == StorageType.String)
                    {
                        return disciplineParam.AsString();
                    }
                    else if (disciplineParam.StorageType == StorageType.Integer)
                    {
                        return disciplineParam.AsInteger().ToString();
                    }
                    else if (disciplineParam.StorageType == StorageType.Double)
                    {
                        return disciplineParam.AsDouble().ToString();
                    }
                    else if (disciplineParam.StorageType == StorageType.ElementId)
                    {
                        ElementId id = disciplineParam.AsElementId();
                        if (id != ElementId.InvalidElementId)
                        {
                            Element elem = sheet.Document.GetElement(id);
                            return elem?.Name ?? "Unknown";
                        }
                    }
                }

                // Try alternative parameter names if the specific one doesn't work
                string[] alternativeParams = {
                    "000_000_152_Discipline", // Try the previous name too
                    "Discipline",
                    "DISCIPLINE",
                    "Sheet Discipline",
                    "Sheet_Discipline"
                };

                foreach (string paramName in alternativeParams)
                {
                    Parameter altParam = sheet.LookupParameter(paramName);
                    if (altParam != null && altParam.HasValue)
                    {
                        if (altParam.StorageType == StorageType.String)
                        {
                            return altParam.AsString();
                        }
                        else if (altParam.StorageType == StorageType.Integer)
                        {
                            return altParam.AsInteger().ToString();
                        }
                        else if (altParam.StorageType == StorageType.Double)
                        {
                            return altParam.AsDouble().ToString();
                        }
                        else if (altParam.StorageType == StorageType.ElementId)
                        {
                            ElementId id = altParam.AsElementId();
                            if (id != ElementId.InvalidElementId)
                            {
                                Element elem = sheet.Document.GetElement(id);
                                return elem?.Name ?? "Unknown";
                            }
                        }
                    }
                }

                // Debug: List all parameters to help identify the correct one
                System.Diagnostics.Debug.WriteLine($"Parameters for sheet {sheet.SheetNumber}:");
                foreach (Parameter param in sheet.Parameters)
                {
                    System.Diagnostics.Debug.WriteLine($"  {param.Definition.Name}: {param.AsValueString()}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting discipline parameter: {ex.Message}");
            }

            return "Unknown";
        }
    }
}