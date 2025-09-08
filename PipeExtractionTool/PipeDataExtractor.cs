using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace PipeExtractionTool
{
    public class PipeDataExtractor
    {
        private const string SPEC_POSITION_PARAMETER = "SPEC_POSITION";
        public List<SheetPipeData> ExtractPipeData(Document doc, List<DrawingSheetInfo> selectedSheets, BackgroundWorker worker = null)
        {
            var result = new List<SheetPipeData>();
            int totalSheets = selectedSheets.Count;
            int processedSheets = 0;

            foreach (var sheetInfo in selectedSheets)
            {
                if (worker?.CancellationPending == true)
                    break;

                // Report progress
                int progressPercentage = (int)((double)processedSheets / totalSheets * 80);
                worker?.ReportProgress(progressPercentage, $"Processing sheet: {sheetInfo.DisplayName}");

                try
                {
                    var sheetPipeData = ExtractPipesFromSheet(doc, sheetInfo.ViewSheet);
                    if (sheetPipeData != null)
                    {
                        result.Add(sheetPipeData);
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue processing other sheets
                    System.Diagnostics.Debug.WriteLine($"Error processing sheet {sheetInfo.DisplayName}: {ex.Message}");
                }

                processedSheets++;
            }

            return result;
        }

        private SheetPipeData ExtractPipesFromSheet(Document doc, ViewSheet sheet)
        {
            var sheetPipeData = new SheetPipeData
            {
                SheetName = $"{sheet.SheetNumber} - {sheet.Name}",
                PipeSpecPositions = new List<string>()
            };

            try
            {
                // Get all viewports on the sheet
                var viewportIds = sheet.GetAllViewports();

                foreach (ElementId viewportId in viewportIds)
                {
                    if (!(doc.GetElement(viewportId) is Viewport viewport)) continue;

                    if (!(doc.GetElement(viewport.ViewId) is View view)) continue;

                    // Extract pipes from all view types that might contain pipes
                    if (view is ViewPlan || view is ViewSection || view is View3D)
                    {
                        var pipesInView = GetPipesFromView(doc, view);
                        foreach (var pipe in pipesInView)
                        {
                            string specPosition = GetSpecPositionParameter(pipe);
                            if (!string.IsNullOrEmpty(specPosition) &&
                                !sheetPipeData.PipeSpecPositions.Contains(specPosition))
                            {
                                sheetPipeData.PipeSpecPositions.Add(specPosition);
                            }
                        }
                    }
                }

                // Sort the spec positions for consistent output
                sheetPipeData.PipeSpecPositions.Sort();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting pipes from sheet {sheet.SheetNumber}: {ex.Message}");
            }

            return sheetPipeData;
        }

        private List<Element> GetPipesFromView(Document doc, View view)
        {
            var pipes = new List<Element>();

            try
            {
                // Use a more comprehensive approach to collect pipe elements
                var collector = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType();

                pipes.AddRange(collector.ToElements());

                // Also try to get pipes from model groups in the view
                var groupCollector = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Group));

                foreach (Element element in groupCollector)
                {
                    if (element is Group group)
                    {
                        foreach (ElementId memberId in group.GetMemberIds())
                        {
                            Element member = doc.GetElement(memberId);
                            if (member is Pipe)
                            {
                                pipes.Add(member);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting pipes from view {view.Name}: {ex.Message}");
            }

            return pipes.Distinct().ToList(); // Remove duplicates
        }

        private string GetSpecPositionParameter(Element pipe)
        {
            try
            {
                // Try to get the SPEC_POSITION parameter using different methods
                Parameter specParam = pipe.LookupParameter(SPEC_POSITION_PARAMETER);

                // If not found, try alternative parameter names
                if (specParam == null || !specParam.HasValue)
                {
                    string[] alternativeParams = { "Spec Position", "SpecPosition", "SPEC_POS", "Specification Position" };
                    foreach (string paramName in alternativeParams)
                    {
                        specParam = pipe.LookupParameter(paramName);
                        if (specParam != null && specParam.HasValue) break;
                    }
                }

                // If still not found, try to get parameter by built-in parameter if available
                if ((specParam == null || !specParam.HasValue) && pipe is Pipe)
                {
                    specParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                }

                if (specParam != null && specParam.HasValue)
                {
                    switch (specParam.StorageType)
                    {
                        case StorageType.String:
                            return specParam.AsString();
                        case StorageType.Integer:
                            return specParam.AsInteger().ToString();
                        case StorageType.Double:
                            return specParam.AsDouble().ToString();
                        case StorageType.ElementId:
                            ElementId id = specParam.AsElementId();
                            if (id != ElementId.InvalidElementId)
                            {
                                Element elem = pipe.Document.GetElement(id);
                                return elem?.Name ?? "";
                            }
                            break;
                    }
                }

                // Additional fallback: try to get parameter using GetParameters method
                var parameters = pipe.GetParameters(SPEC_POSITION_PARAMETER);
                if (parameters.Any())
                {
                    foreach (Parameter param in parameters)
                    {
                        if (param.HasValue)
                        {
                            if (param.StorageType == StorageType.String)
                                return param.AsString();
                            else if (param.StorageType == StorageType.Integer)
                                return param.AsInteger().ToString();
                            else if (param.StorageType == StorageType.Double)
                                return param.AsDouble().ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting SPEC_POSITION parameter: {ex.Message}");
            }

            return string.Empty;
        }
    }
}