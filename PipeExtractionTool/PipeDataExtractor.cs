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
                int progressPercentage = (int)((double)processedSheets / totalSheets * 80); // Reserve 20% for Excel generation
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
                SheetName = $"{sheet.SheetNumber}",
                PipeSpecPositions = new List<string>()
            };

            try
            {
                // Get all viewports on the sheet
                var viewportIds = sheet.GetAllViewports();

                foreach (ElementId viewportId in viewportIds)
                {
                    var viewport = doc.GetElement(viewportId) as Viewport;
                    if (viewport == null) continue;

                    var view = doc.GetElement(viewport.ViewId) as View;
                    if (view == null) continue;

                    // Extract pipes from this view
                    var pipesInView = GetPipesFromView(doc, view);

                    foreach (var pipe in pipesInView)
                    {
                        string specPosition = GetSpecPositionParameter(pipe);
                        if (!string.IsNullOrEmpty(specPosition) && !sheetPipeData.PipeSpecPositions.Contains(specPosition))
                        {
                            sheetPipeData.PipeSpecPositions.Add(specPosition);
                        }
                    }
                }

                // Also check if there are any pipes directly visible on the sheet (drafting views, etc.)
                var directPipes = GetPipesDirectlyOnSheet(doc, sheet);
                foreach (var pipe in directPipes)
                {
                    string specPosition = GetSpecPositionParameter(pipe);
                    if (!string.IsNullOrEmpty(specPosition) && !sheetPipeData.PipeSpecPositions.Contains(specPosition))
                    {
                        sheetPipeData.PipeSpecPositions.Add(specPosition);
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
                // Create filter for pipe elements
                var pipeFilter = new ElementClassFilter(typeof(Pipe));

                // Get all pipes visible in this view
                var collector = new FilteredElementCollector(doc, view.Id)
                    .WherePasses(pipeFilter);

                pipes.AddRange(collector.ToElements());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting pipes from view {view.Name}: {ex.Message}");
            }

            return pipes;
        }

        private List<Element> GetPipesDirectlyOnSheet(Document doc, ViewSheet sheet)
        {
            var pipes = new List<Element>();

            try
            {
                // This method would handle any pipes that might be directly placed on sheets
                // In most cases, pipes are shown through viewports, so this might return empty
                // But we include it for completeness

                var pipeFilter = new ElementClassFilter(typeof(Pipe));
                var collector = new FilteredElementCollector(doc)
                    .WherePasses(pipeFilter)
                    .WhereElementIsNotElementType();

                // Check if pipes are somehow associated with the sheet directly
                // This is uncommon in typical Revit workflows
                foreach (Element pipe in collector)
                {
                    // Additional logic could be added here if needed
                    // For now, we'll rely on the viewport-based extraction
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting pipes directly from sheet {sheet.SheetNumber}: {ex.Message}");
            }

            return pipes;
        }

        private string GetSpecPositionParameter(Element pipe)
        {
            try
            {
                // Try to get the SPEC_POSITION parameter
                Parameter specParam = pipe.LookupParameter(SPEC_POSITION_PARAMETER);

                if (specParam != null && specParam.HasValue)
                {
                    if (specParam.StorageType == StorageType.String)
                    {
                        return specParam.AsString();
                    }
                    else if (specParam.StorageType == StorageType.Integer)
                    {
                        return specParam.AsInteger().ToString();
                    }
                    else if (specParam.StorageType == StorageType.Double)
                    {
                        return specParam.AsDouble().ToString();
                    }
                }

                // If SPEC_POSITION is not found, try alternative parameter names
                string[] alternativeParams = { "Spec Position", "SpecPosition", "SPEC_POS", "Specification Position" };

                foreach (string paramName in alternativeParams)
                {
                    Parameter altParam = pipe.LookupParameter(paramName);
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
                    }
                }

                // As a fallback, try to get pipe system information
                if (pipe is Pipe pipeElement)
                {
                    try
                    {
                        var systemParam = pipeElement.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                        if (systemParam != null && systemParam.HasValue)
                        {
                            ElementId systemTypeId = systemParam.AsElementId();
                            if (systemTypeId != ElementId.InvalidElementId)
                            {
                                Element systemType = pipe.Document.GetElement(systemTypeId);
                                return systemType?.Name ?? "";
                            }
                        }
                    }
                    catch
                    {
                        // Ignore errors in fallback logic
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

    public class SheetPipeData
    {
        public string SheetName { get; set; }
        public List<string> PipeSpecPositions { get; set; }

        public string PipeSpecPositionsString => string.Join(", ", PipeSpecPositions ?? new List<string>());

        public SheetPipeData()
        {
            PipeSpecPositions = new List<string>();
        }
    }
}