using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace PipeExtractionTool
{
    public class PipeDataExtractor
    {
        // Exact parameter name required (per your requirement)
        private const string SPEC_POSITION_PARAMETER = "SPEC_POSITION";

        // Enable for diagnostics; set to false for production.
        private readonly bool DebugMode = true;

        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "PipeExtractor.log");

        // Exact allowed families and type names (case-sensitive). Adjust if you prefer "AND" instead of "OR".
        private static readonly HashSet<string> AllowedTagFamilyNames = new HashSet<string>
        {
            "DX_TAG_PiP_BW-tag-DoublePipes_DE",
            "DX_TAG_PiP_BW-tag_DE"
        };

        private static readonly HashSet<string> AllowedTagTypeNames = new HashSet<string>
        {
            "BW_Tag(opaque)",
            "BWBW_Tag_Sloped_ohneRA",
            "BW_Tag_vertical",
            "BW_Tag_Sloped_ohneRA(opaque)",
            "BW_Tag"
        };

        public PipeDataExtractor()
        {
            if (DebugMode)
            {
                try
                {
                    File.WriteAllText(LogPath, $"--- PipeExtractor log started {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---{Environment.NewLine}");
                }
                catch { }
            }
        }

        private void Log(string line)
        {
            if (!DebugMode) return;
            try { File.AppendAllText(LogPath, $"{DateTime.Now:HH:mm:ss} {line}{Environment.NewLine}"); } catch { }
        }

        /// <summary>
        /// Primary entry. Now accepts SimpleProgressReporter instead of BackgroundWorker
        /// </summary>
        public List<SheetPipeData> ExtractPipeData(Document doc, List<DrawingSheetInfo> selectedSheets, SimpleProgressReporter progressReporter = null)
        {
            var result = new List<SheetPipeData>();
            int totalSheets = selectedSheets?.Count ?? 0;
            int processedSheets = 0;

            Log($"Starting SPEC_POSITION extraction. SelectedSheets={totalSheets}");

            // Set up progress reporting
            progressReporter?.SetTotalSheets(totalSheets);

            foreach (var sheetInfo in selectedSheets)
            {
                if (progressReporter?.CancellationPending == true)
                {
                    Log("Cancellation requested - stopping sheet processing.");
                    break;
                }

                Log($"Processing sheet: {sheetInfo?.DisplayName ?? "(null)"}");

                var sheetPipeData = new SheetPipeData
                {
                    SheetName = $"{sheetInfo?.ViewSheet?.SheetNumber ?? ""} - {sheetInfo?.ViewSheet?.Name ?? sheetInfo?.DisplayName ?? ""}",
                    PipeSpecPositions = new List<string>()
                };

                try
                {
                    if (sheetInfo?.ViewSheet == null)
                    {
                        Log("  SheetInfo.ViewSheet is null - skipping.");
                        result.Add(sheetPipeData);
                        processedSheets++;
                        progressReporter?.IncrementSheet();
                        continue;
                    }

                    var viewportIds = sheetInfo.ViewSheet.GetAllViewports();
                    Log($"  Sheet has {viewportIds?.Count ?? 0} viewport(s).");

                    foreach (ElementId vpId in viewportIds)
                    {
                        if (progressReporter?.CancellationPending == true)
                        {
                            Log("Cancellation requested during view loop.");
                            break;
                        }

                        var vpElem = doc.GetElement(vpId) as Viewport;
                        if (vpElem == null)
                        {
                            Log($"   Viewport ElementId {vpId} not a Viewport - skipping.");
                            continue;
                        }

                        var view = doc.GetElement(vpElem.ViewId) as View;
                        if (view == null)
                        {
                            Log($"   View for viewport {vpId} not found - skipping.");
                            continue;
                        }

                        Log($"   View: {view.Name} (Id={view.Id.IntegerValue}) Type={view.ViewType}");

                        // STRICT: only FloorPlan and Section
                        if (view.ViewType != ViewType.FloorPlan && view.ViewType != ViewType.Section)
                        {
                            Log("    SKIP: view is not FloorPlan or Section.");
                            continue;
                        }

                        var foundSpecs = ExtractSpecPositionsFromTags(doc, view, progressReporter);
                        foreach (var s in foundSpecs)
                        {
                            if (!string.IsNullOrEmpty(s) && !sheetPipeData.PipeSpecPositions.Contains(s))
                            {
                                sheetPipeData.PipeSpecPositions.Add(s);
                                Log($"     Added SPEC_POSITION '{s}' for sheet {sheetPipeData.SheetName}");
                            }
                        }
                    }
                }
                catch (Exception exSheet)
                {
                    Log($"  Error processing sheet '{sheetInfo?.DisplayName}': {exSheet}");
                }

                sheetPipeData.PipeSpecPositions.Sort(StringComparer.OrdinalIgnoreCase);
                result.Add(sheetPipeData);

                processedSheets++;
                progressReporter?.IncrementSheet();
            }

            Log("SPEC_POSITION extraction finished.");
            if (DebugMode) Log($"Log file: {LogPath}");
            return result;
        }

        /// <summary>
        /// Overload to maintain compatibility with BackgroundWorker calls
        /// </summary>
        public List<SheetPipeData> ExtractPipeData(Document doc, List<DrawingSheetInfo> selectedSheets, BackgroundWorker worker = null)
        {
            // Create a wrapper that converts BackgroundWorker to SimpleProgressReporter
            SimpleProgressReporter progressReporter = null;
            if (worker != null)
            {
                progressReporter = new BackgroundWorkerProgressAdapter(worker);
            }

            return ExtractPipeData(doc, selectedSheets, progressReporter);
        }

        /// <summary>
        /// Extract SPEC_POSITION tokens from tags in the given view.
        /// </summary>
        private List<string> ExtractSpecPositionsFromTags(Document doc, View view, SimpleProgressReporter progressReporter = null)
        {
            var specPositions = new List<string>();

            try
            {
                var tags = GetAllTagsInView(doc, view);
                Log($"    Found {tags.Count} tag elements in view '{view.Name}'");

                foreach (var tag in tags)
                {
                    if (progressReporter?.CancellationPending == true) break;

                    try
                    {
                        // read type/family using the tag's document (safer for linked/host)
                        string typeName = string.Empty;
                        string familyName = string.Empty;
                        try
                        {
                            var typeElem = tag.Document.GetElement(tag.GetTypeId());
                            if (typeElem != null)
                            {
                                typeName = typeElem.Name ?? string.Empty;
                                if (typeElem is FamilySymbol fs && fs.Family != null) familyName = fs.Family.Name ?? string.Empty;
                            }
                        }
                        catch (Exception exType) { Log($"      Error reading type/family for tag {tag.Id.IntegerValue}: {exType.Message}"); }

                        Log($"      Tag {tag.Id.IntegerValue}: Family='{familyName}' Type='{typeName}'");

                        // STRICT exact spelling: accept if family OR type matches exactly (case-sensitive)
                        bool allowed = (AllowedTagFamilyNames.Contains(familyName) || AllowedTagTypeNames.Contains(typeName));
                        if (!allowed)
                        {
                            Log($"        SKIP tag {tag.Id.IntegerValue} => family/type not allowed.");
                            continue;
                        }

                        // 1) Try SPEC_POSITION parameter on tag instance (exact name)
                        string rawValue = string.Empty;
                        try
                        {
                            var p = tag.LookupParameter(SPEC_POSITION_PARAMETER);
                            if (p != null && p.HasValue && p.StorageType == StorageType.String)
                            {
                                rawValue = p.AsString() ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(rawValue))
                                    Log($"        SPEC_POSITION param on tag = '{rawValue}'");
                            }
                            else
                            {
                                Log($"        SPEC_POSITION param not found or empty on tag instance.");
                            }
                        }
                        catch (Exception exP) { Log($"        Error reading SPEC_POSITION param on tag {tag.Id.IntegerValue}: {exP.Message}"); }

                        // 2) If empty, check tagged local elements (host model)
                        if (string.IsNullOrWhiteSpace(rawValue) && tag is IndependentTag indTag)
                        {
                            try
                            {
                                var localElems = indTag.GetTaggedLocalElements();
                                if (localElems != null && localElems.Count > 0)
                                {
                                    foreach (var le in localElems)
                                    {
                                        if (le == null) continue;
                                        var val = GetSpecPositionParameterStrict(le, out string ff, out string pn);
                                        Log($"          Tagged local element {le.Id.IntegerValue} param read: foundFrom={ff} param={pn} val='{val}'");
                                        if (!string.IsNullOrWhiteSpace(val))
                                        {
                                            rawValue = val;
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    Log("          No tagged local elements returned.");
                                }
                            }
                            catch (Exception exLocal) { Log($"          Error reading tagged local elements: {exLocal.Message}"); }
                        }

                        // 3) If still empty, check references (covers linked elements and host element references)
                        if (string.IsNullOrWhiteSpace(rawValue) && tag is IndependentTag indTagRefs)
                        {
                            try
                            {
                                var refs = indTagRefs.GetTaggedReferences();
                                if (refs != null && refs.Count > 0)
                                {
                                    foreach (var tref in refs)
                                    {
                                        if (tref == null) continue;
                                        try
                                        {
                                            Log($"          Reference: ElementId={tref.ElementId} LinkedElementId={tref.LinkedElementId}");

                                            if (tref.LinkedElementId != ElementId.InvalidElementId)
                                            {
                                                // reference into linked document
                                                var linkInstance = tag.Document.GetElement(tref.ElementId) as RevitLinkInstance;
                                                if (linkInstance != null)
                                                {
                                                    var linkDoc = linkInstance.GetLinkDocument();
                                                    Log($"            Link instance found (name='{linkInstance.Name}'), linkDoc title='{linkDoc?.Title ?? "null"}'");
                                                    if (linkDoc != null)
                                                    {
                                                        var linkedElem = linkDoc.GetElement(tref.LinkedElementId);
                                                        if (linkedElem != null)
                                                        {
                                                            var val = GetSpecPositionParameterStrict(linkedElem, out string ff, out string pn);
                                                            Log($"              Linked element param read: foundFrom={ff} param={pn} val='{val}'");
                                                            if (!string.IsNullOrWhiteSpace(val)) { rawValue = val; break; }
                                                        }
                                                        else
                                                        {
                                                            Log($"              Linked element {tref.LinkedElementId} not found in linkDoc.");
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Log("            LinkInstance element not found or not a RevitLinkInstance.");
                                                }
                                            }
                                            else
                                            {
                                                // host element reference
                                                var hostElem = tag.Document.GetElement(tref.ElementId);
                                                if (hostElem != null)
                                                {
                                                    var val = GetSpecPositionParameterStrict(hostElem, out string ff, out string pn);
                                                    Log($"              Host referenced element param read: foundFrom={ff} param={pn} val='{val}'");
                                                    if (!string.IsNullOrWhiteSpace(val)) { rawValue = val; break; }
                                                }
                                                else Log("              Host referenced element not found.");
                                            }
                                        }
                                        catch (Exception exRef) { Log($"            Error processing reference: {exRef.Message}"); }
                                    }
                                }
                                else
                                {
                                    Log("          No references returned from GetTaggedReferences().");
                                }
                            }
                            catch (Exception exRefs) { Log($"          Error getting tagged references: {exRefs.Message}"); }
                        }

                        // 4) If still empty -> visible text fallback (only numeric-dash tokens)
                        if (string.IsNullOrWhiteSpace(rawValue))
                        {
                            var visibleText = GetTagTextContent(tag);
                            Log($"        Visible text fallback for tag {tag.Id.IntegerValue}: '{visibleText}'");
                            var tokens = ExtractSpecPositionsFromText(visibleText);
                            foreach (var t in tokens)
                            {
                                if (!specPositions.Contains(t)) specPositions.Add(t);
                                Log($"          Added token from visible text: {t}");
                            }
                            continue;
                        }

                        // 5) We have rawValue(s) from parameter(s) - split and normalize tokens
                        var rawSplits = rawValue.Split(new[] { ',', ';', '/', '|', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                                .Select(x => x.Trim())
                                                .Where(x => !string.IsNullOrWhiteSpace(x))
                                                .ToList();

                        Log($"        Raw value splits: {string.Join(" | ", rawSplits)}");

                        foreach (var rawTok in rawSplits)
                        {
                            // Remove common prefixes like "Nr." or "No."
                            var cleaned = Regex.Replace(rawTok, @"^(Nr\.?|No\.?)\s*", "", RegexOptions.IgnoreCase);

                            // Normalize spacing around hyphen to canonical form
                            var normalized = Regex.Replace(cleaned, @"\s*-\s*", "-");

                            // Extract numeric-dash tokens (1-4 digits each side)
                            var matches = Regex.Matches(normalized, @"\b\d{1,4}-\d{1,4}\b");
                            foreach (Match m in matches)
                            {
                                var token = m.Value.Trim();
                                if (!specPositions.Contains(token))
                                {
                                    specPositions.Add(token);
                                    Log($"          Added token '{token}' from parameter value (orig='{rawTok}')");
                                }
                            }
                        }
                    }
                    catch (Exception exTag)
                    {
                        Log($"      Error processing tag {tag?.Id?.IntegerValue.ToString() ?? "(unknown)"}: {exTag.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"    Error extracting from tags in view '{view?.Name}': {ex.Message}");
            }

            return specPositions.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>
        /// Strict read of SPEC_POSITION from element instance or its type (no heuristics).
        /// </summary>
        private string GetSpecPositionParameterStrict(Element elem, out string foundFrom, out string foundParamName)
        {
            foundFrom = "None";
            foundParamName = string.Empty;

            try
            {
                if (elem == null) return string.Empty;

                // instance parameter (exact)
                try
                {
                    var p = elem.LookupParameter(SPEC_POSITION_PARAMETER);
                    if (p != null && p.HasValue && p.StorageType == StorageType.String)
                    {
                        var v = p.AsString();
                        if (!string.IsNullOrWhiteSpace(v))
                        {
                            foundFrom = "Instance";
                            foundParamName = p.Definition?.Name ?? SPEC_POSITION_PARAMETER;
                            return v.Trim();
                        }
                    }
                }
                catch { /* ignore */ }

                // type parameter (exact)
                try
                {
                    var typeId = elem.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        var typeElem = elem.Document.GetElement(typeId);
                        if (typeElem != null)
                        {
                            var tp = typeElem.LookupParameter(SPEC_POSITION_PARAMETER);
                            if (tp != null && tp.HasValue && tp.StorageType == StorageType.String)
                            {
                                var tv = tp.AsString();
                                if (!string.IsNullOrWhiteSpace(tv))
                                {
                                    foundFrom = "Type";
                                    foundParamName = tp.Definition?.Name ?? SPEC_POSITION_PARAMETER;
                                    return tv.Trim();
                                }
                            }
                        }
                    }
                }
                catch { /* ignore */ }
            }
            catch (Exception ex)
            {
                Log($"        Error in GetSpecPositionParameterStrict for elem {elem?.Id.IntegerValue}: {ex.Message}");
            }

            return string.Empty;
        }

        /// <summary>
        /// Extract numeric-dash tokens from arbitrary text. Handles "Nr. 474-90" and "474 - 90".
        /// </summary>
        private List<string> ExtractSpecPositionsFromText(string text)
        {
            var results = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return results;

            try
            {
                // Remove common prefix if present at start (helps for "Nr. 474-90")
                var cleanedStart = Regex.Replace(text, @"^(Nr\.?|No\.?)\s*", "", RegexOptions.IgnoreCase);

                // Allow optional spaces around hyphen and normalize to single '-'
                var pattern = @"\b\d{1,4}\s*-\s*\d{1,4}\b";
                var matches = Regex.Matches(cleanedStart, pattern);
                foreach (Match m in matches)
                {
                    var normalized = Regex.Replace(m.Value, @"\s*-\s*", "-").Trim();
                    if (!results.Contains(normalized)) results.Add(normalized);
                }
            }
            catch (Exception ex)
            {
                Log($"        Error extracting numeric-dash tokens from text '{text}': {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// Get visible text for a tag or textnote (fallback).
        /// </summary>
        private string GetTagTextContent(Element tag)
        {
            try
            {
                if (tag == null) return string.Empty;

                // TextNote
                if (tag is TextNote tn) return tn.Text ?? string.Empty;

                // IndependentTag.TagText (may throw for linked tags)
                if (tag is IndependentTag indTag)
                {
                    try
                    {
                        var tt = indTag.TagText;
                        if (!string.IsNullOrWhiteSpace(tt)) return tt;
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // TagText may throw for tags in linked contexts; we fall back to parameters
                        Log($"        TagText threw InvalidOperationException for tag {tag.Id.IntegerValue} (likely linked).");
                    }
                    catch (Exception ex)
                    {
                        Log($"        Unexpected error reading TagText for tag {tag.Id.IntegerValue}: {ex.Message}");
                    }
                }

                // Check common string parameter names
                var textParams = new[] { "Text", "TagText", "Value", "Label" };
                foreach (var pn in textParams)
                {
                    try
                    {
                        var p = tag.LookupParameter(pn);
                        if (p != null && p.HasValue && p.StorageType == StorageType.String)
                        {
                            var s = p.AsString();
                            if (!string.IsNullOrWhiteSpace(s)) return s;
                        }
                    }
                    catch { /* ignore param read errors */ }
                }

                // Last resort: enumerate string parameters and pick the first non-empty
                foreach (Parameter p in tag.Parameters)
                {
                    try
                    {
                        if (p.StorageType == StorageType.String && p.HasValue)
                        {
                            var s = p.AsString();
                            if (!string.IsNullOrWhiteSpace(s)) return s;
                        }
                    }
                    catch { /* ignore */ }
                }
            }
            catch (Exception ex)
            {
                Log($"        Error getting tag text: {ex.Message}");
            }

            return string.Empty;
        }

        /// <summary>
        /// Collect tags in view: IndependentTag and TextNote (deduplicated).
        /// </summary>
        private List<Element> GetAllTagsInView(Document doc, View view)
        {
            var tags = new List<Element>();

            try
            {
                try
                {
                    var indTags = new FilteredElementCollector(doc, view.Id)
                        .OfClass(typeof(IndependentTag))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    if (indTags != null && indTags.Count > 0) tags.AddRange(indTags);
                }
                catch (Exception ex) { Log($"      Error collecting IndependentTag items: {ex.Message}"); }

                try
                {
                    var textNotes = new FilteredElementCollector(doc, view.Id)
                        .OfClass(typeof(TextNote))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    if (textNotes != null && textNotes.Count > 0) tags.AddRange(textNotes);
                }
                catch (Exception ex) { Log($"      Error collecting TextNote items: {ex.Message}"); }
            }
            catch (Exception ex)
            {
                Log($"    Error getting tags from view: {ex.Message}");
            }

            // dedupe by Element.Id
            return tags.GroupBy(e => e.Id.IntegerValue).Select(g => g.First()).ToList();
        }
    }

    /// <summary>
    /// Adapter to convert BackgroundWorker progress reporting to SimpleProgressReporter
    /// </summary>
    public class BackgroundWorkerProgressAdapter : SimpleProgressReporter
    {
        private readonly BackgroundWorker _backgroundWorker;

        public BackgroundWorkerProgressAdapter(BackgroundWorker backgroundWorker) : base(null)
        {
            _backgroundWorker = backgroundWorker;
        }

        public new bool CancellationPending => _backgroundWorker?.CancellationPending ?? false;

        public new void ReportProgress(int percentage, string message)
        {
            _backgroundWorker?.ReportProgress(percentage, message);
        }
    }
}