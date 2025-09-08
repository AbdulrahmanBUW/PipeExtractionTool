using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace PipeExtractionTool
{
    public class PipeDataExtractor
    {
        private const string SPEC_POSITION_PARAMETER = "SPEC_POSITION";
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "PipeExtractor.log");

        // Toggle for verbose debug (parameter dumps etc.). Set to true only when diagnosing.
        private readonly bool DebugMode = false;

        public PipeDataExtractor()
        {
            if (DebugMode)
            {
                try { File.WriteAllText(LogPath, $"--- PipeExtractor log started {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---{Environment.NewLine}"); } catch { }
            }
        }

        private void Log(string line)
        {
            if (!DebugMode) return;
            try { Debug.WriteLine(line); } catch { }
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }

        /// <summary>
        /// Primary entry. Must be called on Revit UI thread for best reliability.
        /// </summary>
        public List<SheetPipeData> ExtractPipeData(Document doc, List<DrawingSheetInfo> selectedSheets, BackgroundWorker worker = null)
        {
            var result = new List<SheetPipeData>();
            int totalSheets = selectedSheets?.Count ?? 0;
            int processedSheets = 0;

            // 1) Pre-collect host pipes and cache their model-space bounding boxes once
            Log("Pre-collecting host pipes and caching bounding boxes...");
            var hostPipeCache = new Dictionary<int, BoundingBoxXYZ>(); // key = ElementId.IntegerValue
            var allHostPipes = new List<Element>();
            try
            {
                allHostPipes = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                foreach (var p in allHostPipes)
                {
                    try
                    {
                        BoundingBoxXYZ bb = null;
                        try { bb = p.get_BoundingBox(null); } catch { bb = null; }
                        if (bb == null)
                        {
                            // fallback to rep point -> tiny bbox
                            var rep = GetRepresentativePoint(p);
                            if (rep != null)
                            {
                                double eps = 0.001; // 1 mm
                                bb = new BoundingBoxXYZ { Min = new XYZ(rep.X - eps, rep.Y - eps, rep.Z - eps), Max = new XYZ(rep.X + eps, rep.Y + eps, rep.Z + eps) };
                            }
                        }
                        if (bb != null) hostPipeCache[p.Id.IntegerValue] = bb;
                    }
                    catch { /* ignore per-element errors */ }
                }
                Log($"  Cached {hostPipeCache.Count} host pipe bounding boxes.");
            }
            catch (Exception ex)
            {
                Log($"Error pre-collecting host pipes: {ex}");
            }

            // 2) Pre-collect link instances and build a cache of linked pipes transformed into host-space (host bounding boxes)
            Log("Pre-collecting link instances and building per-link cache...");
            var linkInstances = new List<RevitLinkInstance>();
            var linkPipeCache = new Dictionary<int, List<LinkedPipeInfo>>(); // key = linkInstance.Id.IntegerValue
            try
            {
                linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                Log($"  Found {linkInstances.Count} RevitLinkInstance(s). Building caches (this may take a moment).");

                foreach (var li in linkInstances)
                {
                    // Honor cancellation frequently
                    if (worker?.CancellationPending == true) { Log("Cancellation requested while building link cache."); break; }

                    try
                    {
                        Document linkDoc = null;
                        try { linkDoc = li.GetLinkDocument(); } catch { linkDoc = null; }
                        if (linkDoc == null)
                        {
                            Log($"    LinkInstance Id={li.Id} not loaded or returns null doc. Skipping cache for this link.");
                            continue;
                        }

                        Transform transformToHost = null;
                        try { transformToHost = li.GetTotalTransform(); } catch { transformToHost = null; }

                        var linkedPipes = new FilteredElementCollector(linkDoc)
                            .OfClass(typeof(Pipe))
                            .WhereElementIsNotElementType()
                            .ToElements()
                            .ToList();

                        var list = new List<LinkedPipeInfo>(linkedPipes.Count);
                        foreach (var lp in linkedPipes)
                        {
                            try
                            {
                                // read link element bbox (model space of link)
                                BoundingBoxXYZ linkBb = null;
                                try { linkBb = lp.get_BoundingBox(null); } catch { linkBb = null; }

                                BoundingBoxXYZ hostBb = null;
                                XYZ repHostPoint = null;

                                if (linkBb != null && transformToHost != null)
                                {
                                    // transform bbox corners to host coords and compute axis-aligned host bbox
                                    var corners = GetBoundingBoxCorners(linkBb);
                                    var tMin = new XYZ(double.MaxValue, double.MaxValue, double.MaxValue);
                                    var tMax = new XYZ(double.MinValue, double.MinValue, double.MinValue);
                                    foreach (var c in corners)
                                    {
                                        var tc = transformToHost.OfPoint(c);
                                        tMin = new XYZ(Math.Min(tMin.X, tc.X), Math.Min(tMin.Y, tc.Y), Math.Min(tMin.Z, tc.Z));
                                        tMax = new XYZ(Math.Max(tMax.X, tc.X), Math.Max(tMax.Y, tc.Y), Math.Max(tMax.Z, tc.Z));
                                    }
                                    hostBb = new BoundingBoxXYZ { Min = tMin, Max = tMax };
                                }
                                else
                                {
                                    // fallback to representative point in link coords and transform
                                    var rep = GetRepresentativePoint(lp);
                                    if (rep != null && transformToHost != null)
                                    {
                                        repHostPoint = transformToHost.OfPoint(rep);
                                        double eps = 0.001;
                                        hostBb = new BoundingBoxXYZ { Min = new XYZ(repHostPoint.X - eps, repHostPoint.Y - eps, repHostPoint.Z - eps), Max = new XYZ(repHostPoint.X + eps, repHostPoint.Y + eps, repHostPoint.Z + eps) };
                                    }
                                }

                                var info = new LinkedPipeInfo
                                {
                                    Element = lp,
                                    HostBoundingBox = hostBb,
                                    LinkBoundingBox = linkBb
                                };

                                list.Add(info);
                            }
                            catch { /* skip individual linked pipe errors */ }
                        }

                        linkPipeCache[li.Id.IntegerValue] = list;
                        Log($"    Cached {list.Count} pipes for link '{li.Name}' (link Id={li.Id}).");
                    }
                    catch (Exception exLi) { Log($"    Error building cache for link Id={li.Id}: {exLi}"); }
                }
            }
            catch (Exception ex)
            {
                Log($"Error pre-collecting link instances: {ex}");
            }

            // Now process sheets/views using caches. We still attempt precise view collectors first where possible.
            foreach (var sheetInfo in selectedSheets)
            {
                if (worker?.CancellationPending == true) { Log("Cancellation requested - stopping sheet processing."); break; }

                int progressPercentage = totalSheets == 0 ? 100 : (int)((double)processedSheets / totalSheets * 80);
                worker?.ReportProgress(progressPercentage, $"Processing sheet: {sheetInfo.DisplayName}");

                Log($"Processing sheet: {sheetInfo.DisplayName}");
                var sheetPipeData = new SheetPipeData { SheetName = $"{sheetInfo.ViewSheet.SheetNumber} - {sheetInfo.ViewSheet.Name}", PipeSpecPositions = new List<string>() };

                try
                {
                    var viewportIds = sheetInfo.ViewSheet.GetAllViewports();
                    foreach (ElementId vpId in viewportIds)
                    {
                        if (worker?.CancellationPending == true) { Log("Cancellation requested during view loop."); break; }

                        if (!(doc.GetElement(vpId) is Viewport viewport)) continue;
                        if (!(doc.GetElement(viewport.ViewId) is View view)) continue;
                        Log($"  View: {view.Name} (Id={view.Id})");

                        // try view-scoped collector first (precise)
                        bool preciseUsed = false;
                        List<Element> preciseHostPipes = null;
                        List<RevitLinkInstance> preciseLinkInstances = null;
                        try
                        {
                            preciseHostPipes = new FilteredElementCollector(doc, view.Id)
                                .OfClass(typeof(Pipe))
                                .WhereElementIsNotElementType()
                                .ToElements()
                                .ToList();

                            preciseLinkInstances = new FilteredElementCollector(doc, view.Id)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>()
                                .ToList();

                            preciseUsed = true;
                            Log($"    [Precise] host pipes found: {preciseHostPipes.Count}, link instances found: {preciseLinkInstances.Count}");
                        }
                        catch (Exception exPrecise)
                        {
                            preciseUsed = false;
                            Log($"    [Precise] collectors failed for view '{view.Name}': {exPrecise.Message}. Using cached bbox-based filtering.");
                        }

                        var candidates = new List<Element>();

                        if (preciseUsed)
                        {
                            // Filter preciseHostPipes by bbox intersection with view crop (use get_BoundingBox(view) first)
                            foreach (var hp in preciseHostPipes)
                            {
                                if (worker?.CancellationPending == true) break;
                                try
                                {
                                    if (ElementIntersectsViewCrop(hp, view, hostDocElementIsHost: true, linkTransform: null, linkInstance: null))
                                        candidates.Add(hp);
                                }
                                catch { candidates.Add(hp); } // conservative
                            }

                            // For each link instance returned by precise collector, use our precomputed linkPipeCache (if available)
                            foreach (var li in preciseLinkInstances)
                            {
                                if (worker?.CancellationPending == true) break;
                                try
                                {
                                    if (!linkPipeCache.TryGetValue(li.Id.IntegerValue, out List<LinkedPipeInfo> cachedList))
                                    {
                                        // link wasn't cached earlier (not loaded at cache time) - do a short dynamic collect (best-effort)
                                        Document linkDoc = null;
                                        try { linkDoc = li.GetLinkDocument(); } catch { linkDoc = null; }
                                        if (linkDoc == null) continue;

                                        var linkedPipes = new FilteredElementCollector(linkDoc)
                                            .OfClass(typeof(Pipe))
                                            .WhereElementIsNotElementType()
                                            .ToElements()
                                            .ToList();

                                        Transform tr = null;
                                        try { tr = li.GetTotalTransform(); } catch { tr = null; }

                                        foreach (var lp in linkedPipes)
                                        {
                                            if (worker?.CancellationPending == true) break;
                                            try
                                            {
                                                // quick test via ElementIntersectsViewCrop using tr
                                                if (ElementIntersectsViewCrop(lp, view, hostDocElementIsHost: false, linkTransform: tr, linkInstance: li))
                                                    candidates.Add(lp);
                                            }
                                            catch { candidates.Add(lp); }
                                        }
                                    }
                                    else
                                    {
                                        // use cached host bboxes
                                        foreach (var info in cachedList)
                                        {
                                            if (worker?.CancellationPending == true) break;
                                            try
                                            {
                                                if (info.HostBoundingBox != null)
                                                {
                                                    if (view.CropBoxActive && view.CropBox != null)
                                                    {
                                                        if (BoundingBoxesIntersect(info.HostBoundingBox, view.CropBox))
                                                            candidates.Add(info.Element);
                                                    }
                                                    else
                                                    {
                                                        candidates.Add(info.Element);
                                                    }
                                                }
                                                else
                                                {
                                                    // No host bbox cached (rare) -> fall back to ElementIntersectsViewCrop to compute small rep and test
                                                    if (ElementIntersectsViewCrop(info.Element, view, hostDocElementIsHost: false, linkTransform: li.GetTotalTransform(), linkInstance: li))
                                                        candidates.Add(info.Element);
                                                }
                                            }
                                            catch { candidates.Add(info.Element); }
                                        }
                                    }
                                }
                                catch (Exception exLi) { Log($"    Error handling precise link instance {li.Id}: {exLi.Message}"); }
                            }
                        }
                        else
                        {
                            // Use cached hostPipeCache and linkPipeCache to determine candidates (strict bbox check)
                            // Host cached bboxes:
                            foreach (var hostPipe in allHostPipes)
                            {
                                if (worker?.CancellationPending == true) break;
                                try
                                {
                                    if (!hostPipeCache.TryGetValue(hostPipe.Id.IntegerValue, out BoundingBoxXYZ hbb))
                                    {
                                        // no cached bbox -> try a direct bbox read
                                        try { hbb = hostPipe.get_BoundingBox(null); } catch { hbb = null; }
                                    }
                                    if (hbb != null && view.CropBoxActive && view.CropBox != null)
                                    {
                                        if (BoundingBoxesIntersect(hbb, view.CropBox))
                                            candidates.Add(hostPipe);
                                    }
                                    else if (hbb != null && (view.CropBox == null || !view.CropBoxActive))
                                    {
                                        candidates.Add(hostPipe);
                                    }
                                    else
                                    {
                                        // fallback rep point
                                        var rep = GetRepresentativePoint(hostPipe);
                                        if (rep != null && IsPointInViewCrop(rep, view))
                                            candidates.Add(hostPipe);
                                    }
                                }
                                catch { candidates.Add(hostPipe); }
                            }

                            // Linked caches:
                            foreach (var kv in linkPipeCache)
                            {
                                if (worker?.CancellationPending == true) break;
                                var cachedList = kv.Value;
                                foreach (var info in cachedList)
                                {
                                    if (worker?.CancellationPending == true) break;
                                    try
                                    {
                                        if (info.HostBoundingBox != null && view.CropBoxActive && view.CropBox != null)
                                        {
                                            if (BoundingBoxesIntersect(info.HostBoundingBox, view.CropBox))
                                                candidates.Add(info.Element);
                                        }
                                        else if (info.HostBoundingBox != null && (view.CropBox == null || !view.CropBoxActive))
                                        {
                                            candidates.Add(info.Element);
                                        }
                                        else
                                        {
                                            // fallback: rep point transformed was attempted when caching; if we don't have it, include conservatively
                                            candidates.Add(info.Element);
                                        }
                                    }
                                    catch { candidates.Add(info.Element); }
                                }
                            }
                        }

                        // Deduplicate candidates
                        var unique = new Dictionary<string, Element>();
                        foreach (var c in candidates)
                        {
                            if (c == null) continue;
                            string marker = c.Document?.PathName ?? c.Document?.Title ?? "host";
                            string key = $"{marker}::{c.Id.IntegerValue}";
                            if (!unique.ContainsKey(key)) unique[key] = c;
                        }

                        Log($"    Candidates before parameter extraction: {unique.Count}");

                        // Extract SPEC_POSITION from candidates
                        foreach (var kv in unique)
                        {
                            if (worker?.CancellationPending == true) break;

                            var elem = kv.Value;
                            string foundFrom, paramName;
                            string spec = GetSpecPositionParameter(elem, out foundFrom, out paramName);

                            if (!string.IsNullOrEmpty(spec))
                            {
                                if (!sheetPipeData.PipeSpecPositions.Contains(spec))
                                {
                                    sheetPipeData.PipeSpecPositions.Add(spec);
                                }
                                Log($"      FOUND: {spec} (from {foundFrom} param {paramName}) on elem {elem.Id.IntegerValue}");
                            }
                            else
                            {
                                if (DebugMode)
                                {
                                    Log($"      NOT FOUND on elem {elem.Id.IntegerValue} - dumping params...");
                                    DumpElementParameters(elem);
                                }
                            }
                        }
                    }
                }
                catch (Exception exSheet) { Log($"Error processing sheet '{sheetInfo.DisplayName}': {exSheet}"); }

                sheetPipeData.PipeSpecPositions.Sort(StringComparer.OrdinalIgnoreCase);
                result.Add(sheetPipeData);

                processedSheets++;
                // report progress occasionally
                worker?.ReportProgress((int)((double)processedSheets / totalSheets * 100), $"Sheets processed: {processedSheets}/{totalSheets}");
            }

            Log("Extraction finished.");
            if (DebugMode) Log($"Log file: {LogPath}");
            return result;
        }

        // --- Helpers & cache types ---

        private class LinkedPipeInfo
        {
            public Element Element { get; set; }
            public BoundingBoxXYZ LinkBoundingBox { get; set; }   // in link model space
            public BoundingBoxXYZ HostBoundingBox { get; set; }   // transformed into host model space if available
        }

        private BoundingBoxXYZ[] GetBoundingBoxCornersArray(BoundingBoxXYZ bb)
        {
            if (bb == null) return new BoundingBoxXYZ[0];
            var corners = GetBoundingBoxCorners(bb);
            return corners.ToArray().Select(c => new BoundingBoxXYZ()).ToArray(); // not used; we use GetBoundingBoxCorners directly
        }

        private List<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ bb)
        {
            var corners = new List<XYZ>();
            if (bb == null) return corners;
            var min = bb.Min;
            var max = bb.Max;
            corners.Add(new XYZ(min.X, min.Y, min.Z));
            corners.Add(new XYZ(min.X, min.Y, max.Z));
            corners.Add(new XYZ(min.X, max.Y, min.Z));
            corners.Add(new XYZ(min.X, max.Y, max.Z));
            corners.Add(new XYZ(max.X, min.Y, min.Z));
            corners.Add(new XYZ(max.X, min.Y, max.Z));
            corners.Add(new XYZ(max.X, max.Y, min.Z));
            corners.Add(new XYZ(max.X, max.Y, max.Z));
            return corners;
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            double aMinX = Math.Min(a.Min.X, a.Max.X), aMaxX = Math.Max(a.Min.X, a.Max.X);
            double aMinY = Math.Min(a.Min.Y, a.Max.Y), aMaxY = Math.Max(a.Min.Y, a.Max.Y);
            double aMinZ = Math.Min(a.Min.Z, a.Max.Z), aMaxZ = Math.Max(a.Min.Z, a.Max.Z);

            double bMinX = Math.Min(b.Min.X, b.Max.X), bMaxX = Math.Max(b.Min.X, b.Max.X);
            double bMinY = Math.Min(b.Min.Y, b.Max.Y), bMaxY = Math.Max(b.Min.Y, b.Max.Y);
            double bMinZ = Math.Min(b.Min.Z, b.Max.Z), bMaxZ = Math.Max(b.Min.Z, b.Max.Z);

            bool overlap = !(aMaxX < bMinX || aMinX > bMaxX ||
                             aMaxY < bMinY || aMinY > bMaxY ||
                             aMaxZ < bMinZ || aMinZ > bMaxZ);
            return overlap;
        }

        private XYZ GetRepresentativePoint(Element e)
        {
            try
            {
                var loc = e.Location;
                if (loc is LocationCurve lc && lc.Curve != null)
                    return (lc.Curve.GetEndPoint(0) + lc.Curve.GetEndPoint(1)) * 0.5;

                BoundingBoxXYZ bb = null;
                try { bb = e.get_BoundingBox(null); } catch { bb = null; }
                if (bb != null)
                    return (bb.Min + bb.Max) * 0.5;
            }
            catch { }
            return null;
        }

        private bool ElementIntersectsViewCrop(Element elem, View view, bool hostDocElementIsHost, Transform linkTransform, RevitLinkInstance linkInstance)
        {
            // Strict bbox-based test where possible, else rep point
            try
            {
                if (hostDocElementIsHost)
                {
                    // Try view-scoped bbox first
                    try
                    {
                        var bbInView = elem.get_BoundingBox(view);
                        if (bbInView != null)
                        {
                            if (view.CropBoxActive && view.CropBox != null)
                                return BoundingBoxesIntersect(bbInView, view.CropBox);
                            return true;
                        }
                    }
                    catch { /* fall through to model-space bbox */ }
                }

                // model-space bbox -> if linked transform to host coords
                BoundingBoxXYZ bbModel = null;
                try { bbModel = elem.get_BoundingBox(null); } catch { bbModel = null; }

                if (bbModel != null)
                {
                    if (!hostDocElementIsHost && linkTransform != null)
                    {
                        var corners = GetBoundingBoxCorners(bbModel);
                        var tMin = new XYZ(double.MaxValue, double.MaxValue, double.MaxValue);
                        var tMax = new XYZ(double.MinValue, double.MinValue, double.MinValue);
                        foreach (var c in corners)
                        {
                            var tc = linkTransform.OfPoint(c);
                            tMin = new XYZ(Math.Min(tMin.X, tc.X), Math.Min(tMin.Y, tc.Y), Math.Min(tMin.Z, tc.Z));
                            tMax = new XYZ(Math.Max(tMax.X, tc.X), Math.Max(tMax.Y, tc.Y), Math.Max(tMax.Z, tc.Z));
                        }
                        var hostBb = new BoundingBoxXYZ { Min = tMin, Max = tMax };
                        if (view.CropBoxActive && view.CropBox != null) return BoundingBoxesIntersect(hostBb, view.CropBox);
                        return true;
                    }
                    else
                    {
                        if (view.CropBoxActive && view.CropBox != null) return BoundingBoxesIntersect(bbModel, view.CropBox);
                        return true;
                    }
                }

                // fallback to representative point
                var rep = GetRepresentativePoint(elem);
                if (rep != null)
                {
                    if (!hostDocElementIsHost && linkTransform != null) rep = linkTransform.OfPoint(rep);
                    return IsPointInViewCrop(rep, view);
                }

                // unknown geometry -> include conservatively
                return true;
            }
            catch { return true; }
        }

        private bool IsPointInViewCrop(XYZ point, View view)
        {
            try
            {
                if (view == null) return true;
                if (view.CropBoxActive && view.CropBox != null)
                {
                    return PointInBoundingBox(view.CropBox.Min, view.CropBox.Max, point);
                }
            }
            catch { return true; }
            return true;
        }

        private bool PointInBoundingBox(XYZ min, XYZ max, XYZ p)
        {
            double minX = Math.Min(min.X, max.X), maxX = Math.Max(min.X, max.X);
            double minY = Math.Min(min.Y, max.Y), maxY = Math.Max(min.Y, max.Y);
            double minZ = Math.Min(min.Z, max.Z), maxZ = Math.Max(min.Z, max.Z);

            const double eps = 1e-9;
            return p.X >= minX - eps && p.X <= maxX + eps
                && p.Y >= minY - eps && p.Y <= maxY + eps
                && p.Z >= minZ - eps && p.Z <= maxZ + eps;
        }

        // --- Parameter reading (unchanged robust logic) ---

        private string GetSpecPositionParameter(Element elem, out string foundFrom, out string foundParamName)
        {
            foundFrom = "None";
            foundParamName = string.Empty;

            string ReadParam(Parameter p, Element owner)
            {
                if (p == null) return null;
                try
                {
                    if (!p.HasValue) return null;
                    switch (p.StorageType)
                    {
                        case StorageType.String: return p.AsString();
                        case StorageType.Integer: return p.AsInteger().ToString();
                        case StorageType.Double: return p.AsDouble().ToString();
                        case StorageType.ElementId:
                            {
                                ElementId id = p.AsElementId();
                                if (id != ElementId.InvalidElementId)
                                {
                                    try { var referenced = owner?.Document?.GetElement(id); return referenced?.Name ?? string.Empty; } catch { return string.Empty; }
                                }
                                break;
                            }
                    }
                }
                catch { }
                return null;
            }

            try
            {
                // 1) direct lookup
                try
                {
                    var p = elem.LookupParameter(SPEC_POSITION_PARAMETER);
                    var v = ReadParam(p, elem);
                    if (!string.IsNullOrEmpty(v)) { foundFrom = "Instance"; foundParamName = p?.Definition?.Name ?? SPEC_POSITION_PARAMETER; return v; }
                }
                catch { }

                // 2) enumerate instance parameters exact
                foreach (Parameter p in elem.Parameters)
                {
                    try
                    {
                        var name = p.Definition?.Name ?? string.Empty;
                        if (string.Equals(name, SPEC_POSITION_PARAMETER, StringComparison.OrdinalIgnoreCase))
                        {
                            var v = ReadParam(p, elem);
                            if (!string.IsNullOrEmpty(v)) { foundFrom = "Instance"; foundParamName = name; return v; }
                        }
                    }
                    catch { }
                }

                // 3) alternatives and heuristics
                string[] alts = { "Spec Position", "SpecPosition", "SPEC_POS", "Specification Position" };
                foreach (var alt in alts)
                {
                    foreach (Parameter p in elem.Parameters)
                    {
                        try
                        {
                            var name = p.Definition?.Name ?? string.Empty;
                            if (string.Equals(name, alt, StringComparison.OrdinalIgnoreCase))
                            {
                                var v = ReadParam(p, elem);
                                if (!string.IsNullOrEmpty(v)) { foundFrom = "Instance"; foundParamName = name; return v; }
                            }
                        }
                        catch { }
                    }
                }

                foreach (Parameter p in elem.Parameters)
                {
                    try
                    {
                        var name = p.Definition?.Name ?? string.Empty;
                        var up = name.ToUpperInvariant();
                        if ((up.Contains("SPEC") || up.Contains("SPECIFICATION")) && (up.Contains("POS") || up.Contains("POSITION")))
                        {
                            var v = ReadParam(p, elem);
                            if (!string.IsNullOrEmpty(v)) { foundFrom = "Instance"; foundParamName = name; return v; }
                        }
                    }
                    catch { }
                }

                // 4) shared param (ExternalDefinition) heuristic
                foreach (Parameter p in elem.Parameters)
                {
                    try
                    {
                        var ext = p.Definition as ExternalDefinition;
                        if (ext != null)
                        {
                            var nm = ext.Name ?? string.Empty;
                            var up = nm.ToUpperInvariant();
                            if ((up.Contains("SPEC") || up.Contains("SPECIFICATION")) && (up.Contains("POS") || up.Contains("POSITION")))
                            {
                                var v = ReadParam(p, elem);
                                if (!string.IsNullOrEmpty(v)) { foundFrom = "Instance (SharedGUID)"; foundParamName = nm; return v; }
                            }
                        }
                    }
                    catch { }
                }

                // 5) type
                try
                {
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        var typeElem = elem.Document.GetElement(typeId);
                        if (typeElem != null)
                        {
                            var tp = typeElem.LookupParameter(SPEC_POSITION_PARAMETER);
                            var tv = ReadParam(tp, typeElem);
                            if (!string.IsNullOrEmpty(tv)) { foundFrom = "Type"; foundParamName = tp?.Definition?.Name ?? SPEC_POSITION_PARAMETER; return tv; }

                            foreach (Parameter p in typeElem.Parameters)
                            {
                                try
                                {
                                    var name = p.Definition?.Name ?? string.Empty;
                                    if (string.Equals(name, SPEC_POSITION_PARAMETER, StringComparison.OrdinalIgnoreCase))
                                    {
                                        var v = ReadParam(p, typeElem);
                                        if (!string.IsNullOrEmpty(v)) { foundFrom = "Type"; foundParamName = name; return v; }
                                    }
                                }
                                catch { }
                            }

                            foreach (Parameter p in typeElem.Parameters)
                            {
                                try
                                {
                                    var name = p.Definition?.Name ?? string.Empty;
                                    var up = name.ToUpperInvariant();
                                    if ((up.Contains("SPEC") || up.Contains("SPECIFICATION")) && (up.Contains("POS") || up.Contains("POSITION")))
                                    {
                                        var v = ReadParam(p, typeElem);
                                        if (!string.IsNullOrEmpty(v)) { foundFrom = "Type"; foundParamName = name; return v; }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }

                // 6) piping system fallback
                try
                {
                    if (elem is Pipe)
                    {
                        var p = elem.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                        var val = ReadParam(p, elem);
                        if (!string.IsNullOrEmpty(val)) { foundFrom = "PipingSystem"; foundParamName = p?.Definition?.Name ?? "RBS_PIPING_SYSTEM_TYPE_PARAM"; return val; }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                if (DebugMode) Log($"Error in GetSpecPositionParameter for elem {elem?.Id.IntegerValue}: {ex}");
            }

            return string.Empty;
        }

        private void DumpElementParameters(Element e)
        {
            if (!DebugMode) return;
            try
            {
                Log($"        PARAM DUMP for ElementId={e.Id.IntegerValue} (Doc='{e.Document?.Title ?? e.Document?.PathName ?? "host"}'):");
                foreach (Parameter p in e.Parameters)
                {
                    try
                    {
                        string pname = p.Definition?.Name ?? "(null)";
                        string pval = p.HasValue ? p.AsValueString() ?? "(empty string)" : "(no value)";
                        var def = p.Definition as ExternalDefinition;
                        string extra = def != null ? $" [GUID={def.GUID}]" : "";
                        Log($"           Param: '{pname}' -> {pval}{extra}");
                    }
                    catch (Exception px)
                    {
                        Log($"           Param read error: {px}");
                    }
                }
            }
            catch (Exception ex) { Log($"        Error dumping parameters: {ex}"); }
        }
    }
}
