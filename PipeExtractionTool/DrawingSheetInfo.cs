using Autodesk.Revit.DB;

namespace PipeExtractionTool
{
    public class DrawingSheetInfo
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
        public string Title { get; set; }
        public string Discipline { get; set; }
        public ViewSheet ViewSheet { get; set; }
        public bool IsSelected { get; set; } = false;

        public string DisplayName => $"{Number} - {Name}";

        public override string ToString()
        {
            return DisplayName;
        }
    }
}