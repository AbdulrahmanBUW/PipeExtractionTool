using System.Collections.Generic;

namespace PipeExtractionTool
{
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