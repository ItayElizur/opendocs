using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Chart-type name to Office's xlChartType code, shared by Word, Excel and
    /// PowerPoint. Single-sourced to avoid the apps' tables drifting apart;
    /// add a chart type here AND to each app's entry.ts chartType enum
    /// (ChartTypesTests guards the two stay in step). Full history: ChartTypes.cs.md.
    /// </summary>
    public static class ChartTypes
    {
        public static readonly Dictionary<string, int> ByName = new Dictionary<string, int>
        {
            ["column"] = 51,        // xlColumnClustered
            ["columnStacked"] = 52, // xlColumnStacked
            ["bar"] = 57,           // xlBarClustered
            ["barStacked"] = 58,    // xlBarStacked
            ["line"] = 4,           // xlLine
            ["area"] = 1,           // xlArea
            ["pie"] = 5,            // xlPie
            ["doughnut"] = -4120,   // xlDoughnut
        };
    }
}
