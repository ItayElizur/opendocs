using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Chart-type name to Office's xlChartType code, shared by Word, Excel and
    /// PowerPoint. All three drive the same Office charting engine and so use
    /// the same numeric codes.
    ///
    /// Previously three separate copies, one per app - a duplication that had
    /// already caused one shipped bug and one capability gap:
    ///   - PptChartTypeMap's "bar" was independently wrong (51/xlColumnClustered
    ///     instead of 57/xlBarClustered) and had to be fixed on that side alone.
    ///   - WordChartTypeMap was missing "barStacked" entirely, so Word could not
    ///     produce a stacked bar chart while the other two apps could.
    /// Both are why this table is now single-source. Add a chart type HERE, and
    /// to each app's entry.ts chartType enum - ChartTypesTests guards that the
    /// two stay in step.
    ///
    /// Plain int values, no Office interop type anywhere - so unlike ShapeTypes
    /// this needed no special handling to cross the assembly boundary; it was
    /// already int-valued in all three apps.
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
