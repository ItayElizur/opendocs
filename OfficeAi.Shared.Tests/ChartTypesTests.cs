using Xunit;
using OfficeAi.Shared;

public class ChartTypesTests
{
    // Every chart-type name advertised by any app's entry.ts. All three now
    // advertise the same 8 (Word gained "barStacked" to reach parity):
    //   Word:       edit_chart.chartType
    //   Excel:      add_chart.chartType, edit_chart.chartType
    //   PowerPoint: add_chart.kind, edit_chart.chartType
    private static readonly string[] EntryTsNames =
    {
        "column", "columnStacked", "bar", "barStacked", "line", "area", "pie", "doughnut",
    };

    // Unlike ShapeTypes (where asserting raw ints would be brittle and
    // meaningless), asserting exact values here is the whole point: these are
    // documented xlChartType constants, and a wrong one is a SILENT wrong
    // result - a successful call that draws the wrong chart. That bug has
    // actually shipped in this repo: PowerPoint's copy mapped "bar" to 51
    // (xlColumnClustered) instead of 57 (xlBarClustered), so chartType:'bar'
    // produced a column chart and reported success.
    [Theory]
    [InlineData("column", 51)]        // xlColumnClustered
    [InlineData("columnStacked", 52)] // xlColumnStacked
    [InlineData("bar", 57)]           // xlBarClustered - NOT 51
    [InlineData("barStacked", 58)]    // xlBarStacked   - NOT 52
    [InlineData("line", 4)]           // xlLine
    [InlineData("area", 1)]           // xlArea
    [InlineData("pie", 5)]            // xlPie
    [InlineData("doughnut", -4120)]   // xlDoughnut
    public void ByName_MapsEachNameToItsDocumentedXlChartTypeCode(string name, int expected)
    {
        Assert.Equal(expected, ChartTypes.ByName[name]);
    }

    [Fact]
    public void ByName_BarIsDistinctFromColumn()
    {
        // The exact confusion behind the shipped bug - guard it explicitly
        // rather than relying on the table above being read carefully.
        Assert.NotEqual(ChartTypes.ByName["column"], ChartTypes.ByName["bar"]);
        Assert.NotEqual(ChartTypes.ByName["columnStacked"], ChartTypes.ByName["barStacked"]);
    }

    [Fact]
    public void ByName_ContainsEveryNameTheAppsEntryTsAdvertise()
    {
        foreach (string name in EntryTsNames)
        {
            Assert.True(ChartTypes.ByName.ContainsKey(name), "Missing chart type: " + name);
        }
    }

    [Fact]
    public void ByName_AdvertisesNothingTheAppsDoNot()
    {
        // The other direction of the drift guard: a name here but absent from
        // every entry.ts enum would be a capability no model can reach.
        foreach (string key in ChartTypes.ByName.Keys)
        {
            Assert.Contains(key, EntryTsNames);
        }
    }

    [Fact]
    public void ByName_UnknownKey_ReturnsFalse()
    {
        int value;
        Assert.False(ChartTypes.ByName.TryGetValue("not-a-chart-type", out value));
    }
}
