using Xunit;
using OfficeAi.Shared;

public class ShapeTypesTests
{
    // Excel's add_shape shapeType enum (ExcelAiAddIn/web-src/entry.ts), minus
    // "textbox" (handled separately, not an AutoShape).
    private static readonly string[] ExcelEntryTsNames =
    {
        "rect", "roundRect", "ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid",
        "diamond", "pentagon", "hexagon", "octagon", "pie", "chord", "donut", "foldedCorner", "heart",
        "lightningBolt", "sun", "moon", "cloud", "arc", "star5", "rightArrow", "leftArrow", "upArrow", "downArrow",
    };

    // PowerPoint's add_shape shapeType enum (PowerPointAiAddIn/web-src/entry.ts).
    private static readonly string[] PowerPointEntryTsNames =
    {
        "rect", "rectangle", "roundRect", "ellipse", "oval", "triangle", "rtTriangle",
        "parallelogram", "trapezoid", "diamond", "pentagon", "hexagon", "octagon",
        "pie", "chord", "donut", "foldedCorner", "heart", "lightningBolt", "sun", "moon",
        "cloud", "arc", "star5", "rightArrow", "leftArrow", "upArrow", "downArrow",
    };

    [Fact]
    public void ByName_KnownKey_Resolves()
    {
        int value;
        Assert.True(ShapeTypes.ByName.TryGetValue("rect", out value));
    }

    [Fact]
    public void ByName_LookupIsCaseInsensitive()
    {
        int fromExact, fromMixedCase;
        Assert.True(ShapeTypes.ByName.TryGetValue("RoundRect", out fromExact));
        Assert.True(ShapeTypes.ByName.TryGetValue("roundrect", out fromMixedCase));
        Assert.Equal(fromExact, fromMixedCase);
    }

    [Fact]
    public void ByName_RectangleAlias_ResolvesToSameValueAsRect()
    {
        Assert.Equal(ShapeTypes.ByName["rect"], ShapeTypes.ByName["rectangle"]);
    }

    [Fact]
    public void ByName_OvalAlias_ResolvesToSameValueAsEllipse()
    {
        Assert.Equal(ShapeTypes.ByName["ellipse"], ShapeTypes.ByName["oval"]);
    }

    [Fact]
    public void ByName_UnknownKey_ReturnsFalse()
    {
        int value;
        Assert.False(ShapeTypes.ByName.TryGetValue("not-a-real-shape", out value));
    }

    [Fact]
    public void ByName_ContainsEveryNameExcelsEntryTsAdvertises()
    {
        foreach (string name in ExcelEntryTsNames)
        {
            Assert.True(ShapeTypes.ByName.ContainsKey(name), "Missing key: " + name);
        }
    }

    [Fact]
    public void ByName_ContainsEveryNamePowerPointsEntryTsAdvertises()
    {
        foreach (string name in PowerPointEntryTsNames)
        {
            Assert.True(ShapeTypes.ByName.ContainsKey(name), "Missing key: " + name);
        }
    }
}
