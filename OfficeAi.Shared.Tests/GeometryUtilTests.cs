using Xunit;
using OfficeAi.Shared;

public class GeometryUtilTests
{
    [Fact]
    public void ResolveImageSize_BothDimensionsGiven_BothHonoredVerbatim()
    {
        GeometryUtil.ResolveImageSize(400, 300, 200, 50, out float w, out float h);
        Assert.Equal(200, w);
        Assert.Equal(50, h);
    }

    [Fact]
    public void ResolveImageSize_WidthOnly_HeightScalesProportionally()
    {
        // Natural 400x300 (4:3), width halved to 200 -> height should also halve to 150.
        GeometryUtil.ResolveImageSize(400, 300, 200, null, out float w, out float h);
        Assert.Equal(200, w);
        Assert.Equal(150, h);
    }

    [Fact]
    public void ResolveImageSize_HeightOnly_WidthScalesProportionally()
    {
        GeometryUtil.ResolveImageSize(400, 300, null, 150, out float w, out float h);
        Assert.Equal(200, w);
        Assert.Equal(150, h);
    }

    [Fact]
    public void ResolveImageSize_NeitherGiven_ReturnsNaturalSizeUnchanged()
    {
        GeometryUtil.ResolveImageSize(400, 300, null, null, out float w, out float h);
        Assert.Equal(400, w);
        Assert.Equal(300, h);
    }

    [Fact]
    public void ResolveCornerPosition_TopLeft_InsetsFromOrigin()
    {
        GeometryUtil.ResolveCornerPosition(720, 540, 50, 20, 12, "topLeft", out float left, out float top);
        Assert.Equal(12, left);
        Assert.Equal(12, top);
    }

    [Fact]
    public void ResolveCornerPosition_TopRight_InsetsFromRightEdge()
    {
        GeometryUtil.ResolveCornerPosition(720, 540, 50, 20, 12, "topRight", out float left, out float top);
        Assert.Equal(720 - 50 - 12, left);
        Assert.Equal(12, top);
    }

    [Fact]
    public void ResolveCornerPosition_BottomLeft_InsetsFromBottomEdge()
    {
        GeometryUtil.ResolveCornerPosition(720, 540, 50, 20, 12, "bottomLeft", out float left, out float top);
        Assert.Equal(12, left);
        Assert.Equal(540 - 20 - 12, top);
    }

    [Fact]
    public void ResolveCornerPosition_BottomRight_InsetsFromBothFarEdges()
    {
        GeometryUtil.ResolveCornerPosition(720, 540, 50, 20, 12, "bottomRight", out float left, out float top);
        Assert.Equal(720 - 50 - 12, left);
        Assert.Equal(540 - 20 - 12, top);
    }

    [Fact]
    public void ResolveCornerPosition_UnknownCorner_ThrowsWithValidList()
    {
        var ex = Assert.Throws<System.ArgumentException>(() =>
            GeometryUtil.ResolveCornerPosition(720, 540, 50, 20, 12, "middle", out _, out _));
        Assert.Contains("topLeft", ex.Message);
    }
}
