using System.Collections.Generic;
using Xunit;
using OfficeAi.Shared;

public class ObjectAnchorsTests
{
    private static List<ObjectAnchor> Anchors(params int[] starts)
    {
        var list = new List<ObjectAnchor>();
        foreach (int s in starts) list.Add(new ObjectAnchor(s, "o" + s));
        return list;
    }

    [Fact]
    public void LabelsFor_AnchorInsideParagraph_IsAttached()
    {
        var a = Anchors(5);
        int next = 0;
        Assert.Equal("  {o5}", ObjectAnchors.LabelsFor(a, ref next, 0, 10, 8));
    }

    [Fact]
    public void LabelsFor_NoAnchorInParagraph_ReturnsEmpty()
    {
        var a = Anchors(50);
        int next = 0;
        Assert.Equal("", ObjectAnchors.LabelsFor(a, ref next, 0, 10, 8));
        Assert.Equal(0, next);
    }

    [Fact]
    public void LabelsFor_AnchorAtParagraphEnd_BelongsToNextParagraph()
    {
        var a = Anchors(10);
        int next = 0;
        Assert.Equal("", ObjectAnchors.LabelsFor(a, ref next, 0, 10, 8));
        Assert.Equal("  {o10}", ObjectAnchors.LabelsFor(a, ref next, 10, 20, 8));
    }

    [Fact]
    public void LabelsFor_AnchorsBeforeRequestedRange_AreSkipped()
    {
        var a = Anchors(2, 30);
        int next = 0;
        // Reading starts at the paragraph [20, 40): the anchor at 2 is outside the range.
        Assert.Equal("  {o30}", ObjectAnchors.LabelsFor(a, ref next, 20, 40, 8));
    }

    [Fact]
    public void LabelsFor_OverCap_ReportsTheRemainderCount()
    {
        var a = Anchors(1, 2, 3, 4);
        int next = 0;
        Assert.Equal("  {o1, o2, +2 more}", ObjectAnchors.LabelsFor(a, ref next, 0, 10, 2));
    }

    [Fact]
    public void LabelsFor_SameStartTwice_BothAttached()
    {
        var a = Anchors(4, 4);
        int next = 0;
        Assert.Equal("  {o4, o4}", ObjectAnchors.LabelsFor(a, ref next, 0, 10, 8));
    }
}
