using System.Collections.Generic;
using Xunit;
using OfficeAi.Shared;

public class ActionHistoryTests
{
    private sealed class Entry
    {
        public string Id;
        public Entry(string id) { Id = id; }
    }

    private static Entry TakeUndo(ActionHistory<Entry> h)
    {
        Entry e;
        string barrier;
        Assert.True(h.TryTakeUndo(out e, out barrier));
        Assert.Null(barrier);
        return e;
    }

    [Fact]
    public void Empty_NothingToUndoOrRedo()
    {
        var h = new ActionHistory<Entry>();
        Entry e;
        string barrier;
        Assert.False(h.TryTakeUndo(out e, out barrier));
        Assert.Null(e);
        Assert.Null(barrier);
        Assert.False(h.TryTakeRedo(out e));
    }

    [Fact]
    public void Undo_ReturnsNewestFirst()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        h.Push(new Entry("b"));
        Assert.Equal("b", TakeUndo(h).Id);
        Assert.Equal("a", TakeUndo(h).Id);
    }

    [Fact]
    public void UndoThenRedo_RoundTripsInOrder()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        h.Push(new Entry("b"));
        h.CompleteUndo(TakeUndo(h)); // b
        h.CompleteUndo(TakeUndo(h)); // a

        Entry e;
        Assert.True(h.TryTakeRedo(out e));
        Assert.Equal("a", e.Id);
        h.CompleteRedo(e);
        Assert.True(h.TryTakeRedo(out e));
        Assert.Equal("b", e.Id);
        h.CompleteRedo(e);

        Assert.Equal(0, h.RedoCount);
        Assert.Equal("b", TakeUndo(h).Id);
    }

    [Fact]
    public void Push_ClearsRedo()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        h.CompleteUndo(TakeUndo(h));
        Assert.Equal(1, h.RedoCount);

        h.Push(new Entry("b"));
        Assert.Equal(0, h.RedoCount);
    }

    [Fact]
    public void FailedUndo_DroppedEntryIsNotRedoable()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        TakeUndo(h); // caller failed - no CompleteUndo
        Entry e;
        Assert.False(h.TryTakeRedo(out e));
        Assert.Equal(0, h.UndoCount);
    }

    [Fact]
    public void Barrier_BlocksUndoAndStaysInPlace()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        h.PushBarrier("send_email");

        for (int i = 0; i < 2; i++)
        {
            Entry e;
            string barrier;
            Assert.False(h.TryTakeUndo(out e, out barrier));
            Assert.Null(e);
            Assert.Equal("send_email", barrier);
        }
        Assert.Equal(2, h.UndoCount);
    }

    [Fact]
    public void Barrier_ActionsAfterItAreStillUndoable()
    {
        var h = new ActionHistory<Entry>();
        h.PushBarrier("send_email");
        h.Push(new Entry("b"));
        Assert.Equal("b", TakeUndo(h).Id);

        Entry e;
        string barrier;
        Assert.False(h.TryTakeUndo(out e, out barrier));
        Assert.Equal("send_email", barrier);
    }

    [Fact]
    public void Barrier_ClearsRedo()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("a"));
        h.CompleteUndo(TakeUndo(h));
        h.PushBarrier("send_email");
        Assert.Equal(0, h.RedoCount);
    }

    [Fact]
    public void Capacity_DropsOldest()
    {
        var h = new ActionHistory<Entry>(2);
        h.Push(new Entry("a"));
        h.Push(new Entry("b"));
        h.Push(new Entry("c"));
        Assert.Equal(2, h.UndoCount);
        Assert.Equal("c", TakeUndo(h).Id);
        Assert.Equal("b", TakeUndo(h).Id);
        Entry e;
        string barrier;
        Assert.False(h.TryTakeUndo(out e, out barrier));
    }

    [Fact]
    public void ForEachEntry_VisitsBothStacksAndSkipsBarriers()
    {
        var h = new ActionHistory<Entry>();
        h.Push(new Entry("old"));
        h.PushBarrier("send_email");
        h.Push(new Entry("old"));
        h.Push(new Entry("other"));
        h.CompleteUndo(TakeUndo(h)); // "other" now on redo

        var seen = new List<string>();
        h.ForEachEntry(e => { seen.Add(e.Id); if (e.Id == "old") e.Id = "new"; });
        Assert.Equal(3, seen.Count);

        var after = new List<string>();
        h.ForEachEntry(e => after.Add(e.Id));
        Assert.Equal(new[] { "new", "new", "other" }, after);
    }
}
