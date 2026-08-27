using System;
using Xunit;
using OfficeAi.Shared;

public class ComRetryTests
{
    // The three HRESULTs that mean "Office was busy, ask again" rather than
    // "your request was wrong". Asserting the exact values matters: the whole
    // design is an allowlist, and a wrong constant either retries something
    // that should fail fast, or fails fast on something that should retry.
    [Theory]
    [InlineData(unchecked((int)0x800706BE))] // RPC_S_CALL_FAILED
    [InlineData(unchecked((int)0x8001010A))] // RPC_E_SERVERCALL_RETRYLATER
    [InlineData(unchecked((int)0x800706BA))] // RPC_S_SERVER_UNAVAILABLE
    public void IsTransient_RecognisesTheKnownTransientHResults(int hr)
    {
        Assert.True(ComRetry.IsTransient(hr));
    }

    [Fact]
    public void IsTransient_RejectsEverythingElse()
    {
        // E_INVALIDARG - a genuine "your request was wrong" error, which must
        // NOT be retried.
        Assert.False(ComRetry.IsTransient(unchecked((int)0x80070057)));
        Assert.False(ComRetry.IsTransient(0));
    }

    [Fact]
    public void Run_SucceedsFirstTime_InvokesActionExactlyOnce()
    {
        int calls = 0;
        ComRetry.Run(() => calls++, "test");
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Run_NonTransientFailure_ThrowsImmediatelyWithoutRetrying()
    {
        // The point of the allowlist: a real logic error must fail on attempt
        // 1, not be masked behind 3 attempts and ~600ms of delay.
        int calls = 0;
        Assert.Throws<ArgumentException>(() => ComRetry.Run(() =>
        {
            calls++;
            throw new ArgumentException("bad range");
        }, "test"));
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Run_TransientFailureThatThenSucceeds_Retries()
    {
        int calls = 0;
        ComRetry.Run(() =>
        {
            calls++;
            if (calls < 2) throw new COMLikeException(unchecked((int)0x800706BE));
        }, "test");
        Assert.Equal(2, calls);
    }

    [Fact]
    public void Run_TransientFailureEveryTime_GivesUpAfterThreeAttempts()
    {
        int calls = 0;
        Assert.Throws<COMLikeException>(() => ComRetry.Run(() =>
        {
            calls++;
            throw new COMLikeException(unchecked((int)0x800706BE));
        }, "test"));
        Assert.Equal(3, calls);
    }

    // A COM error surfaced through dynamic late-binding is not guaranteed to
    // arrive as a raw COMException, which is why ComRetry filters on HResult
    // rather than on exception type. This stand-in proves that.
    private class COMLikeException : Exception
    {
        public COMLikeException(int hr) { HResult = hr; }
    }
}

public class SmartArtLayoutsTests
{
    [Fact]
    public void ByName_ContainsTheSevenSupportedLayouts()
    {
        Assert.Equal(7, SmartArtLayouts.ByName.Count);
        foreach (string key in new[] { "list", "process", "cycle", "hierarchy", "pyramid", "matrix", "venn" })
        {
            Assert.True(SmartArtLayouts.ByName.ContainsKey(key), "Missing layout: " + key);
        }
    }

    [Fact]
    public void DisplayNameFor_KnownKey_ReturnsGalleryName()
    {
        Assert.Equal("Organization Chart", SmartArtLayouts.DisplayNameFor("hierarchy", "add_smartart"));
    }

    [Fact]
    public void DisplayNameFor_UnknownKey_ThrowsListingValidKeys()
    {
        var ex = Assert.Throws<ArgumentException>(() => SmartArtLayouts.DisplayNameFor("bogus", "add_smartart"));
        Assert.Contains("bogus", ex.Message);
        Assert.Contains("hierarchy", ex.Message);   // lists what IS valid
        Assert.StartsWith("add_smartart:", ex.Message); // names the calling tool
    }

    [Fact]
    public void NotInGallery_MessageNamesTheLocalisationCause()
    {
        // This error is deliberately distinct from the unknown-key one - it is
        // the only diagnosis that points at a non-English Office install.
        var ex = SmartArtLayouts.NotInGallery("Basic Venn", "add_smartart");
        Assert.Contains("Basic Venn", ex.Message);
        Assert.Contains("non-English", ex.Message);
    }
}
