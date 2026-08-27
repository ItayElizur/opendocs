using System;
using System.Collections.Generic;
using System.Text.Json;
using Xunit;
using OfficeAi.Shared;

public class ToolArgsTests
{
    private static readonly Dictionary<string, string[]> Required = new Dictionary<string, string[]>
    {
        ["do_thing"] = new[] { "a", "b" },
    };

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement;

    [Fact]
    public void ValidateRequired_UnknownKind_IsASilentNoOp()
    {
        // Deliberate current behavior: a kind with no entry in RequiredFields
        // is simply not validated, rather than treated as an error.
        ToolArgs.ValidateRequired("no_such_kind", Parse("{}"), Required, "Command");
    }

    [Fact]
    public void ValidateRequired_AllFieldsPresent_Passes()
    {
        ToolArgs.ValidateRequired("do_thing", Parse("{\"a\":1,\"b\":2}"), Required, "Command");
    }

    [Fact]
    public void ValidateRequired_MissingField_ThrowsWithExactMessageAndNoun()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ToolArgs.ValidateRequired("do_thing", Parse("{\"a\":1}"), Required, "Command"));
        Assert.Equal("Command \"do_thing\" is missing required field \"b\".", ex.Message);
    }

    [Fact]
    public void ValidateRequired_UsesTheGivenNoun()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ToolArgs.ValidateRequired("do_thing", Parse("{\"a\":1}"), Required, "Operation"));
        Assert.StartsWith("Operation ", ex.Message);
    }

    [Fact]
    public void ValidateKnownFields_KnownSet_Passes()
    {
        var known = new HashSet<string> { "bold", "italic" };
        var fields = new HashSet<string> { "bold" };
        ToolArgs.ValidateKnownFields(fields, known, "updateTextStyle");
    }

    [Fact]
    public void ValidateKnownFields_UnknownField_ThrowsListingValidFields()
    {
        var known = new HashSet<string> { "bold", "italic" };
        var fields = new HashSet<string> { "bogus" };
        var ex = Assert.Throws<ArgumentException>(() => ToolArgs.ValidateKnownFields(fields, known, "updateTextStyle"));
        Assert.Contains("bold", ex.Message);
        Assert.Contains("italic", ex.Message);
        Assert.Contains("bogus", ex.Message);
    }
}
