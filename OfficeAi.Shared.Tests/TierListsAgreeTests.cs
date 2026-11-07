using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    // The editing-mode tool lists exist twice: in C# (what the executor enforces) and in each add-in's entry.ts (what
    // the model is shown). CLAUDE.md warns they drift silently; these tests read both source files and compare.
    public class TierListsAgreeTests
    {
        private static string Repo()
        {
            string dir = AppContext.BaseDirectory;
            while (dir != null && !File.Exists(Path.Combine(dir, "CLAUDE.md"))) dir = Path.GetDirectoryName(dir);
            Assert.True(dir != null, "repo root (CLAUDE.md) not found above the test folder");
            return dir;
        }

        private static string Read(string relative)
        {
            string text = File.ReadAllText(Path.Combine(Repo(), relative));
            return Regex.Replace(text, @"//[^\r\n]*", ""); // drop line comments so quoted words in them are not picked up
        }

        private static HashSet<string> CSharpSet(string relative, string field)
        {
            Match m = Regex.Match(Read(relative), field + @"\s*=\s*new\s+(?:System\.Collections\.Generic\.)?HashSet<string>\s*\{(.*?)\};", RegexOptions.Singleline);
            Assert.True(m.Success, field + " not found in " + relative);
            return new HashSet<string>(Regex.Matches(m.Groups[1].Value, "\"([^\"]+)\"").Cast<Match>().Select(x => x.Groups[1].Value));
        }

        private static HashSet<string> TsArray(string relative, string startMarker)
        {
            string text = Read(relative);
            int start = text.IndexOf(startMarker, StringComparison.Ordinal);
            Assert.True(start >= 0, startMarker + " not found in " + relative);
            int open = text.IndexOf('[', start);
            int close = text.IndexOf(']', open);
            return new HashSet<string>(Regex.Matches(text.Substring(open, close - open), "'([^']+)'").Cast<Match>().Select(x => x.Groups[1].Value));
        }

        private static void Same(HashSet<string> csharp, HashSet<string> ts, string what)
        {
            Assert.True(csharp.Count > 0 && ts.Count > 0, what + ": one side is empty");
            Assert.True(csharp.SetEquals(ts),
                what + ": only in C# [" + string.Join(", ", csharp.Except(ts)) + "] only in entry.ts [" + string.Join(", ", ts.Except(csharp)) + "]");
        }

        [Fact]
        public void Word_ReadOnlyTools()
        {
            Same(CSharpSet("WordAiAddIn/WordTools.cs", "AlwaysAllowedTools"),
                 TsArray("WordAiAddIn/web-src/entry.ts", "readOnlyTools:"), "Word read-only tools");
        }

        [Fact]
        public void Excel_ReadOnlyTools()
        {
            Same(CSharpSet("ExcelAiAddIn/ExcelTools.cs", "AlwaysAllowedTools"),
                 TsArray("ExcelAiAddIn/web-src/tools.ts", "const READ_ONLY_TOOL_NAMES ="), "Excel read-only tools");
        }

        [Fact]
        public void PowerPoint_ReadOnlyTools()
        {
            string ts = Read("PowerPointAiAddIn/web-src/tools.ts");
            int from = ts.IndexOf("const READER_TOOLS = [", StringComparison.Ordinal);
            int to = ts.IndexOf("const MUTATION_TOOLS = [", StringComparison.Ordinal);
            Assert.True(from >= 0 && to > from, "READER_TOOLS / MUTATION_TOOLS markers not found");
            var names = new HashSet<string>(Regex.Matches(ts.Substring(from, to - from), @"^    name: '([^']+)'", RegexOptions.Multiline)
                .Cast<Match>().Select(x => x.Groups[1].Value));
            Same(CSharpSet("PowerPointAiAddIn/PowerPointTools.cs", "AlwaysAllowedTools"), names, "PowerPoint read-only tools");
        }

        [Fact]
        public void Outlook_AllTiers()
        {
            const string cs = "OutlookAiAddIn/OutlookTools.cs";
            const string ts = "OutlookAiAddIn/web-src/entry.ts";
            Same(CSharpSet(cs, "AlwaysAllowedTools"), TsArray(ts, "readOnlyTools:"), "Outlook read-only tools");
            Same(CSharpSet(cs, "DraftTierTools"), TsArray(ts, "commentOnlyExtraTools:"), "Outlook draft tier");
            Same(CSharpSet(cs, "ApprovalTierTools"), TsArray(ts, "trackChangesExtraTools:"), "Outlook approval tier");

            // The send tools are advertised only in Full autonomy, so no lower list may mention them.
            var lower = TsArray(ts, "readOnlyTools:").Union(TsArray(ts, "commentOnlyExtraTools:")).Union(TsArray(ts, "trackChangesExtraTools:"));
            Assert.Empty(CSharpSet(cs, "SendTierTools").Intersect(lower));
        }
    }
}
