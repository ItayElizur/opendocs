using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    // Each add-in lists its tools in web-src/tools.ts and the English/Hebrew labels for them in web-src/display.ts.
    // The display type only checks the entries that exist, so these tests catch a tool with no label (the settings
    // screen would show its raw name) and a label left behind for a tool that was removed.
    public class ToolDisplayCoverageTests
    {
        private static string Read(string relative)
        {
            string dir = AppContext.BaseDirectory;
            while (dir != null && !File.Exists(Path.Combine(dir, "CLAUDE.md"))) dir = Path.GetDirectoryName(dir);
            Assert.True(dir != null, "repo root (CLAUDE.md) not found above the test folder");
            string text = File.ReadAllText(Path.Combine(dir, relative));
            return Regex.Replace(text, @"//[^\r\n]*", ""); // line comments can hold quoted words
        }

        private static void Check(string app)
        {
            string folder = app + "AiAddIn/web-src/";
            var tools = new HashSet<string>(Regex.Matches(Read(folder + "tools.ts"), @"\bname: '([a-z][a-z_]*)'")
                .Cast<Match>().Select(m => m.Groups[1].Value));
            var labelled = new HashSet<string>(Regex.Matches(Read(folder + "display.ts"), @"^  ([a-z][a-z_]*): (?:\{|d\()", RegexOptions.Multiline)
                .Cast<Match>().Select(m => m.Groups[1].Value));

            Assert.True(tools.Count > 0 && labelled.Count > 0, app + ": one side is empty");
            Assert.True(tools.SetEquals(labelled),
                app + ": tools without a label [" + string.Join(", ", tools.Except(labelled)) +
                "], labels without a tool [" + string.Join(", ", labelled.Except(tools)) + "]");
        }

        [Fact] public void Word() { Check("Word"); }
        [Fact] public void Excel() { Check("Excel"); }
        [Fact] public void PowerPoint() { Check("PowerPoint"); }
        [Fact] public void Outlook() { Check("Outlook"); }
    }
}
