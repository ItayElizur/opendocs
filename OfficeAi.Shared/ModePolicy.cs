using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace OfficeAi.Shared
{
    // Administrator policy from policy.json in the install folder (Program Files\OpenDocs\policy.json, next to the
    // per-app folders - only administrators can change it there), per app:
    //   disabledModes   - editing modes switched off (enforced here, in PaneHostBase, and hidden in the page);
    //   requireApproval - tool -> modes in which the user must Approve the call first ("*" = every tool that is not
    //                     read-only; Full autonomy never asks). Enforced by the page only; see ModePolicy.cs.md.
    // A wrong app or mode name, or a broken file, switches every mode off in every app rather than silently leaving
    // one open. Tool names are not checked: a name the app does not have is ignored (the administrator's to get right).
    public static class ModePolicy
    {
        public const string FileName = "policy.json";

        // Wire names, the same strings the page sends in "set-mode".
        private static readonly string[] AllModes = { "readOnly", "commentOnly", "trackChanges", "fullAutonomy" };
        private static readonly string[] Apps = { "Word", "Excel", "PowerPoint", "Outlook" };

        public sealed class Rules
        {
            /// <summary>App -> disabled modes.</summary>
            public Dictionary<string, HashSet<string>> Disabled = new Dictionary<string, HashSet<string>>();
            /// <summary>App -> tool (or "*") -> modes that need an Approve.</summary>
            public Dictionary<string, Dictionary<string, HashSet<string>>> Approvals = new Dictionary<string, Dictionary<string, HashSet<string>>>();
        }

        private static readonly Lazy<Rules> Policy = new Lazy<Rules>(() =>
            // BaseDirectory is the add-in's own folder (InstallRoot\WordAiAddIn\); the policy sits one level up, shared by all four apps.
            Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", FileName)));

        /// <summary>Disabled modes (wire names) for an add-in, by its folder name ("WordAiAddIn") or app name ("Word").</summary>
        public static HashSet<string> DisabledFor(string addIn)
        {
            HashSet<string> set;
            return Policy.Value.Disabled.TryGetValue(AppOf(addIn), out set) ? set : new HashSet<string>(AllModes); // unknown add-in: fail closed
        }

        /// <summary>Tool (or "*") -> modes that need an Approve, for an add-in.</summary>
        public static Dictionary<string, HashSet<string>> ApprovalsFor(string addIn)
        {
            Dictionary<string, HashSet<string>> map;
            return Policy.Value.Approvals.TryGetValue(AppOf(addIn), out map) ? map : new Dictionary<string, HashSet<string>>();
        }

        internal static string AppOf(string addIn)
        {
            return addIn.EndsWith("AiAddIn", StringComparison.Ordinal) ? addIn.Substring(0, addIn.Length - "AiAddIn".Length) : addIn;
        }

        /// <summary>True when a tool call must be refused: the policy switches some mode off and the pane is in one of
        /// them, or has not reported a mode yet.</summary>
        public static bool Blocks(HashSet<string> disabled, EditingMode? mode)
        {
            return disabled.Count > 0 && (mode == null || disabled.Contains(WireName(mode.Value)));
        }

        public static string WireName(EditingMode mode)
        {
            return AllModes[(int)mode];
        }

        public static Rules Load(string path)
        {
            try
            {
                if (!File.Exists(path)) return Empty(false);
                return Parse(File.ReadAllText(path));
            }
            catch (Exception e)
            {
                DebugLog.Write("ModePolicy: could not read " + path + ": " + e.Message);
                return Empty(true);
            }
        }

        /// <summary>{ "disabledModes": { "Outlook": ["fullAutonomy"] }, "requireApproval": { "Excel": { "*": ["trackChanges"] } } }.
        /// App and mode names are case-insensitive; tool names are as the tools are named. An app left out has no rule.
        /// An unknown app or mode, a wrong shape or invalid JSON switches every mode off in every app; tool names are
        /// taken as written and one the app does not have is ignored by the page.</summary>
        public static Rules Parse(string json)
        {
            var result = Empty(false);
            try
            {
                using (var doc = JsonDocument.Parse(json, new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true }))
                {
                    var root = doc.RootElement;
                    if (root.ValueKind != JsonValueKind.Object) return Empty(true);

                    if (root.TryGetProperty("disabledModes", out var disabled))
                    {
                        if (disabled.ValueKind != JsonValueKind.Object) return Empty(true);
                        foreach (var entry in disabled.EnumerateObject())
                        {
                            string app = Match(Apps, entry.Name);
                            if (app == null || !AddModes(entry.Value, result.Disabled[app])) return Empty(true);
                        }
                    }

                    if (root.TryGetProperty("requireApproval", out var approvals))
                    {
                        if (approvals.ValueKind != JsonValueKind.Object) return Empty(true);
                        foreach (var entry in approvals.EnumerateObject())
                        {
                            string app = Match(Apps, entry.Name);
                            if (app == null || entry.Value.ValueKind != JsonValueKind.Object) return Empty(true);
                            foreach (var tool in entry.Value.EnumerateObject())
                            {
                                var modes = new HashSet<string>();
                                if (!AddModes(tool.Value, modes)) return Empty(true);
                                result.Approvals[app][tool.Name] = modes;
                            }
                        }
                    }
                }
                return result;
            }
            catch (JsonException)
            {
                return Empty(true);
            }
        }

        // Adds an array of mode names to `into`; false on anything that is not an array of known mode names.
        private static bool AddModes(JsonElement list, HashSet<string> into)
        {
            if (list.ValueKind != JsonValueKind.Array) return false;
            foreach (var item in list.EnumerateArray())
            {
                string mode = item.ValueKind == JsonValueKind.String ? Match(AllModes, item.GetString()) : null;
                if (mode == null) return false;
                into.Add(mode);
            }
            return true;
        }

        private static string Match(string[] names, string value)
        {
            return Array.Find(names, n => string.Equals(n, value.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        // No rule in any app, or (allOff) every mode switched off in every app.
        private static Rules Empty(bool allOff)
        {
            var r = new Rules();
            foreach (var app in Apps)
            {
                r.Disabled[app] = allOff ? new HashSet<string>(AllModes) : new HashSet<string>();
                r.Approvals[app] = new Dictionary<string, HashSet<string>>();
            }
            return r;
        }
    }
}
