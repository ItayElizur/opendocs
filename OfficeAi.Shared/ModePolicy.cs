using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace OfficeAi.Shared
{
    // Administrator policy: editing modes switched off per app by policy.json in the install folder
    // (Program Files\OpenDocs\policy.json, next to the per-app folders - only administrators can change it there).
    // A broken file disables every mode in every app rather than silently enabling one the admin meant to block.
    public static class ModePolicy
    {
        public const string FileName = "policy.json";

        // Wire names, the same strings the page sends in "set-mode".
        private static readonly string[] AllModes = { "readOnly", "commentOnly", "trackChanges", "fullAutonomy" };
        private static readonly string[] Apps = { "Word", "Excel", "PowerPoint", "Outlook" };

        private static readonly Lazy<Dictionary<string, HashSet<string>>> Policy = new Lazy<Dictionary<string, HashSet<string>>>(() =>
            // BaseDirectory is the add-in's own folder (InstallRoot\WordAiAddIn\); the policy sits one level up, shared by all four apps.
            Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", FileName)));

        /// <summary>Disabled modes (wire names) for an add-in, by its folder name ("WordAiAddIn") or app name ("Word").</summary>
        public static HashSet<string> DisabledFor(string addIn)
        {
            string app = addIn.EndsWith("AiAddIn", StringComparison.Ordinal) ? addIn.Substring(0, addIn.Length - "AiAddIn".Length) : addIn;
            HashSet<string> set;
            return Policy.Value.TryGetValue(app, out set) ? set : new HashSet<string>(AllModes); // unknown add-in: fail closed
        }

        public static string WireName(EditingMode mode)
        {
            return AllModes[(int)mode];
        }

        public static Dictionary<string, HashSet<string>> Load(string path)
        {
            try
            {
                if (!File.Exists(path)) return Uniform(false);
                return Parse(File.ReadAllText(path));
            }
            catch (Exception e)
            {
                DebugLog.Write("ModePolicy: could not read " + path + ": " + e.Message);
                return Uniform(true);
            }
        }

        /// <summary>{ "disabledModes": { "Outlook": ["fullAutonomy"], ... } }, one entry per app ("Word", "Excel",
        /// "PowerPoint", "Outlook"); an app left out has every mode on. App and mode names are case-insensitive; an
        /// unknown name, a wrong shape or invalid JSON disables every mode in every app.</summary>
        public static Dictionary<string, HashSet<string>> Parse(string json)
        {
            var result = Uniform(false);
            try
            {
                using (var doc = JsonDocument.Parse(json, new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true }))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Object) return Uniform(true);
                    if (!doc.RootElement.TryGetProperty("disabledModes", out var byApp)) return result;
                    if (byApp.ValueKind != JsonValueKind.Object) return Uniform(true);
                    foreach (var entry in byApp.EnumerateObject())
                    {
                        string app = Match(Apps, entry.Name);
                        if (app == null || entry.Value.ValueKind != JsonValueKind.Array) return Uniform(true);
                        foreach (var item in entry.Value.EnumerateArray())
                        {
                            string mode = item.ValueKind == JsonValueKind.String ? Match(AllModes, item.GetString()) : null;
                            if (mode == null) return Uniform(true);
                            result[app].Add(mode);
                        }
                    }
                }
                return result;
            }
            catch (JsonException)
            {
                return Uniform(true);
            }
        }

        private static string Match(string[] names, string value)
        {
            return Array.Find(names, n => string.Equals(n, value.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        // Every app with no mode disabled, or (allOff) with every mode disabled.
        private static Dictionary<string, HashSet<string>> Uniform(bool allOff)
        {
            var d = new Dictionary<string, HashSet<string>>();
            foreach (var app in Apps) d[app] = allOff ? new HashSet<string>(AllModes) : new HashSet<string>();
            return d;
        }
    }
}
