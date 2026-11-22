using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeAi.Shared
{
    // One skill folder (Agent Skills format: <name>\SKILL.md plus any other files), parsed once when the pane starts.
    public sealed class Skill
    {
        public string Name;
        public string Description;
        public string ArgumentHint;
        public string Dir;
        public string Body;
        public List<string> Files;
        public bool ModelInvocable;
        public bool UserInvocable;
        public string Source;
    }

    // Reads an app's skills from the machine folder (InstallRoot\skills\<App>\, next to policy.json) and the user's
    // folder (%LOCALAPPDATA%\<App>AiAddIn\Skills\), and serves read_skill_file. The model only ever names a skill and a
    // file from that skill's scanned list, so it can never pick the folder it reads from. See SkillCatalog.cs.md.
    public static class SkillCatalog
    {
        public const string ReadToolName = "read_skill_file";
        public const int MaxFiles = 200;
        public const int MaxFileBytes = 256 * 1024;

        private static readonly Regex NamePattern = new Regex("^[a-z0-9]+(-[a-z0-9]+)*$");
        private static readonly Dictionary<string, List<Skill>> Cache = new Dictionary<string, List<Skill>>();
        private static readonly object CacheLock = new object();

        /// <summary>The add-in's skills, read on first use and kept for the life of the Office process.</summary>
        public static List<Skill> For(string addIn)
        {
            lock (CacheLock)
            {
                List<Skill> skills;
                if (!Cache.TryGetValue(addIn, out skills))
                {
                    string machine = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "skills", ModePolicy.AppOf(addIn));
                    string user = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), addIn, "Skills");
                    skills = Load(machine, user);
                    Cache[addIn] = skills;
                }
                return skills;
            }
        }

        /// <summary>Both folders merged; a machine skill wins over a user skill of the same name.</summary>
        public static List<Skill> Load(string machineRoot, string userRoot)
        {
            var byName = new Dictionary<string, Skill>(StringComparer.Ordinal);
            var result = new List<Skill>();
            foreach (var root in new[] { Tuple.Create(machineRoot, "machine"), Tuple.Create(userRoot, "user") })
            {
                foreach (Skill s in LoadRoot(root.Item1, root.Item2))
                {
                    if (byName.ContainsKey(s.Name)) { DebugLog.Write("skills: " + s.Dir + " skipped, '" + s.Name + "' already exists in " + byName[s.Name].Dir); continue; }
                    byName[s.Name] = s;
                    result.Add(s);
                }
            }
            return result;
        }

        private static IEnumerable<Skill> LoadRoot(string root, string source)
        {
            var found = new List<Skill>();
            DirectoryInfo rootInfo;
            try
            {
                rootInfo = new DirectoryInfo(Path.GetFullPath(root));
                if (!rootInfo.Exists) return found;
                foreach (DirectoryInfo dir in rootInfo.GetDirectories().OrderBy(d => d.Name, StringComparer.Ordinal))
                {
                    try
                    {
                        if (IsReparsePoint(dir)) { DebugLog.Write("skills: " + dir.FullName + " skipped, it is a link"); continue; }
                        var md = new FileInfo(Path.Combine(dir.FullName, "SKILL.md"));
                        if (!md.Exists || IsReparsePoint(md) || md.Length > MaxFileBytes) continue;
                        string error;
                        Skill s = Parse(File.ReadAllText(md.FullName, Encoding.UTF8), dir.FullName, ScanFiles(dir), source, out error);
                        if (s == null) DebugLog.Write("skills: " + dir.FullName + " skipped, " + error);
                        else found.Add(s);
                    }
                    catch (Exception ex) { DebugLog.WriteException("skills: " + dir.FullName, ex); }
                }
            }
            catch (Exception ex) { DebugLog.WriteException("skills: " + root, ex); }
            return found;
        }

        // Every file under the skill folder except SKILL.md, as '/'-separated relative paths. Links (symlinks, junctions)
        // and dot-entries are never followed, so the list can only name files that physically sit inside the folder.
        private static List<string> ScanFiles(DirectoryInfo skillDir)
        {
            var files = new List<string>();
            var pending = new Stack<Tuple<DirectoryInfo, string>>();
            pending.Push(Tuple.Create(skillDir, ""));
            while (pending.Count > 0 && files.Count < MaxFiles)
            {
                var current = pending.Pop();
                foreach (FileSystemInfo entry in current.Item1.GetFileSystemInfos().OrderBy(e => e.Name, StringComparer.Ordinal))
                {
                    if (entry.Name.StartsWith(".", StringComparison.Ordinal) || IsReparsePoint(entry)) continue;
                    string rel = current.Item2 + entry.Name;
                    if (entry is DirectoryInfo sub) pending.Push(Tuple.Create(sub, rel + "/"));
                    else if (rel != "SKILL.md" && files.Count < MaxFiles) files.Add(rel);
                }
            }
            files.Sort(StringComparer.Ordinal);
            return files;
        }

        /// <summary>Parses SKILL.md. Returns null (with the reason) when the frontmatter has no valid name or description.</summary>
        public static Skill Parse(string text, string dir, List<string> files, string source, out string error)
        {
            error = null;
            string[] lines = text.TrimStart('﻿').Replace("\r\n", "\n").Split('\n');
            var meta = new Dictionary<string, string>(StringComparer.Ordinal);
            int bodyStart = 0;
            if (lines.Length > 0 && lines[0].Trim() == "---")
            {
                int end = Array.FindIndex(lines, 1, l => l.Trim() == "---");
                if (end < 0) { error = "the frontmatter has no closing ---"; return null; }
                ReadFrontmatter(lines, 1, end, meta);
                bodyStart = end + 1;
            }

            string name;
            if (!meta.TryGetValue("name", out name) || name == "") name = Path.GetFileName(dir.TrimEnd('\\', '/'));
            if (name.Length > 64 || !NamePattern.IsMatch(name)) { error = "the name '" + name + "' must be lowercase letters, digits and hyphens (at most 64)"; return null; }
            string description;
            if (!meta.TryGetValue("description", out description) || description == "") { error = "the frontmatter has no description"; return null; }
            if (description.Length > 1024) { error = "the description is longer than 1024 characters"; return null; }

            string hint;
            meta.TryGetValue("argument-hint", out hint);
            return new Skill
            {
                Name = name,
                Description = description,
                ArgumentHint = hint ?? "",
                Dir = dir,
                Body = string.Join("\n", lines, bodyStart, lines.Length - bodyStart).Trim(),
                Files = files ?? new List<string>(),
                ModelInvocable = !IsTrue(meta, "disable-model-invocation"),
                UserInvocable = !meta.ContainsKey("user-invocable") || IsTrue(meta, "user-invocable"),
                Source = source,
            };
        }

        // Flat "key: value" lines. A value of "", ">", "|" (or their "-" forms) takes the indented lines below it,
        // folded with spaces (">") or kept as lines ("|"). shortcut: no nested maps or lists; switch to a YAML library if
        // real skills need them.
        private static void ReadFrontmatter(string[] lines, int from, int to, Dictionary<string, string> meta)
        {
            for (int i = from; i < to; i++)
            {
                string line = lines[i];
                if (line.Length == 0 || char.IsWhiteSpace(line[0]) || line.TrimStart().StartsWith("#", StringComparison.Ordinal)) continue;
                int colon = line.IndexOf(':');
                if (colon <= 0) continue;
                string key = line.Substring(0, colon).Trim();
                string value = line.Substring(colon + 1).Trim();
                if (value == "" || value == ">" || value == "|" || value == ">-" || value == "|-")
                {
                    var block = new List<string>();
                    while (i + 1 < to && (lines[i + 1].Length == 0 || char.IsWhiteSpace(lines[i + 1][0]))) block.Add(lines[++i].Trim());
                    value = string.Join(value.StartsWith("|", StringComparison.Ordinal) ? "\n" : " ", block).Trim();
                }
                else if (value.Length >= 2 && (value[0] == '"' || value[0] == '\'') && value[value.Length - 1] == value[0])
                {
                    value = value.Substring(1, value.Length - 2);
                }
                meta[key] = value;
            }
        }

        private static bool IsTrue(Dictionary<string, string> meta, string key)
        {
            string v;
            return meta.TryGetValue(key, out v) && string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsReparsePoint(FileSystemInfo info)
        {
            return (info.Attributes & FileAttributes.ReparsePoint) != 0;
        }

        /// <summary>The catalog as the page reads it (window.__opendocsSkills).</summary>
        public static string ToPageJson(List<Skill> skills)
        {
            return JsonSerializer.Serialize(skills.Select(s => new
            {
                name = s.Name,
                description = s.Description,
                argumentHint = s.ArgumentHint,
                dir = s.Dir,
                body = s.Body,
                files = s.Files,
                modelInvocable = s.ModelInvocable,
                userInvocable = s.UserInvocable,
                source = s.Source,
            }));
        }

        /// <summary>read_skill_file {skill, path}: one file from the skill's scanned list, as text.</summary>
        public static ToolResult ReadSkillFile(List<Skill> skills, JsonElement input)
        {
            string skillName = Str(input, "skill");
            string path = Str(input, "path").Replace('\\', '/').Trim();
            Skill skill = skills.FirstOrDefault(s => s.Name == skillName);
            if (skill == null) return Error("There is no skill named '" + skillName + "'.");
            // Membership in the scanned list is the containment check: it never holds "..", a drive or a root.
            string rel = skill.Files.FirstOrDefault(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase));
            if (rel == null) return Error("'" + path + "' is not a file of the skill '" + skill.Name + "'. Its files are listed when the skill loads.");

            string full = Path.GetFullPath(Path.Combine(skill.Dir, rel.Replace('/', '\\')));
            string root = Path.GetFullPath(skill.Dir).TrimEnd('\\') + "\\";
            if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return Error("'" + path + "' is outside the skill folder.");
            // The folder may have changed since it was scanned: refuse a link anywhere between the file and the skill folder.
            for (FileSystemInfo at = new FileInfo(full); at.FullName.Length >= root.Length; at = ((at as FileInfo)?.Directory ?? ((DirectoryInfo)at).Parent))
            {
                if (!at.Exists) return Error("'" + path + "' no longer exists.");
                if (IsReparsePoint(at)) return Error("'" + path + "' goes through a link, which skills may not use.");
            }

            var file = new FileInfo(full);
            if (file.Length > MaxFileBytes) return Error("'" + path + "' is larger than " + (MaxFileBytes / 1024) + " KB. Its full path is " + full + ".");
            byte[] bytes = File.ReadAllBytes(full);
            if (Array.IndexOf(bytes, (byte)0) >= 0)
                return new ToolResult { Output = "'" + path + "' is a binary file. To use it with a tool that takes a file path (such as inserting an image), pass " + full, Summary = ReadToolName };
            return new ToolResult { Output = new UTF8Encoding(false).GetString(bytes).TrimStart('﻿'), Summary = ReadToolName };
        }

        private static string Str(JsonElement input, string key)
        {
            JsonElement v;
            return input.ValueKind == JsonValueKind.Object && input.TryGetProperty(key, out v) && v.ValueKind == JsonValueKind.String ? v.GetString() : "";
        }

        private static ToolResult Error(string message)
        {
            return new ToolResult { Output = message, IsError = true, Summary = ReadToolName };
        }
    }
}
