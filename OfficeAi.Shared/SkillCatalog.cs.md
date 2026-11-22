# SkillCatalog.cs

Skills follow the Agent Skills format that Claude Code uses: a folder named after the skill, holding a `SKILL.md`
(YAML frontmatter, then markdown instructions) and any references, templates or assets the instructions point to.
Each Office app has its own skills, in two folders:

- `InstallRoot\skills\<App>\` (for example `C:\Program Files\OpenDocs\skills\Word\`), installed by IT next to
  `policy.json`. Found the same way `ModePolicy` finds the policy: one level up from the add-in's own folder.
- `%LOCALAPPDATA%\<App>AiAddIn\Skills\`, the user's own, next to their chat history and settings.

A machine skill wins over a user skill with the same name, so users cannot replace one IT installed.

The folders are read once per Office process (`For` caches them), like `policy.json`: a new or changed skill shows
after Office restarts. The page gets the whole catalog, bodies included, as `window.__opendocsSkills` before its
scripts run, so the `/` list, the system prompt and restoring a saved chat never wait on a message from the host.

## Frontmatter

There is no YAML library in the add-ins, and the keys a skill uses are flat: `name`, `description`,
`argument-hint`, `disable-model-invocation`, `user-invocable`. The parser reads `key: value` lines, strips one pair of
matching quotes, and takes the indented lines under a `>` or `|` value (folded with spaces or kept as lines). Nested
maps and lists are skipped, and other keys (`allowed-tools`, `model`, ...) are read and ignored. If real skills need
more, replace `ReadFrontmatter` with a YAML library.

A folder is skipped, with a line in `DebugLog`, when the name is not lowercase letters, digits and hyphens (at most
64), or the description is missing or over 1024 characters. A missing `name` falls back to the folder name, as in
Claude Code. A skipped folder never breaks the pane and cannot be read with `read_skill_file`.

## Why the model can only read listed files

`read_skill_file` takes a skill name and a relative path. The skill is looked up in this app's catalog, so the model
never chooses which folder is read. The path must match an entry of that skill's file list, which is built when the
folder is scanned. That list can only hold files that sit inside the folder:

- the scan skips links (any reparse point: symlinks and junctions) and dot-entries, so a junction in a user's skill
  folder pointing at their Documents never adds files;
- the list never holds `..`, a drive or a root, so membership alone refuses traversal and absolute paths.

The folder can change after the scan (a user replaces `refs\` with a junction). So before reading, the full path is
checked to start with the skill folder, and the file and every folder between it and the skill folder are checked
again for a reparse point. `Path.GetFullPath` alone would not catch this: it normalizes `..` but does not resolve
links.

Files over 256 KB are refused. A file with a NUL byte is treated as binary: instead of its bytes, the model gets its
full path, for tools that take a local file path (inserting an image). `SKILL.md` itself is not in the list; the
model gets its body from `load_skill`.

## Why read_skill_file is answered in PaneHostBase

It is read-only and identical in every app, so `PaneHostBase.ExecuteToolUnderPolicy` answers it after the policy
check and before the app's own `ExecuteTool`. It never reaches the four apps' tier gates, and it is in no app's
`AlwaysAllowedTools` and no `entry.ts` tier list: the page treats both skill tools as read-only itself
(`bootstrap.ts`). That is the same exception as Outlook's `preview_action`.
