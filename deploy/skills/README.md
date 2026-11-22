# OpenDocs skills

Skills installed here reach every user of the computer. Put each one in the folder of the app it is for:

```
skills\
  Word\<skill-name>\SKILL.md
  Excel\<skill-name>\SKILL.md
  PowerPoint\<skill-name>\SKILL.md
  Outlook\<skill-name>\SKILL.md
```

A skill is a folder in the Agent Skills format (the one Claude Code uses): a `SKILL.md`, plus any references,
templates or images it mentions. `SKILL.md` starts with frontmatter:

```markdown
---
name: weekly-report
description: Writes the weekly status report from the notes in the document. Use when asked for the weekly report.
argument-hint: [week number]
---

1. Read the whole document.
2. ...
```

- `name`: lowercase letters, digits and hyphens, at most 64 characters. Defaults to the folder name.
- `description`: required, at most 1024 characters. The assistant reads it to decide when the skill applies.
- `argument-hint`: optional, shown next to the name in the `/` list.
- `disable-model-invocation: true`: only the user can start it, by typing `/name`.
- `user-invocable: false`: hidden from the `/` list; only the assistant loads it.

`build-msi.ps1 -SkillsDir <folder>` builds a folder like this one into the MSI, and `install.ps1` copies the
package's `skills` folder. Users can add their own skills in `%LOCALAPPDATA%\<App>AiAddIn\Skills\`. A skill installed
here wins over a user's skill with the same name. Office reads skills when an app starts.
