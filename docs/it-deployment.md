# OpenDocs: deployment guide for IT

This guide is for the people who install OpenDocs on company computers and decide what the AI assistant may do. It
covers installing with the MSI, the administrator policy (`policy.json`), upgrades and removal.

## What gets installed

One MSI installs the OpenDocs add-in for **Word, Excel, PowerPoint and Outlook**, for every user of the computer.
There are two builds of it; deploy the one that matches **Windows** (not Office):

- `OpenDocs-<version>-x64.msi` for 64-bit Windows, with either 32-bit or 64-bit Office;
- `OpenDocs-<version>-x86.msi` for 32-bit Windows. It refuses to install on 64-bit Windows.

It installs:

| What | Where |
|---|---|
| The four add-ins | `C:\Program Files\OpenDocs\WordAiAddIn\`, `ExcelAiAddIn\`, `PowerPointAiAddIn\`, `OutlookAiAddIn\` |
| The administrator policy | `C:\Program Files\OpenDocs\policy.json` |
| Skills for every user (optional) | `C:\Program Files\OpenDocs\skills\<App>\<skill>\` |
| Office add-in registration | `HKLM\SOFTWARE\Microsoft\Office\<App>\Addins\<App>AiAddIn`; on 64-bit Windows in both the 64-bit and the 32-bit registry view, so it works with either Office |

Only administrators can change files under Program Files, so users cannot change the add-in or the policy. Each
user's own settings (which AI provider and model to use, the API key, theme, language) are created the first time
they open the panel, in their own profile. API keys are stored encrypted for that Windows user.

OpenDocs does not need internet access to install or run. It talks only to the AI provider each user configures,
which can be a server on your own network.

## Before you install

OpenDocs needs three Microsoft runtimes. `OpenDocs.msi` does not install them, but **it checks for them and refuses
to install when one is missing**, with a message naming it (in the install log, `msiexec` exit code 1603). That way
it never installs add-ins that Office then cannot load. Deploy what is missing first, then OpenDocs:

| Prerequisite | Usually already there? | How to deploy it as an MSI |
|---|---|---|
| **.NET Framework 4.8** or later | Yes: it is part of Windows 10 version 1903 and later and of Windows 11. | Nothing to deploy on those. Older Windows (for example Windows 10 LTSC 2019, which has 4.7.2) gets it through Windows Update / WSUS. Microsoft publishes no MSI for it. |
| **Visual Studio 2010 Tools for Office Runtime** (VSTO runtime) | Often, through Office; check with an inventory. | Microsoft's own MSI, taken out of `vstor_redist.exe` and delivered next to the OpenDocs MSIs: `prerequisites\VSTO-Runtime\x64\vstor40_x64.msi` (64-bit Windows; `x86\vstor40_x86.msi` for 32-bit Windows). Keep each `.msi` next to its `.cab`. Install it silently with `msiexec /i vstor40_x64.msi /qn` (or `vstor40_x86.msi`). |
| **Microsoft Edge WebView2 Runtime** | Yes on Windows 11; on Windows 10, Microsoft 365 Apps installs it. | Microsoft publishes it only as an .exe (`MicrosoftEdgeWebView2RuntimeInstallerX64.exe /silent /install`, or the X86 one on 32-bit Windows). If computers without it remain, it needs a wrapper MSI. |

To check a computer (all three must be present):

| Prerequisite | Registry value OpenDocs checks |
|---|---|
| .NET 4.8+ | `HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full`, `Release` ≥ 528040 |
| VSTO runtime | `HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R`, `Version` (in either the 64-bit or the 32-bit registry view) |
| WebView2 | `HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}`, `pv` present and not `0.0.0.0` |

Microsoft Office desktop apps (Word, Excel, PowerPoint, Outlook), 32- or 64-bit, are needed too, of course.

With Configuration Manager, make the VSTO runtime (and WebView2, if you package it) a dependency of OpenDocs. With
Group Policy software installation, put them in an earlier GPO or above OpenDocs in the same GPO. If OpenDocs runs
first anyway, it fails cleanly and installs on the next attempt once the prerequisite is there.

If a computer had OpenDocs installed earlier with the `install.ps1` script, run `uninstall.ps1` from that package
first (elevated). The MSI does not own the files and registry keys the script created and may leave old files in place.

## Install

Close Word, Excel, PowerPoint and Outlook on the target computer first: Office keeps the add-in files locked while
it runs. Then, from an elevated prompt:

```
msiexec /i OpenDocs-<version>-x64.msi /qn /l*v C:\Windows\Temp\OpenDocs-install.log
```

- `/qn` installs silently. Leave it out to see the installer's progress window.
- With Intune (Win32 app or line-of-business app), Configuration Manager or Group Policy software installation,
  deploy the MSI as a per-machine (device) install. It needs no command-line properties.
- Users get the add-ins the next time they start each Office app.

To check one computer: start Excel, open the OpenDocs panel, and open the editing-mode menu next to the message box.
It should show the four modes, with any your policy switched off grayed out.

## Upgrade

Install the newer MSI the same way. It removes the old version first and then installs the new one, so **every file
is replaced, including `policy.json`**. If someone edited the installed `policy.json` by hand, that edit is lost on
upgrade: keep your policy in the MSI (see below) or deploy it again after every upgrade.

## Uninstall

```
msiexec /x OpenDocs-<version>-x64.msi /qn
```

Or remove **OpenDocs** in Settings → Apps. This removes the add-ins, their registration and `policy.json`. Each
user's own settings and chat history stay in their profile.

## Skills

A skill is a set of written instructions for one kind of task, such as "write the weekly report the way we do it",
with any reference files or images it needs. Users start one by typing `/` in the panel and picking it, and the
assistant loads one by itself when a request matches its description. Skills use the Agent Skills format (the one
Claude Code uses); `deploy\skills\README.md` shows the layout and the `SKILL.md` header.

- **Skills for everyone:** put them in a folder with one subfolder per app (`Word`, `Excel`, `PowerPoint`,
  `Outlook`) and build the MSI with `-SkillsDir` (see below). They are installed under Program Files, where users
  cannot change them, and every upgrade replaces them with the package's copy.
- **A user's own skills:** `%LOCALAPPDATA%\<App>AiAddIn\Skills\` (for example `WordAiAddIn\Skills\`). A skill IT
  installed wins over a user's skill of the same name.
- Office reads skills when an app starts: restart it after adding or changing one.
- **What a skill can do:** only what the user's editing mode, the policy and the Approve card already allow. A skill
  adds no tools and no permissions, but its text does steer the assistant, so a user's own skill is as trusted as
  what that user types. The assistant can read a skill's own files and nothing else through it. If the assistant
  must never act on instructions without a person checking each step, switch Full autonomy off with
  `disabledModes` (in Outlook, Full autonomy is what sends mail).

## The policy: what the assistant may do

`policy.json` controls two things, separately for each app:

1. **`disabledModes`**: editing modes the user cannot pick. They stay visible in the menu, grayed out.
2. **`requireApproval`**: actions the user must approve, on an Approve/Deny card, before the assistant carries them out.

### The editing modes

The user picks a mode in the panel. Each mode allows more than the one before it. In `policy.json` the modes always
have the same four names, but each app shows its own label:

| Name in policy.json | Word | Excel and PowerPoint | Outlook |
|---|---|---|---|
| `readOnly` | Read only: reads the document | Read only | Read only: reads and searches mail and calendar |
| `commentOnly` | Comment only: adds comments, never changes text | Comment only (these apps have no comment tool yet, so this is the same as Read only) | Draft only: triage, tasks, and drafts the user reviews and sends; nothing is sent |
| `trackChanges` | Track changes: edits as Word tracked changes | **Edits with approval**: any edit, asking for approval as the policy sets | **Actions with approval**: also answers meeting invitations and deletes permanently, asking for approval as the policy sets |
| `fullAutonomy` | Full autonomy: edits directly | Full autonomy | Full autonomy: also sends mail and calendar events immediately |

**Full autonomy never asks for approval**, whatever the policy lists: it is the mode without a person checking each
step. If the assistant must not act on its own, switch Full autonomy off with `disabledModes`.

The mode is enforced twice: the panel only offers the assistant the tools the mode allows, and the add-in itself
refuses any other call. A disabled mode is refused by the add-in even if someone tampered with the panel.

### The file

`policy.json` is JSON. Comments (`// ...`) and trailing commas are allowed. The shipped file:

```jsonc
{
  "disabledModes": {
    "Word": [],
    "Excel": [],
    "PowerPoint": [],
    "Outlook": []
  },
  "requireApproval": {
    "Word": {},
    "Excel": { "*": ["trackChanges"] },
    "PowerPoint": { "*": ["trackChanges"] },
    "Outlook": { "*": ["trackChanges"], "delete_email": ["commentOnly", "trackChanges"] }
  }
}
```

With these defaults:
- every mode is available;
- Excel, PowerPoint and Outlook ask before every action in their third mode (Edits / Actions with approval);
- Outlook also asks before deleting a message in Draft only;
- Word never asks, because its Track changes mode already leaves every edit for the user to accept or reject.

**`disabledModes`**: for each app, the list of modes to switch off.

```jsonc
"disabledModes": { "Outlook": ["fullAutonomy"] }   // Outlook can never send mail on its own
```

**`requireApproval`**: for each app, a map from a tool name to the modes in which the user must approve it.
- `"*"` means every tool that changes something; tools that only read never ask.
- A tool's own entry overrides `"*"`, so `"save": []` means "never ask before save" even when `"*"` asks.
- An app or tool left out never asks.

```jsonc
"requireApproval": {
  "Excel": { "*": ["trackChanges"], "save": [] },               // ask before every change except saving
  "Outlook": { "respond_meeting": ["trackChanges"] }            // ask before answering a meeting invitation
}
```

The tool names for each app are listed in the project's tool reference: `docs/tools/word.md`, `excel.md`,
`powerpoint.md` and `outlook.md`.

When an approval is required, the panel shows a card with what is about to happen (for Outlook, the message's type,
subject, sender and date; otherwise the action and its details) and waits for **Approve** or **Deny**. Deny means
the action does not run, and the assistant is told not to retry it.

### Mistakes in the file

| Mistake | Effect |
|---|---|
| Invalid JSON, a misspelled **app** name or a misspelled **mode** name | **Every mode is switched off in every app.** The panel says the administrator turned the modes off. This is deliberate: a typo must never leave a mode open that you meant to close. |
| A **tool** name the app does not have | Ignored. Nothing else changes and the user sees nothing, so check tool names against the tool reference. |
| The file is missing | No restrictions: every mode, no approvals. |

`build-msi.ps1` checks the file before building the MSI and stops if it would switch every mode off.

### Changing the policy

The add-ins read `policy.json` when an Office app starts, so a change takes effect when users restart Office.

- **Recommended: build the policy into the MSI.** Ask whoever builds OpenDocs to run
  `deploy\build-msi.ps1 -PolicyFile <your policy.json>`, or run it yourself on a build machine (see below). Then
  every install and upgrade gets your policy.
- **Or replace the file after installing**, for example with a Group Policy Preferences file item (action:
  Replace) or an Intune script that copies your `policy.json` to `C:\Program Files\OpenDocs\policy.json`. Remember
  that an upgrade puts the MSI's own copy back, so the copy has to run again afterwards.

## Building the MSI

On a machine with the OpenDocs source, Visual Studio 2022 (with the Office/SharePoint workload), Node.js, and the
add-ins' signing certificate:

```
dotnet tool install --global wix --version 5.0.2
cd deploy
.\build-msi.ps1 -PolicyFile C:\it\policy.json -SkillsDir C:\it\skills
```

- The version comes from the newest `vX.Y.Z` git tag, or pass `-Version X.Y.Z`.
- The MSIs are written to `deploy\msi\OpenDocs-<version>-x64.msi` and `OpenDocs-<version>-x86.msi`.
- The script prints, per app, which modes are off and what asks for approval, so you can check the policy before
  you deploy it.
- Without `-SkillsDir` the MSI installs no skills (`deploy\skills` holds only its README). The script lists each
  app's skills and stops if a folder has a `SKILL.md` the add-in would skip (a bad name, no description, or the
  same name twice).

## Security notes

- **What a mode switch or disabled mode guarantees:** the add-in refuses tool calls outside the user's mode, and
  any call while the panel is in a disabled mode.
- **How approvals are enforced:** by the panel. To make the panel hard to tamper with, the installed add-ins switch
  off the browser developer tools (F12) in the panel.
- **Calls the user types with `!`** (an opt-in developer feature in the panel's settings) never ask for approval:
  the user wrote the call themselves. The editing mode still limits them.
- **Data and the network:** OpenDocs has no server of its own and no built-in internet features. It sends document
  content only to the AI provider the user configured.

## Troubleshooting

- **The panel does not appear.** In the Office app: File → Options → Add-ins → Manage: COM Add-ins → Go. Check the
  add-in is ticked, and look under Manage: Disabled Items. Outlook disables add-ins it thinks start slowly; re-enable
  it there.
- **The panel says every mode is turned off.** `policy.json` has a typo (see "Mistakes in the file"). Each user's
  `%TEMP%\OpenDocsDebug.log` has a line starting `ModePolicy:` if the file could not be read.
- **A mode is grayed out.** That mode is in `disabledModes` for that app.
- **The add-in fails to load with a runtime error.** Check the prerequisites above, then look in Event Viewer →
  Windows Logs → Application for a `.NET Runtime` or `VSTO 4.0` error at the time Office started.
