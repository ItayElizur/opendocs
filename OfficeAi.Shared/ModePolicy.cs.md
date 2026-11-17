# ModePolicy.cs

`policy.json` is the administrator's file. It sits in the install folder next to the four per-app folders, and an
all-users install puts it in Program Files, where only administrators can change it. It has two per-app sections.

## disabledModes

`disabledModes` switches whole editing modes off. This is enforced on both sides:

- **C#:** `PaneHostBase` refuses every tool call while the pane is in a disabled mode, or before the page has reported
  a mode (`Blocks`). It never passes a disabled mode on to the app's `SetMode`. Before every call it stores the pane's
  mode again under its current chat id.
- **Page:** the page leaves disabled modes out of the menu.

## requireApproval

`requireApproval` maps a tool name to the modes in which the user must click Approve before the model's call runs.

- `"*"` stands for every tool that is not read-only in that app. A tool's own entry overrides `"*"`, so
  `"add_comment": []` exempts one tool.
- The page shows the card (`bootstrap.ts`, `executeConfirmed`). How the card looks comes from the app's `confirmActions`
  rules, or from a plain "Allow …?" card for a tool that has no rule.
- **This is enforced by the page only.** A C# check cannot help, because the page is also what shows the card and
  reports the click, so a page that has been tampered with could report an approval it never asked for. To make that
  harder, WebView2's developer tools are switched off in any app whose policy has a rule (`WebViewBridgeHost`).
- Calls the user types with `!` are never asked: the user wrote the call themselves.

## Fail closed

Any error in the file switches every mode off in every app, so a typo never leaves a mode open or an approval
missing. That covers invalid JSON, a wrong shape, an unknown app or mode name, and a tool name that is neither `"*"`
nor a lower-case identifier. C# does not know each app's tool names, so the page checks them. When the page finds a
tool name it does not know, it requires an Approve for every non-read-only tool in every mode and shows an error that
names the tool.

## Changes

The file is read once per Office process, so changes take effect when Office is restarted.
