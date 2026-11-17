# build-msi.ps1 and OpenDocs.wxs

`build-msi.ps1` runs `package.ps1` (Release build of the four add-ins into `deploy\dist`), checks the policy file,
and builds `deploy\msi\OpenDocs-<version>.msi` from `OpenDocs.wxs` with WiX Toolset 5 (`wix build`). The MSI does
what `install.ps1` does for an all-users install, as one Windows Installer package that IT can push with Intune,
SCCM or Group Policy.

## What the MSI installs

- `C:\Program Files\OpenDocs\<App>AiAddIn\` for Word, Excel, PowerPoint and Outlook: the whole build output from
  `deploy\dist`, harvested with WiX's `Files` element, so a new file in the build needs no change here.
- `C:\Program Files\OpenDocs\policy.json`: the file passed as `-PolicyFile` (default `deploy\policy.json`).
- `HKLM\SOFTWARE\Microsoft\Office\<App>\Addins\<App>AiAddIn` with `Description`, `FriendlyName`, `LoadBehavior=3`
  and `Manifest`, written in both registry views. A 64-bit Office reads `SOFTWARE`, a 32-bit Office on 64-bit
  Windows reads `SOFTWARE\WOW6432Node`. `install.ps1` writes only the view of the installed Office when it can
  tell; the MSI cannot run that detection without a custom action, and the unused view is harmless.
- `Manifest` is `<folder>\<App>AiAddIn.vsto|vstolocal`, a plain path, the form Microsoft documents for VSTO add-ins
  deployed by Windows Installer. `install.ps1` writes a `file:///` URL instead; both load.

The package is per-machine and x64 (Program Files, not Program Files (x86)), the same folder `install.ps1` uses.

## Upgrades replace everything, policy.json included

`MajorUpgrade Schedule="afterInstallInitialize"` removes the installed version completely before installing the new
one. Windows Installer would otherwise keep an unversioned file that someone edited (it treats a file whose
modified date differs from its created date as user data), so an edited `policy.json` would survive an upgrade.
Removing first means the package's `policy.json` always wins, which is the intended behavior: the package is the
source of truth, as with `install.ps1`. `AllowSameVersionUpgrades` lets a rebuild with the same version replace the
installed one.

The UpgradeCode (`A51AEEE1-4241-4A44-8A75-4E09D9C3358F`) must never change, or a new MSI installs side by side
with the old one instead of replacing it.

## The policy check

Before building, the script loads `OfficeAi.Shared.dll` from the package and parses the policy with
`ModePolicy.Load`, the same code the add-ins run. A policy that switches every mode off in every app (what a
misspelled app or mode name, or invalid JSON, does) stops the build. Without this, the mistake would only show on
users' machines as a panel with no modes. The resolve handler stands in for the binding redirects in the add-in's
`.dll.config`, which Office applies but plain PowerShell does not: without it, `System.Text.Json` fails to load its
dependencies and every policy looks broken.

## Not handled by the MSI

- The prerequisites (.NET Framework 4.8, VSTO 2010 runtime, WebView2 runtime) are not checked or installed;
  `install.ps1` only warns about them too. Deploy them first.
- Outlook's per-user "do not disable this add-in" list (`Resiliency\DoNotDisableAddinList`) lives in each user's
  HKCU, which a per-machine package does not write, the same as an all-users `install.ps1`.
- A machine that was installed with `install.ps1` should be cleaned with `uninstall.ps1` first. The MSI does not
  own those files or registry keys, and Windows Installer may keep the old files instead of replacing them.

## Verified

Not yet built on this machine: WiX was not installed when this was written. To verify, install WiX 5
(`dotnet tool install --global wix --version 5.0.2`), run the script, install the MSI on a test machine, and
check the four add-ins load and `policy.json` is in place.
