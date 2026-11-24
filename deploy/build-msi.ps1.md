# build-msi.ps1 and OpenDocs.wxs

`build-msi.ps1` runs `package.ps1` (Release build of the four add-ins into `deploy\dist`), checks the policy file,
and builds two MSIs from `OpenDocs.wxs` with WiX Toolset 5 (`wix build`): `deploy\msi\OpenDocs-<version>-x64.msi` for
64-bit Windows and `OpenDocs-<version>-x86.msi` for 32-bit Windows. Each MSI does what `install.ps1` does, as one
Windows Installer package: started elevated (as IT pushes it with Intune, SCCM or Group Policy) it installs for all
users, and started without admin rights it installs for the current user only.

## One package, two scopes

`Scope="perUserOrMachine"` makes it a dual-purpose package: `ALLUSERS=2` with `MSIINSTALLPERUSER=1`, so a plain
install is per-user and needs no elevation. A first install that is running elevated sets `ALLUSERS=1` before
`FindRelatedProducts`, which makes it per-machine. `MSIUSEREALADMINDETECTION=1` makes `AdminUser` mean "this process
is really elevated"; without it, Windows reports any member of the Administrators group as an admin, even
un-elevated. A deployment tool running the MSI as SYSTEM is elevated, so it installs for all users. Repair and
uninstall are left alone (`NOT Installed`): Windows sets `ALLUSERS` from the scope the product was installed with.

Windows Installer picks the per-user folders (`ProgramFiles64Folder` and friends point at
`%LOCALAPPDATA%\Programs\`) from the starting `ALLUSERS`/`MSIINSTALLPERUSER`, before any custom action can change
them. A probe package that only flipped `ALLUSERS` registered under HKLM but put its files in `%LOCALAPPDATA%`. So
an all-users install also sets `INSTALLFOLDER` to `[%ProgramW6432]\OpenDocs` (`[%ProgramFiles]` on 32-bit Windows),
which is what WiX's own install-scope dialog does. The registry entries use `HKMU`, which is HKLM for all users and
HKCU for one.

## What the MSI installs

- `C:\Program Files\OpenDocs\<App>AiAddIn\` for Word, Excel, PowerPoint and Outlook (per-user:
  `%LOCALAPPDATA%\Programs\OpenDocs\<App>AiAddIn\`): the whole build output from `deploy\dist`, harvested with WiX's
  `Files` element, so a new file in the build needs no change here.
- `policy.json` next to those folders: the file passed as `-PolicyFile` (default `deploy\policy.json`).
- `HKLM\SOFTWARE\Microsoft\Office\<App>\Addins\<App>AiAddIn` (per-user: the same key under HKCU) with
  `Description`, `FriendlyName`, `LoadBehavior=3` and `Manifest`. For all users, the x64 MSI writes both registry
  views: a 64-bit Office reads `SOFTWARE`, a 32-bit Office on 64-bit Windows reads `SOFTWARE\WOW6432Node`. The x86
  MSI writes the one view 32-bit Windows has (only 32-bit Office runs there). `install.ps1` writes only the view of
  the installed Office when it can tell; the MSI cannot run that detection without a custom action, and the unused
  view is harmless. `HKCU\Software` has no separate 32-bit view, so for one user the x64 MSI's two components write
  the same key.
- `Manifest` is `<folder>\<App>AiAddIn.vsto|vstolocal`, a plain path, the form Microsoft documents for VSTO add-ins
  deployed by Windows Installer. `install.ps1` writes a `file:///` URL instead; both load.

For all users, both packages install into Program Files (`ProgramFiles6432Folder`: the 64-bit one, not Program
Files (x86), on 64-bit Windows), the same folder `install.ps1` uses. The add-ins are AnyCPU and the package carries
the x86, x64 and arm64 WebView2 loaders, so the two MSIs differ only in folder, registry views and one extra check:
the x86 MSI refuses to install on 64-bit Windows (`NOT VersionNT64`), where it would register the add-ins only for
32-bit Office. Both share one `UpgradeCode`; a computer only ever gets one of them.

## Upgrades replace everything, policy.json included

`MajorUpgrade Schedule="afterInstallInitialize"` removes the installed version completely before installing the new
one. Windows Installer would otherwise keep an unversioned file that someone edited (it treats a file whose
modified date differs from its created date as user data), so an edited `policy.json` would survive an upgrade.
Removing first means the package's `policy.json` always wins, which is the intended behavior: the package is the
source of truth, as with `install.ps1`. `AllowSameVersionUpgrades` lets a rebuild with the same version replace the
installed one.

The UpgradeCode (`A51AEEE1-4241-4A44-8A75-4E09D9C3358F`) must never change, or a new MSI installs side by side
with the old one instead of replacing it. The per-machine packages built before the per-user option share it, so
the dual-purpose package upgrades them when run elevated.

## Never two installs at once

`FindRelatedProducts` only sees an older version installed in the same scope, and knows nothing about `install.ps1`.
So the package looks for each app's add-in registration (`Manifest` under HKCU, and under HKLM in both views) and
refuses, with a message saying what to remove, when the install would leave two copies:

- **For all users while the current user has their own copy** (HKCU): Office prefers a user's own registration, so
  that user would keep the old copy.
- **For one user while an all-users copy exists** (HKLM): the all-users copy already serves the user. More
  importantly, both installs use this package's component IDs, so Windows counts each as a client of the other's
  components. In a test, removing a per-user copy while the all-users one stayed logged "Disallowing uninstallation
  of component ... since another client exists" and left the HKCU registrations behind, pointing at deleted files.
- **Over an `install.ps1` install of the same scope** (a registration without `WIX_UPGRADE_DETECTED`): the package
  would not reliably replace its files (see below). A Visual Studio build of a VSTO project also registers itself
  under HKCU, so on a dev machine the per-user check fires until the registrations are cleaned.

The checks read the HKCU of the account running the installer. Another user's own copy is not seen; under SYSTEM
(a deployment tool), HKCU is SYSTEM's own, which has none.

## The policy check

Before building, the script loads `OfficeAi.Shared.dll` from the package and parses the policy with
`ModePolicy.Load`, the same code the add-ins run. A policy that switches every mode off in every app (what a
misspelled app or mode name, or invalid JSON, does) stops the build. Without this, the mistake would only show on
users' machines as a panel with no modes. The resolve handler stands in for the binding redirects in the add-in's
`.dll.config`, which Office applies but plain PowerShell does not: without it, `System.Text.Json` fails to load its
dependencies and every policy looks broken.

## Not handled by the MSI

- The prerequisites (.NET Framework 4.8, VSTO 2010 runtime, WebView2 runtime) are not installed: an MSI cannot run
  another installer while it installs. They are checked instead (registry searches, the same values `install.ps1`
  looks at). Each missing one sets a `MISSING_*` property to a line of text, and one `Launch` condition per scope
  shows them all in one message. `Installed OR` keeps uninstall and repair working on a computer that has since lost
  a prerequisite. The VSTO runtime looks in both registry views because it is registered in either, depending on
  what installed it (only the 32-bit view on the dev machine). .NET and the VSTO runtime only install machine-wide.
  WebView2 can also be installed for one user (HKCU), which counts for a per-user install. An all-users install needs
  the machine-wide runtime, and says so when it finds only the user's: a per-user runtime does not reach other
  users, and under SYSTEM its HKCU says nothing about anyone.
- Outlook's per-user "do not disable this add-in" list (`Resiliency\DoNotDisableAddinList`) lives in each user's
  HKCU, which the package does not write, the same as `install.ps1`.
- An `install.ps1` install has to be removed with `uninstall.ps1` first, and the package refuses until it is (see
  "Never two installs at once"). The MSI does not own those files or registry keys. On a first install, Windows
  Installer keeps an existing file of the same version (the add-ins' assembly versions do not change between
  builds) and an unversioned file it thinks a user edited, so old files would survive.

## The VSTO runtime MSI

Microsoft ships the VSTO runtime only as `vstor_redist.exe`, but inside it is an ordinary Windows Installer package
that installs on its own (no launch conditions; it finds .NET 4 through its own registry search). It is not in the
repo. To take it out again with 7-Zip, without running anything:

```
7z e vstor_redist.exe -oouter "*vstor40_x64.exe" "*vstor40_x86.exe" -r
7z e outer\vstor40_x64.exe -ox64 vstor40_x64.msi vstor40_x64.cab
7z e outer\vstor40_x86.exe -ox86 vstor40_x86.msi vstor40_x86.cab
```

Each `.msi` needs its `.cab` next to it. Both are signed by Microsoft (`Get-AuthenticodeSignature`). The copy
delivered with OpenDocs 0.2.0 came from `vstor_redist.exe` 10.0.60910 and is version 10.0.60915.

## Verified

Built with WiX 5.0.2. The x64 MSI holds every file of the package (checked against `deploy\dist` after an
administrative extract), registers both registry views, and schedules `RemoveExistingProducts` right after
`InstallInitialize`. Installed on the dev machine (Windows 11, 64-bit Office), Excel loaded the add-in from it.
The prerequisite checks were tested with throwaway builds whose registry searches point at missing keys: each one
stops the install at `LaunchConditions` with its own message. Not tested: a computer that really lacks a
prerequisite, 32-bit Office, and deployment through Intune, Configuration Manager or Group Policy.

The x86 MSI was only built and decompiled (`wix msi decompile`): it installs into `ProgramFilesFolder`, has only the
32-bit registry components, and carries the `NOT VersionNT64` check. It was never installed - there is no 32-bit
Windows here to install it on.

The two scopes (x64 MSI, Windows 11, `msiexec /qn` with a verbose log, several builds of one version so each has its
own ProductCode):

- **Without admin:** files in `%LOCALAPPDATA%\Programs\OpenDocs`, HKCU registrations, product registered per-user.
  A newer per-user build upgraded it, and removing it left nothing behind.
- **Elevated:** files in `C:\Program Files\OpenDocs`, both HKLM views, product registered per-machine. It upgraded
  the earlier per-machine-only package (0.3.1) and a dual-purpose all-users build. A non-elevated removal is refused
  with error 1730, as for any per-machine product; an elevated removal left nothing behind.
- **Refused, each with its own message:** an elevated install while a per-user copy exists; a per-user install
  while an all-users copy exists; a per-user install over a Visual Studio build's HKCU registrations.
- **A rollback:** a 0.3.0 build run elevated over the installed 0.3.1 showed only WiX's "A newer version of OpenDocs
  is already installed." A newer version installed sets `WIX_DOWNGRADE_DETECTED`, not `WIX_UPGRADE_DETECTED`, so
  without `OR WIX_DOWNGRADE_DETECTED` the `install.ps1` check would fail too, and which of the two messages Windows
  shows is not defined.
- **The prerequisite message,** with throwaway builds whose VSTO and WebView2 searches point at missing keys: per-user
  and all-users each listed the missing items, one per line. With a WebView2 key added only under HKCU, the
  all-users message said it is installed only for this account.

Not tested: Office loading the add-ins from a per-user MSI install (`install.ps1 -CurrentUser`, which registers the
same way, was), the dialogs a double-click shows, and the x86 MSI in either scope.
