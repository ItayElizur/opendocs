# install.ps1 / uninstall.ps1

## Scope: machine-wide by default, per user on request

`install.ps1` used to be per-user only (files in `%LOCALAPPDATA%\OpenDocs`, registered under `HKCU`, certificate in
`CurrentUser\TrustedPublisher`). That left every user of a shared computer to run the installer themselves, and on one
machine the certificate import was denied ("access denied" into `CurrentUser\TrustedPublisher`, probably a Group Policy or
permission restriction). The default is now a machine-wide install from an elevated PowerShell; `-CurrentUser` keeps the old
per-user behaviour (no admin). `uninstall.ps1` mirrors the switches.

| | machine-wide (default) | `-CurrentUser` |
|---|---|---|
| files | `%ProgramW6432%\OpenDocs\<App>AiAddIn` | `%LOCALAPPDATA%\OpenDocs\<App>AiAddIn` |
| registration | `HKLM\SOFTWARE\Microsoft\Office\<App>\Addins\<App>AiAddIn` (and the `WOW6432Node` view, see below) | `HKCU\Software\Microsoft\Office\<App>\Addins\<App>AiAddIn` |
| certificate | `Cert:\LocalMachine\TrustedPublisher` | `Cert:\CurrentUser\TrustedPublisher` |
| Outlook slow-add-in whitelist | not set (it is a per-user `HKCU` list an elevated run cannot reach) | set, as before |

The registered values are the same in both modes: `Description`, `FriendlyName`, `LoadBehavior` = 3, and `Manifest` =
`file:///<path>/<App>AiAddIn.vsto|vstolocal`.

### Why Program Files and not ProgramData

VSTO loads the add-in's managed assemblies inside every user's Office process. A folder that standard users can write to
would let one user replace those DLLs and run code in the others' sessions. `Program Files` is writable by administrators
only and readable and executable by everyone, which is what is needed. `ProgramW6432` is used so a 32-bit PowerShell still
lands in the 64-bit `Program Files`.

### Bitness

A 32-bit Office on 64-bit Windows reads add-in registrations from `HKLM\SOFTWARE\WOW6432Node`, a 64-bit Office from
`HKLM\SOFTWARE`. `Get-OfficeBitness` reads `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\Platform` (Click-to-Run),
then Outlook's `Bitness` value for MSI installs, and the matching view is written. When it cannot tell, both views are written:
the one Office does not read is harmless. `HKCU` has no such split, so per-user is unchanged.

### The certificate

- It is only imported when it is not already trusted. A machine-wide trust also covers a per-user install; the reverse is not
  true, because `CurrentUser` belongs to whichever account ran the script (an elevated shell run as a different administrator
  puts the certificate in that administrator's store, not the user's).
- A failed import is a WARNING, never a stop. It names the likely cause (a Group Policy or permission restriction on Trusted
  Publishers; `gpresult /h gp.html` shows which) and the options (the elevated machine-wide install, or having IT deploy
  `OpenDocsDevCert.cer` through Group Policy), then the install continues. It used to throw, which on the machine that first
  showed "access denied" turned a certificate problem into a failed install, although the add-in loads without the certificate
  (see "What the certificate is for" below). The `E_ACCESSDENIED` path was reproduced by running the certificate block
  non-elevated against the machine store: it warned and continued.
- It is NOT added to `LocalMachine\Root`. The per-user install has always worked with `TrustedPublisher` alone, and trusting a
  self-signed certificate as a root for the whole computer is a bigger security decision than this installer should take. If
  a machine turns out to need it, that should be an explicit, separate step.

### Other behaviour

- `-WhatIf` (the script is `SupportsShouldProcess`) prints every file copy, registry write and certificate import and changes
  nothing; it also works without elevation, so a machine-wide install can be reviewed first. Registry values are described in
  a line of their own because `New-Item -WhatIf` does not create the key.
- The script stops with a message when an Office app being installed is running (any user's session for a machine-wide
  install, this user's otherwise), instead of failing part way through replacing locked files. Under `-WhatIf` it only warns.
- If this account also has a per-user registration of the same add-in, a machine-wide install says so: two registrations would
  load the add-in twice. Run `uninstall.ps1 -CurrentUser` first.

## Verification status

NOT VERIFIED on a second user account, and the machine-wide write path was never run for real, because the tool-driven
session that wrote it cannot elevate. What was checked, on one Windows 11 machine (Office x64):

- both scripts parse with the PowerShell parser (0 errors) and are ASCII with CRLF, as before;
- a non-elevated run without `-CurrentUser` stops with the "needs an elevated PowerShell" message and changes nothing;
- `-WhatIf` for the machine-wide install lists the certificate import, the `Program Files` copy and the `HKLM` registration with
  the right manifest URI, and Program Files, the `HKLM` key and the machine certificate store were confirmed untouched;
- `-CurrentUser -WhatIf` recognises the already-trusted certificate and skips the import;
- the "Office is running" guard stops a real run while Word is open;
- Office bitness is detected as x64 here;
- real `-CurrentUser` runs (Word closed): `install.ps1 -CurrentUser -App Word` skipped the already-trusted certificate, copied the
  files, and re-pointed the `HKCU` registration at the installed copy (a VSTO build re-registers the add-in at the repo's
  `bin\Release` on every MSBuild run, so before the install Word's `Manifest` pointed there); a hidden Word then listed the
  add-in with `Connect = True`. `uninstall.ps1 -CurrentUser -App Word` removed the key and the folder and left the other
  apps' folders alone, and a reinstall restored the working state.

Machine-wide install, run for real by the repository owner from an elevated PowerShell (Windows 11, Office x64, one user
account on the machine): `uninstall.ps1 -CurrentUser -App Word`, then `install.ps1 -App Word` completed without errors. Afterwards
the files were in `C:\Program Files\OpenDocs\WordAiAddIn`, `HKLM\SOFTWARE\Microsoft\Office\Word\Addins\WordAiAddIn` had
`LoadBehavior` 3 and the Program Files manifest (the `WOW6432Node` view was correctly not written for x64 Office), no per-user key
was left, `CN=WordAiAddIn Temporary Key` was in `LocalMachine\TrustedPublisher`, and Word started and worked with the add-in. A
`Get-ChildItem Cert:\...` check pasted after a formatted `Get-ItemProperty ... | Select` printed nothing in that session; that
is PowerShell's display of a second object type after a formatted table, not a missing certificate (a separate read showed it).
That also settles item 3 below: Office resolved the manifest path with a space.

Caveat that mattered: that account ALSO had the certificate in its own `CurrentUser\TrustedPublisher`, so that run alone did not
show whether the `LocalMachine` copy suffices. The later runs under "What the certificate is for" answer a different question:
whether the certificate is needed at all (no, in the configurations tried).

## What the certificate is for (measured)

Run on one Windows 11 machine, Office x64, Word, with the add-in registered machine-wide (`|vstolocal` manifest in Program Files),
always from a fresh Word start:

1. Certificate in neither `CurrentUser\TrustedPublisher` nor `LocalMachine\TrustedPublisher`: the add-in loads, with no security
   prompt and nothing logged.
2. The same, with this account's per-user VSTO trust data (`HKCU\Software\Microsoft\VSTO\Security\Inclusion` and
   `SolutionMetadata`) backed up and cleared, which makes the account look like a brand-new user: still loads, no prompt; VSTO
   recorded only a `SolutionMetadata` entry and wrote nothing to the inclusion list (which is what an accepted trust prompt writes).
3. The same, with Word's Trust Center option "Require that application add-ins are signed by Trusted Publisher" on
   (`HKCU\Software\Microsoft\Office\16.0\Word\Security\RequireAddinSig` = 1): Word showed a yellow "Some active content has
   been disabled" bar listing two COM add-ins, but ours was not among the disabled ones and still loaded. The two that were
   disabled are almost certainly the other add-ins registered for Word on that machine (Adobe `PDFMaker.OfficeAddin` and the
   Visual Studio design-time adaptor `VS15WordAdaptor`).

So, in the default configuration and with the "require signed add-ins" option on, the certificate was not needed. The test values
were removed and the trust data restored from the backup afterwards. What this does NOT cover: a machine whose Group Policy
enforces publisher trust some other way (for instance the VSTO runtime's own policies), which is the situation the original
"access denied" machine may be in; there the certificate has to be trusted, by an elevated machine-wide install or by deploying
`OpenDocsDevCert.cer` through Group Policy. The installer still imports the certificate when it can, because on such machines it
is what makes the add-in load.

Still to test:

1. As a standard user, and as a user created after the install (neither has the certificate in its own store), start Word: does
   the add-in load? Quicker on one account: delete this account's own `CurrentUser\TrustedPublisher` copy and start Word.
2. If item 1 fails, the machine-wide per-machine inclusion list (`HKLM\SOFTWARE\Microsoft\VSTO\Security\Inclusion`) is the
   documented alternative to relying on the certificate store.
3. (Done above: the `Manifest` path containing a space resolves.)
4. An `HKLM` add-in shows in Office as not user-manageable (expected for a machine-wide add-in).
5. On a managed machine, whether Group Policy blocks the `LocalMachine\TrustedPublisher` import (then deploy it through GPO).
6. 32-bit Office on 64-bit Windows: the `WOW6432Node` view is the one written.
7. `uninstall.ps1` (machine-wide) run for real, with and without `-RemoveTrustedCert`.
