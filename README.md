<p align="right">
  <strong>English</strong> | <a href="README.zh-CN.md">简体中文</a>
</p>

# OutlookQuickMove

A keyboard-first VSTO add-in for **classic Outlook on Windows** that moves the selected mail
to a folder you find by typing — no dragging, no expanding the folder tree.

## Features

- **Quick Move dialog** — type to filter folders by full path, `Up`/`Down` to pick, `Enter` to
  move. Works on one or many selected messages.
- **Go to Folder** — search the same folder list and switch the active Outlook window without
  moving mail.
- **Frequent folders first** — destinations you use most float to the top, so common moves are
  one keystroke away (an empty search box pre-selects your most-used folder).
- **Mark as read before moving** — optional, and your choice is remembered across restarts.
- **Per-data-file filtering** — choose which Outlook data files (stores) are searched.
- **Safe by design** — never deletes mail, skips non-mail items, skips moves into the same
  folder, releases Outlook COM objects deterministically, and logs problems instead of failing
  silently.

## Requirements

- Windows with **classic Win32 Outlook** (Microsoft 365 / 2016+; not "New Outlook" or Outlook on
  the web). Developed against `Microsoft Outlook for Microsoft 365 MSO (Version 2605 Build
  16.0.20026.20076) 64-bit`.
- **.NET Framework 4.8**.
- To build: **Visual Studio 2022** with the *Office/SharePoint development* workload (VSTO). See
  `DEV_SETUP.md`.

## Build

- Open `OutlookQuickMove.sln` in Visual Studio and build. Pressing **F5** launches Outlook with
  the add-in loaded for debugging.
- Or from a command line: `msbuild OutlookQuickMove.sln /p:Configuration=Debug`.

ClickOnce manifest signing is **disabled by default** (`SignManifests` is `false`) and no
signing key is included, so the project builds from a clean clone with no extra setup. Before
distributing to other machines, supply your own code-signing certificate and set
`SignManifests` / `ManifestKeyFile` (and a thumbprint) in
`OutlookQuickMove/OutlookQuickMove.csproj`. Never commit the private key (`*.pfx` is git-ignored).

## Install

Build the project first so the `.vsto` deployment manifest exists under the output folder.

Debug build manifest:

```powershell
C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto
```

Use `VSTOInstaller.exe` with a `file:///` URI. This avoids path parsing issues when the local path contains line breaks, copied prompt markers, or other hidden characters.

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri

& $installer /Install $vstoUri
```

Do not split the `.vsto` path across multiple lines inside the quoted string.

After installation, restart Outlook. The add-in creates its own top-level Ribbon tab:

`Quick Move` ribbon tab -> `Actions` group

Available buttons:

- `Quick Move`: search and move the selected mail items.
- `Go to Folder`: search for a folder and switch the active Outlook window to it (navigate, no
  move). Uses the same Data Files selection and the same type-to-filter dialog as Quick Move.
- `Undo Quick Move...`: open a checklist of recent moves and send the chosen items back to
  their original folders. Enabled only when there is move history to undo.
- `Settings`: a tabbed dialog — `Data Files` (which Outlook data files are searched),
  `Frequent Folders` (the remembered-targets cap and list), and `Undo History` (the
  remembered-moves cap and a clear option).

In the Quick Move dialog, type to filter folders, use `Up` / `Down` to change the highlighted candidate, and press `Enter` to confirm.

## Go to a folder

`Go to Folder` reuses the Quick Move folder picker to navigate instead of move: press the button,
type to filter, use `Up` / `Down`, and press `Enter` (or `Go`) to switch the active explorer to
the chosen folder. It searches the same folders as Quick Move (the saved Data Files selection)
and floats your most-used Quick Move destinations to the top for quick access, but it never
changes any mail and never records a move — so the frequent-folders and undo history are
untouched.

## Frequent folders

Quick Move remembers the folders you move to most often and floats them to the top of the
search results (most-used first), so your common destinations are one keystroke away — with an
empty search box, the most-used folder is already selected. Each confirmed move updates the
usage count. Usage is tracked for every folder you use, so a newly used folder accumulates
counts and rises into the list over time rather than being kept out once the list is full.

In `Settings` -> `Frequent Folders` you can set how many folders to **show** at the top of the
Quick Move list (0–100, default 20; `0` turns the feature off), review all tracked folders
ordered by usage, and delete or clear entries. The data persists across Outlook restarts in:

```powershell
$env:APPDATA\OutlookQuickMove\frequent-targets.txt
$env:APPDATA\OutlookQuickMove\frequent-targets-max.txt
```

## Undo a move

Outlook's native `Ctrl+Z` cannot reverse moves made by the add-in (the object model has no API
to push onto Outlook's undo stack), so Quick Move keeps its own history. Every move records, per
item, where it came from, where it went, and its original read state.

Press `Undo Quick Move...` to open a checklist of recent moves (newest first). The most recent
action's items are pre-checked for the common "undo what I just did" case; check or uncheck any
combination across actions, then press `Undo Selected` to move them back to their original
folders (read/unread state is restored too). Items that can no longer be found (deleted or moved
again since) are reported and dropped from the list; anything that fails for another reason is
kept so you can retry. `Clear History` empties the list without moving any mail.

In `Settings` -> `Undo History` you can set how many recent moves are remembered (0-500, default
50; `0` turns recording off but keeps existing history until cleared) and clear the history. The
data persists across Outlook restarts in:

```powershell
$env:APPDATA\OutlookQuickMove\undo-history.txt
$env:APPDATA\OutlookQuickMove\undo-max.txt
```

If no store filter has been saved yet, Quick Move searches folders from all Outlook data files. The Settings dialog supports selecting multiple data files.

The selected data files are persisted in:

```powershell
$env:APPDATA\OutlookQuickMove\store-filter.txt
```

PST/OST-backed stores are matched by file path so the selection survives Outlook restarts and VSTO deployment refreshes.

### Folder list caching

Building the folder list walks every folder in every selected data file, which is the heaviest
operation, so the result is cached and reused for a short window (about 2 minutes) instead of being
rebuilt on every Quick Move / Go to Folder. This noticeably reduces memory churn and MAPI resource
pressure on large or multi-mailbox profiles. The cache refreshes automatically after the window, or
immediately when you save Settings. A folder you create or rename directly in Outlook may therefore
take up to about 2 minutes to appear; save Settings to refresh it right away. If some folders cannot
be read during enumeration, the Quick Move summary groups the warnings by cause and points you to
the diagnostic log for the exact folders.

The `Mark as read before moving` checkbox remembers its last confirmed state across dialog
closes and Outlook restarts, persisted next to the store filter in:

```powershell
$env:APPDATA\OutlookQuickMove\mark-as-read.txt
```

Dialog windows use the embedded icon asset at `OutlookQuickMove\Assets\QuickMove.ico`.

After changing and rebuilding the add-in, do a clean reinstall with the same `.vsto`
path to make sure Outlook refreshes the VSTO deployment cache. See
`Clean reinstall` below.

If the button is not visible, check whether Outlook loaded the add-in:

`File` -> `Options` -> `Add-ins` -> `Manage: COM Add-ins` -> `Go...`

The add-in writes a lightweight diagnostic log (startup/Ribbon trace plus any handled
errors) in both Debug and Release builds to:

```powershell
$env:TEMP\OutlookQuickMove.log
```

Useful entries include `ThisAddIn_Startup`, `CreateRibbonExtensibilityObject`, and
`GetCustomUI ribbonId=Microsoft.Outlook.Explorer`, as well as any folder-enumeration or
move failures. The file is best-effort, never blocks the add-in, and is capped at ~1 MB
(it is recreated when it exceeds that size). Note it may contain folder paths and the
subjects of mail that failed to move.

## Clean reinstall

Use this when refreshing a local Debug build after code changes.

1. Close Outlook.
2. Uninstall the current VSTO deployment with the same `.vsto` URI.
3. Clear the current user's ClickOnce application cache.
4. Install again.

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri
$mage = 'C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe'

Get-Process Outlook,VSTOInstaller,dfsvc -ErrorAction SilentlyContinue |
    Stop-Process -Force

& $installer /Uninstall $vstoUri
& $mage -cc
& $installer /Install $vstoUri
```

If uninstall reports that the customization is not installed or has already been
uninstalled, that is acceptable when you are cleaning up a broken local install.
Continue with the cache-clear and install steps.

## Uninstall

Use the same `.vsto` deployment manifest location that was used for installation.

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri

& $installer /Uninstall $vstoUri
```

Restart Outlook after uninstalling.

If the local build output has been cleaned or moved, rebuild the same configuration first so the `.vsto` manifest path exists again, then run the uninstall command.

## Reinstall troubleshooting

If reinstall still fails with this error:

```text
Unable to install this application because an application with the same identity is already installed.
```

and a second uninstall says:

```text
Attempting to uninstall a customization that has not been installed on this computer or has already been uninstalled from this computer.
```

then Outlook's add-in registration is already gone, but the ClickOnce/VSTO cache still
has the old deployment identity. The cache identity includes values such as
`OutlookQuickMove.vsto`, `version="1.0.0.0"`, and the manifest public key token.

First confirm there is no regular Outlook add-in registration left:

```powershell
Get-ItemProperty -Path `
  'HKLM:\Software\Microsoft\Office\Outlook\Addins\*',`
  'HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins\*',`
  'HKCU:\Software\Microsoft\Office\Outlook\Addins\*' `
  -ErrorAction SilentlyContinue |
  Where-Object {
      $_.PSChildName -like '*OutlookQuickMove*' -or
      $_.FriendlyName -like '*OutlookQuickMove*' -or
      $_.Manifest -like '*OutlookQuickMove*'
  } |
  Select-Object PSPath,PSChildName,FriendlyName,LoadBehavior,Manifest
```

Then look for stale ClickOnce/VSTO cache entries:

```powershell
$cacheRoot = Join-Path $env:LOCALAPPDATA 'Apps\2.0'

Get-ChildItem $cacheRoot -Recurse -ErrorAction SilentlyContinue |
  Where-Object {
      $_.Name -like '*.vsto' -or
      $_.Name -like '*.manifest' -or
      $_.Name -like '*.cdf-ms'
  } |
  Select-String -Pattern 'OutlookQuickMove' -List -ErrorAction SilentlyContinue |
  Select-Object Path,LineNumber,Line

reg query "HKCU\Software\Microsoft\VSTO" /f OutlookQuickMove /s
```

If stale entries remain after `mage.exe -cc`, remove only the entries that match this
add-in. Do **not** delete the whole `%LOCALAPPDATA%\Apps\2.0` folder; it may contain
other Office add-ins.

The first script below derives the manifest public key token from the current `.vsto`
file and previews the matched cache targets.

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri
$cacheRoot = (Resolve-Path (Join-Path $env:LOCALAPPDATA 'Apps\2.0')).Path

[xml]$vstoManifest = Get-Content -LiteralPath $vsto
$identity = $vstoManifest.assembly.assemblyIdentity
$token = $identity.publicKeyToken
if ([string]::IsNullOrWhiteSpace($token)) {
    throw 'Could not read publicKeyToken from the .vsto manifest.'
}

Get-Process Outlook,VSTOInstaller,dfsvc -ErrorAction SilentlyContinue |
    Stop-Process -Force

$manifestTargets = Get-ChildItem $cacheRoot -Recurse -File -ErrorAction SilentlyContinue |
  Where-Object {
      ($_.Name -like '*.manifest' -or $_.Name -like '*.cdf-ms') -and
      $_.Name -like "*$token*"
  }

$payloadTargets = Get-ChildItem $cacheRoot -Recurse -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -like "outl*$token*" }

$targets = @($manifestTargets.FullName) + @($payloadTargets.FullName) |
  Sort-Object -Unique

$targets
```

Review the preview. If it lists only OutlookQuickMove cache entries, run the cleanup:

```powershell
foreach ($target in $targets) {
    $resolved = (Resolve-Path -LiteralPath $target).Path
    if (-not $resolved.StartsWith($cacheRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to delete outside ClickOnce cache: $resolved"
    }
    if ($resolved -notlike "*$token*") {
        throw "Refusing to delete a path that does not contain the manifest token: $resolved"
    }

    Remove-Item -LiteralPath $resolved -Recurse -Force
}

Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion' -ErrorAction SilentlyContinue |
  Where-Object {
      (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).Url -eq $vstoUri
  } |
  Remove-Item -Recurse -Force

& $installer /Install $vstoUri
```

As a development-only workaround, incrementing the deployment manifest version also
changes the ClickOnce identity and avoids the collision. Prefer cleaning the stale
cache first so local install state stays understandable.

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) for development
setup, the manual test checklist, and the project conventions (COM hygiene, never deleting
mail, and logging instead of failing silently).

## License

[MIT](LICENSE) © r404r
