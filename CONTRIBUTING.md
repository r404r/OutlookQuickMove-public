# Contributing to OutlookQuickMove

Thanks for your interest in improving OutlookQuickMove! This is a small, focused VSTO add-in
for classic Outlook on Windows. The guiding principle is **reliability and predictable behavior
over cleverness** — it touches people's mailboxes, so correctness and safety come first.

## Development setup

- Windows with classic Win32 Outlook (Microsoft 365 / 2016+). New Outlook and Outlook on the
  web are out of scope.
- Visual Studio 2022 with the **Office/SharePoint development** workload (VSTO) and **.NET
  Framework 4.8** targeting pack. See `DEV_SETUP.md` for the full list.
- Open `OutlookQuickMove.sln`, then **F5** to launch Outlook with the add-in loaded for
  debugging. Manifest signing is disabled by default, so a clean clone builds with no extra
  setup.

## Building

- IDE: build in Visual Studio (Debug or Release).
- CLI: `msbuild OutlookQuickMove.sln /p:Configuration=Debug`.

## Testing

There is no automated test suite — the add-in is driven by the live Outlook object model, which
is hard to mock. Verify changes manually in Outlook. A good baseline checklist:

- No message selected; one message; many messages.
- Selection includes a non-mail item (meeting, task) — it is skipped, not crashed on.
- Mark-as-read on vs. off (and that the setting is remembered after restarting Outlook).
- Target folder under a secondary / shared / archive store.
- Duplicate folder names disambiguated by full path.
- Moving into the folder the mail is already in (should be skipped).
- Frequent folders: repeated moves raise a folder to the top; a brand-new folder accumulates and
  eventually appears; Settings cap / delete / clear behave and persist.
- `Esc` cancels, `Enter` confirms; restart Outlook and confirm the add-in still loads.

## Conventions (please follow)

These keep the add-in safe and stable — most existing code already does this:

- **Release COM objects.** Every Outlook RCW you obtain (explorers, selections, items,
  `NameSpace`, `Stores`, `MAPIFolder`, `Folders`, the result of `MailItem.Move`, …) must be
  released in a `finally` via `ComUtil.Release`. Avoid chained `a.b.c` access on COM objects —
  the intermediates leak. The only object you do not release is `Globals.ThisAddIn.Application`.
- **Never delete mail** or perform destructive Outlook operations.
- **Never silently swallow failures.** Show the user a concise summary and route diagnostics
  through `QuickMoveLog` (which works in Release, unlike `Debug.WriteLine`).
- **Persist per-user state** under `%APPDATA%\OutlookQuickMove\` via `QuickMovePaths.DataFile`
  (survives Outlook restarts and VSTO deployment refreshes); don't add a second copy in
  per-version application settings.
- **Match the surrounding style** — naming, brace placement, and comment density of nearby code.
- **Never commit secrets** — signing keys (`*.pfx`) are git-ignored; keep it that way.

## Submitting changes

- Keep pull requests focused; describe what you changed and how you verified it in Outlook.
- For UI or behavior changes, include the manual test cases you ran.
- By contributing, you agree your contributions are licensed under the project's
  [MIT License](LICENSE).

## Reporting issues

Please include your Outlook version/build, Windows version, and steps to reproduce. If relevant,
attach the diagnostic log from `%TEMP%\OutlookQuickMove.log` (note it may contain folder paths
and the subjects of mail that failed to move — redact as needed).
