# CLDNT-mcsIdentifier

An offline lookup tool for AWS managed-services teams. Paste an AWS account ID, a client name, or a whole alert email, and it searches every sheet of your client workbook. It shows which client the account belongs to, the team that looks after it, and every related row across the workbook.

Built at **Cloudelligent** for the Managed Cloud Services team. The first use was identifying clients from FinOps cost alerts, which don't name the client, and it's now used for day-to-day account and client lookups.

The whole tool is **one HTML file**. There's nothing to install, no server and no build step, and the data never leaves your computer.

---

## Features

**Search**
- Finds 12-digit AWS account IDs anywhere in pasted text: alert emails, tickets, invoice lines. It handles several IDs at once.
- Matches IDs written as `1234-5678-9012`, and IDs where Excel dropped the leading zeros.
- Searches by client name or part of a name, and by partial account ID.
- Searches as you type, and keeps a list of recent lookups.

**Understanding the workbook**
- Reads every sheet and every column of an `.xlsx`, `.xlsm` or `.csv` file, including hidden sheets.
- Finds the real header row (Client, Pod, Pod Lead, PM, Primary/Secondary Engineer, Account ID…) even when there are title rows above it.
- Supports **client heading rows**: a row containing just a client name, with that client's accounts listed underneath. It also supports merged cells and normal "Client" columns.
- Connects the same client across sheets that are laid out differently.

**Results**
- **Account ID search:** the client and how they were identified, the team (pod, pod lead, PM, engineers), the full row for the account, every other account for that client, and the client's rows from all other sheets.
- **Client name search:** a client profile with the team, every account ID, and the related rows grouped by sheet.
- **Copy details** copies a clean, ready-to-paste summary for a ticket or reply.

**Security and privacy**
- **Runs fully offline.** A strict Content Security Policy blocks every network request, so the page can't load or send anything.
- The workbook is read in the browser using a built-in `.xlsx` reader. There are no third-party libraries.
- **Passcode lock:** the saved copy of the workbook is encrypted with AES-256-GCM, using a key derived from your passcode with PBKDF2-SHA256 (310,000 iterations). It unlocks once per browser session and locks again when the browser or tab closes, or when you press **Lock**.
- The passcode can't be recovered. If it's forgotten, the saved copy is erased and the workbook is loaded again. The Excel file itself is never changed.

**Convenience**
- **Refresh** re-reads the workbook from disk to pick up saved changes, without choosing the file again (Chromium-based browsers).
- Light and dark mode, following the Windows setting by default.
- Optional company logo, chosen once from your computer.
- Keyboard shortcuts: `/` to focus the search box, `Esc` to clear.

---

## Getting started

1. Download `index.html`. You can rename it, for example to `AWS Account Finder.html`.
2. Double-click it to open it in **Microsoft Edge, Google Chrome or Dia**.
3. Set a passcode.
4. Drag your client workbook onto the page, or click **Choose .xlsx file**.
5. Search.

The next time, open the file and enter your passcode. The workbook is still loaded. When the Excel file changes, click **Refresh**.

### Sharing with a team

Put `index.html` and the workbook in a SharePoint or OneDrive folder that only your team can access, and have each person sync it to their laptop. Access is then controlled by everyone's Microsoft work sign-in, and **Refresh** picks up the latest synced workbook.

> SharePoint Online doesn't run custom HTML pages, so open the synced copy from your computer instead of opening it in SharePoint.

---

## Excel version (optional)

`AccountFinder.bas` is a VBA module that adds a **Finder** sheet to the workbook itself, with the same lookup logic, for people who prefer to stay in Excel on the desktop.

1. Back up the workbook.
2. Press `Alt + F11`, then **File → Import File…** and choose `AccountFinder.bas`.
3. Run **Debug → Compile VBAProject**.
4. Press `Alt + F8` and run **SetupFinder**.
5. Save as **Excel Macro-Enabled Workbook (`.xlsm`)**.

Type an ID or client name in `C3` on the **Finder** sheet, then press `Ctrl + Shift + F` or click **Find**. To run the search when you press Enter, paste this into the Finder sheet's code window (right-click the tab, then **View Code**):

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then FindAccount
End Sub
```

The macro needs Excel on the desktop with macros allowed. It doesn't run in Excel for the web.

---

## Files

| File | What it is |
|---|---|
| `index.html` | The Account Finder app: HTML, CSS and JavaScript in one file |
| `AccountFinder.bas` | Optional VBA module for the in-Excel Finder sheet |

---

## Browser support

Works best in current versions of **Edge, Chrome and Dia**. Firefox and Safari can search and use the passcode lock, but they don't support **Refresh**. In those browsers, use **Choose another file** to load an updated workbook.

---

## Limitations

- Old `.xls` files and password-protected workbooks aren't supported. Re-save them as a normal `.xlsx` first.
- A row with only one filled cell is treated as a client heading. A stray note on its own row can be mistaken for a client name, so the result always shows which row the client came from.
- Clients are matched across sheets by name. Small differences (`Ltd`, `&`, spacing) are ignored, but short forms or nicknames may not match.

---

## Data safety

**Never commit client data to this repository.** Keep workbooks out of Git, for example with this `.gitignore`:

```gitignore
*.xlsx
*.xlsm
*.xls
*.csv
```

---

Made by [github.com/mhd-sami](https://www.github.com/mhd-sami) at Cloudelligent.

