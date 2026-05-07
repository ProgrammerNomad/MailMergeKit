# MailMergeKit

> A Microsoft Word add-in that creates personalized Outlook draft emails using Word's mail merge data source - giving you full review control before sending.

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-0.0.2-orange.svg)](https://github.com/ProgrammerNomad/MailMergeKit/releases/tag/v0.0.2)
[![Office](https://img.shields.io/badge/Office-2007%20SP2%2B-blue.svg)](#system-requirements)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](#system-requirements)

---

## What It Does

MailMergeKit adds a **"Send via MailMergeKit"** button to Word's **Mailings** tab. When clicked, it reads your Word mail merge data source and creates one **Outlook draft email per recipient** - personalized and ready to review before you send anything.

**No SMTP setup. No cloud services. Everything runs locally using your existing Outlook account.**

---

## Features (v0.0.2)

- **Word ribbon button** - appears in the Mailings tab, zero extra setup
- **Reads Word's native mail merge data source** - works with Excel, CSV, Access, or any source Word supports
- **Personalized subject lines** - use merge fields like `Hello «FirstName»` in the subject
- **CC / BCC support** - map columns from your data source
- **Attachment support** - static file path per recipient (semicolon-separated for multiple files)
- **Draft mode** - all emails go to Outlook Drafts so you review before sending
- **Settings dialog** - pick your email field and subject template from a simple UI
- **100% local** - no data leaves your machine

---

## System Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10 or later |
| **Word** | Microsoft Word 2007 SP2 / 2010 / 2013 / 2016 / 2019 / 2021 / 365 (desktop) |
| **Outlook** | Microsoft Outlook 2007 SP2 / 2010 / 2013 / 2016 / 2019 / 2021 / 365 (desktop, must be open) |
| **Runtime** | [VSTO Runtime](https://aka.ms/vsto-runtime) - installed automatically by setup.exe |

> **Web versions of Word/Outlook (browser) are not supported.** Requires the desktop Office apps installed on Windows.
>
> **Office 2007 users:** Service Pack 2 (SP2) or later is required. Office 2007 RTM and SP1 are not supported by the VSTO Runtime.

---

## Installation

1. Go to the [Releases page](https://github.com/ProgrammerNomad/MailMergeKit/releases/tag/v0.0.2)
2. Download **`MailMergeKit-v0.0.2-installer.zip`**
3. Extract the zip
4. Run **`setup.exe`**
5. Follow the prompts - the VSTO runtime installs automatically if missing
6. Open (or restart) Microsoft Word
7. Go to the **Mailings** tab - you'll see the **MailMergeKit** section

> **Security note:** Windows may show a warning because the installer is signed with a self-signed certificate. Click **Install** to proceed - this is normal for open-source tools without a paid code-signing certificate.

---

## How to Use

### Step 1 - Set up your Word document

1. Open Word and write your email as a document
2. Go to **Mailings → Select Recipients** and connect your data source (Excel, CSV, etc.)
3. Insert merge fields in the document body using **Mailings → Insert Merge Field**

**Your data source should have at minimum:**

| Column | Required | Example |
|---|---|---|
| `Email` | Yes | `john@example.com` |
| Any name/detail fields | No | `FirstName`, `Company`, `Domain` |
| `CC` | No | `manager@example.com` |
| `BCC` | No | `archive@example.com` |
| `Attachment` | No | `invoice.pdf` or `invoice.pdf;receipt.pdf` |

### Step 2 - Run MailMergeKit

1. Click **Mailings → Send via MailMergeKit**
2. A settings dialog opens showing all fields from your data source
3. Select the column that contains email addresses (e.g. `Email`)
4. Enter a subject template - you can use merge field names in `«»` brackets, e.g.:
   ```
   Your domain «Domain» expires on «ExpiryDate»
   ```
5. Click **OK**

### Step 3 - Review and send

1. Open Outlook
2. Go to your **Drafts** folder
3. You'll see one draft per recipient, fully personalized
4. Review, edit if needed, then click **Send**

---

## Example

**Data source (Excel):**

| FirstName | Email | Domain | ExpiryDate |
|---|---|---|---|
| John | john@example.com | example.com | 2026-06-01 |
| Jane | jane@company.com | company.com | 2026-07-15 |

**Subject template:**
```
Hi «FirstName» - your domain «Domain» expires on «ExpiryDate»
```

**Result in Outlook Drafts:**
- Email 1 to `john@example.com` - Subject: `Hi John - your domain example.com expires on 2026-06-01`
- Email 2 to `jane@company.com` - Subject: `Hi Jane - your domain company.com expires on 2026-07-15`

---

## Current Limitations (v0.0.1)

This is an early prototype. The following are **not yet implemented:**

- No preview before running the merge
- No progress bar during merge
- No test email mode (send to yourself first)
- No retry on failure
- Body merge uses Word's content as-is - complex field formatting may vary
- No undo - drafts must be deleted manually if something goes wrong

See [GETTING_STARTED.md](GETTING_STARTED.md) for the full known issues list.

---

## Roadmap

### v0.1.0 (next)
- Progress bar during merge
- Test email mode (send first draft to yourself before the full run)
- Better error messages in a dialog
- Auto-start Outlook if not running
- HTML body merge improvements

### v0.2.0
- Preview individual emails before running
- Per-recipient dynamic attachments from a data column
- Merge field picker UI

---

## Building from Source

### Prerequisites

- **Visual Studio 2022** with:
  - .NET Desktop Development workload
  - Office/SharePoint Development workload
- **Microsoft Office** (Word + Outlook desktop)

### Steps

```bash
git clone https://github.com/ProgrammerNomad/MailMergeKit.git
cd MailMergeKit
```

1. Open `MailMergeKit.sln` in Visual Studio 2022
2. Press `Ctrl+Shift+B` to build
3. Press `F5` - Word launches with the add-in loaded for debugging

### Generate a new installer

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" `
  "MailMergeKit.sln" /t:Publish /p:Configuration=Release /p:Platform="Any CPU" `
  /p:PublishDir="src\MailMergeKit.WordAddin\publish\" /p:ApplicationVersion="0.0.2.0"

Compress-Archive -Path "src\MailMergeKit.WordAddin\publish\*" `
  -DestinationPath "MailMergeKit-v0.0.2-installer.zip" -Force
```

---

## Project Structure

```
MailMergeKit/
├── src/
│   └── MailMergeKit.WordAddin/
│       ├── ThisAddIn.cs             # VSTO entry point
│       ├── Globals.cs               # VSTO globals
│       ├── Ribbon/
│       │   └── MailMergeRibbon.cs   # Word ribbon button
│       ├── Services/
│       │   ├── MergeController.cs   # Reads data source, drives the merge
│       │   └── OutlookMailer.cs     # Creates Outlook draft emails via COM
│       ├── Models/
│       │   └── RecipientData.cs     # Per-recipient data model
│       └── UI/
│           └── SettingsForm.cs      # Settings dialog
├── docs/                            # Documentation
├── examples/                        # Sample data and templates
├── installer/                       # Installer project (future)
└── tests/                           # Tests (future)
```

---

## Contributing

Contributions are welcome. Please open an issue first to discuss what you'd like to change.

1. Fork the repo
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -m "Add my feature"`
4. Push and open a Pull Request

---

## License

[MIT](LICENSE) - free to use, modify, and distribute.

---

## Support

- **Bug reports:** [Open an issue](https://github.com/ProgrammerNomad/MailMergeKit/issues)
- **Questions:** [Start a discussion](https://github.com/ProgrammerNomad/MailMergeKit/discussions)
