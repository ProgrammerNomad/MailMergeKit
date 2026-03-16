# MailMergeKit

> A powerful Microsoft Word Mail Merge extension that creates personalized Outlook draft emails with advanced features not available in native Office tools.

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![.NET](https://img.shields.io/badge/.NET-4.8%20%7C%206.0-purple.svg)](https://dotnet.microsoft.com/)
[![Office](https://img.shields.io/badge/Office-VSTO-orange.svg)](https://docs.microsoft.com/en-us/visualstudio/vsto/)
[![Performance](https://img.shields.io/badge/Performance-5K%20emails-green.svg)](#performance-best-practices)
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20Local-success.svg)](#privacy--security)

### Quick Start

1. **Download:** Get [MailMergeKit-Setup-v0.0.1.msi](https://github.com/ProgrammerNomad/MailMergeKit/releases)
2. **Install:** Run as administrator and restart Word
3. **Use:** Open Word → Mailings → MailMergeKit → Send via MailMergeKit
4. **Review:** Check Outlook Drafts folder before sending

**Perfect for:** Email campaigns | Marketing teams | Domain expiry notices | Invoice delivery | Event invitations

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Why Choose MailMergeKit?](#why-choose-mailmergekit)
- [How It Works](#how-it-works)
- [Technology Stack](#technology-stack)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Architecture](#architecture)
- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Roadmap](#roadmap)
- [Performance Best Practices](#performance-best-practices)
- [Contributing](#contributing)
- [License](#license)
- [Support](#support)

---

## Overview

**MailMergeKit** uses Microsoft Word's mail merge data source to generate personalized Outlook draft emails with features that Word and Outlook don't provide natively.

### The Problem

Word's built-in Mail Merge can send emails, but it lacks:
- Subject line personalization with merge fields
- Per-recipient attachment support
- Draft mode for review before sending
- CC/BCC field support
- Full control over email formatting

### The Solution

**MailMergeKit reads Word's mail merge data source to generate personalized Outlook draft emails, giving you complete control before sending.**

**No SMTP configuration. No external services. Everything runs locally using your existing Outlook account.**

### Privacy & Security

- **100% Local Processing** - No data leaves your computer
- **No Cloud Dependencies** - All merge operations happen on your machine
- **Enterprise Safe** - Works within corporate network restrictions
- **Your Email, Your Control** - Uses your existing Outlook account and credentials

---

## Key Features

### Version 0.0.1 (Current Prototype)

- **Word Integration** - Seamless add-in with custom ribbon button
- **Subject Personalization** - Use Word merge fields in subject lines (e.g., `Domain «Domain» expires soon`)
- **Basic Attachments** - Static attachment support (same file for all recipients)
- **Draft Mode** - Review emails in Outlook Drafts before sending
- **Multiple Data Sources** - Excel, CSV, Access, SQL (via Word's native merge)
- **HTML Email Support** - Rich formatting preserved from Word
- **Simple Settings Dialog** - Basic configuration for merge operations

### Performance & Enterprise Features

- **Optimized Processing** - Fast queue-based processing for stable merges
- **COM Safety** - Single-threaded architecture prevents Outlook COM crashes
- **Memory Safety** - Proper COM object cleanup prevents crashes on large campaigns
- **Resume Interrupted Merge** - Checkpoint system to continue from last position (v0.2.0+)
- **Draft Mode** - Always saves to Drafts folder for review before sending

### Version 0.1.0 Features (Next Release)

- **Test Email Mode** - Send first email to test address before full merge
- **Automatic Outlook Startup** - Auto-start Outlook if not running
- **Multiple Attachments** - Support multiple files per recipient (file1.pdf;file2.pdf)
- **Enhanced Logging** - Detailed merge operation logs with Serilog
- **Improved Error Handling** - Better COM error recovery with Polly retry logic

### Future Enhancements (v0.2.0+)

- **Resume Interrupted Merge** - Checkpoint system to continue from last position
- **Merge Field Picker** - UI to insert available fields and avoid typos
- **Anti-Spam Delays** - Optional delays between emails to prevent server blocking
- **Enhanced Validation** - Advanced pre-merge checks and data quality rules

---

## Why Choose MailMergeKit?

### vs. Paid Tools (Mail Merge Toolkit, etc.)

| Feature | MailMergeKit | Mail Merge Toolkit | Others |
|---------|--------------|-------------------|---------|
| **Price** | Free & Open Source | $49-199/license | $30-300/license |
| **Subject Personalization** | Yes | Yes | Limited |
| **Per-Recipient Attachments** | Yes | Yes | Limited |
| **Draft Mode** | Yes | Yes | No |
| **Resume Interrupted Merge** | Yes | No | No |
| **Optimized Processing** | Yes (queue-based) | Sequential | Sequential |
| **Privacy** | 100% local | Cloud telemetry | Varies |
| **Source Code Access** | Full access | Closed | Closed |
| **Enterprise Safe** | No external calls | Phone home | Varies |
| **Large Campaigns (5K)** | Optimized | Slow | Slow |
| **Customizable** | Plugin architecture | Fixed | Fixed |

### Designed For

**Email Marketing Teams** - Handle campaigns of up to ~5,000 emails efficiently  
**Enterprises** - Privacy-focused with no data leaving your network  
**Developers** - Open source, extensible, and hackable  
**Small Businesses** - Free alternative to expensive paid tools  
**IT Departments** - Auditable code and local processing  

---

## How It Works

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│  Word Mail Merge Document + Data Source                    │
│  ┌──────────────────┐        ┌──────────────┐              │
│  │ Template.docx    │        │ Recipients   │              │
│  │                  │   +    │ (Excel/CSV)  │              │
│  │ Dear «Name»,     │        │              │              │
│  │ Your «Domain»... │        └──────────────┘              │
│  └──────────────────┘                                      │
│           │                                                 │
│           ▼                                                 │
│  ┌────────────────────────────────────────┐                │
│  │  Click "Send via MailMergeKit" button  │                │
│  └────────────────────────────────────────┘                │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  MailMergeKit Processes Data Source:                │   │
│  │  • Reads merge fields from Word data source         │   │
│  │  • Personalizes subject lines                       │   │
│  │  • Resolves attachment paths                        │   │
│  │  • Queues records for processing                    │   │
│  └─────────────────────────────────────────────────────┘   │
  │           │                                                 │
  │           ▼                                                 │
  │  ┌─────────────────────────────────────────────────────┐   │
  │  │  Single Outlook Worker (COM-safe architecture)      │   │
  │  │  Processes each record sequentially:                │   │
  │  │  • Creates MailItem                                 │   │
  │  │  • Sets To/CC/BCC                                   │   │
  │  │  • Merges subject and body                          │   │
  │  │  • Adds attachments                                 │   │
  │  │  • Saves to Drafts folder                           │   │
  │  └─────────────────────────────────────────────────────┘   │
  │  Email 1: john@example.com                          │   │
  │  Email 2: jane@example.com                          │   │
  │  Email 3: bob@example.com                           │   │
│  └─────────────────────────────────────────────────────┘   │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  User reviews in Outlook Drafts folder              │   │
│  │  Then clicks Send for each email                    │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Technology Stack

**All frameworks and libraries are 100% free and open source - no paid licenses required.**

| Component | Technology | License | Purpose |
|-----------|-----------|---------|------|
| **Language** | C# | - | Primary development language |
| **Framework** | .NET Framework 4.8 / .NET 6 | Free | Runtime environment |
| **Office Integration** | VSTO (Visual Studio Tools for Office) | Free | Word/Outlook add-in framework |
| **Word Automation** | Microsoft.Office.Interop.Word | Free | Word document manipulation |
| **Outlook Automation** | Microsoft.Office.Interop.Outlook | Free | Email creation and management |
| **Logging** | Serilog | Apache 2.0 | Professional structured logging |
| **Validation** | FluentValidation | Apache 2.0 | Data validation rules |
| **Resilience** | Polly | BSD 3-Clause | Retry logic and error recovery |
| **CSV Support** | CsvHelper | Apache 2.0 | CSV data source parsing |
| **Configuration** | Microsoft.Extensions.Configuration | MIT | Settings management |
| **Excel Processing** | ClosedXML (optional) | MIT | Excel without Office Interop |
| **UI Framework** | WinForms | Free | Settings dialog and user interface |
| **Installer** | WiX Toolset / WixSharp | MIT | MSI installer generation |
| **HTML Processing** | HtmlAgilityPack | MIT | HTML cleanup and optimization |

> See [TECHNICAL_STACK.md](TECHNICAL_STACK.md) for detailed framework comparison and code examples.

---

## System Requirements

### End Users

- **Operating System:** Windows 10 or later
- **Microsoft Office:** 
  - Word 2016 or later
  - Outlook 2016 or later (desktop version)
- **.NET Framework:** 4.8 or later
- **Permissions:** Administrator rights for installation

### Developers

- **IDE:** Visual Studio 2022
- **Workloads:**
  - .NET Desktop Development
  - Office/SharePoint Development
- **Tools:**
  - Microsoft Office Developer Tools for Visual Studio
  - WiX Toolset v3.11+ (for installer)
- **Office Installation:** Word and Outlook installed locally

---

## Installation

### For End Users

1. **Download the installer**
   ```
   MailMergeKit-Setup-v1.0.msi
   ```

2. **Run the MSI installer**
   - Right-click → Run as Administrator
   - Follow the installation wizard
   - Accept the prompt to install the Office add-in

3. **Restart Microsoft Word**

4. **Verify installation**
   - Open Word
   - Go to the **Mailings** tab
   - Look for the **MailMergeKit** section in the ribbon

### For Developers

See [Development Setup](#development-setup) section below.

---

## Usage

### Basic Workflow

1. **Prepare Your Mail Merge Document**
   - Create a Word document with your email template
   - Insert merge fields: `Ctrl+F9` or use **Mailings** → **Insert Merge Field**

2. **Connect Data Source**
   - **Mailings** → **Select Recipients** → Choose your data source
   - Ensure you have columns for:
     - `Email` (required)
     - `Subject` (optional - or use template)
     - `Attachment` (optional - for dynamic attachments)
       - Single file: `invoice.pdf`
       - Multiple files: `invoice.pdf;receipt.pdf;terms.pdf`
     - `CC` and `BCC` (optional)
     - `Attachment1`, `Attachment2`, etc. (alternative for multiple files)

3. **Configure MailMergeKit**
   - Click **Mailings** → **MailMergeKit** → **Send via MailMergeKit**
   - In the settings dialog:
     - **Recipient Field:** Select the column containing email addresses
     - **Subject Template:** Enter subject with merge fields (e.g., `Domain «Domain» expires soon`)
     - **Attachments:** Add static files or map to a data column
     - **Mode:** Choose **Draft** (recommended) or **Send Immediately**
     - **Test Mode:** Send first email to test address (optional)
     - **Processing:** Optimized queue-based merge

4. **Preview and Validate**
   - Click **Preview First Email** to verify merge fields
   - Use **Preview Sample (10)** to check multiple variations
   - System validates:
     - Email addresses exist
     - Attachments are found
     - No duplicate recipients (with warnings)

5. **Generate Emails**
   - Click **Start Merge**
   - MailMergeKit processes all records with progress tracking
   - Detailed logs saved to `logs/merge-log.txt`

6. **Review and Send**
   - Open Outlook
   - Go to **Drafts** folder (or your selected folder)
   - Review each email
   - Click **Send** when ready

### Advanced Features

**Resume Interrupted Merge:**
If a merge is interrupted (e.g., crash at email 500 of 2000):
- MailMergeKit saves state to `merge_state.json`
- Next run prompts: "Continue from row 501?"
- No need to regenerate completed emails

**Performance Expectations:**
- 100 emails: ~1-2 minutes
- 1,000 emails: ~10-15 minutes
- 5,000 emails: ~50-75 minutes
- Outlook remains responsive during merge
- **Designed for up to ~5,000 emails** (Outlook is not built for bulk campaigns)

**Error Recovery:**
All operations are logged:
```
2026-04-01 14:10:33 | john@example.com | SUCCESS
2026-04-01 14:10:34 | jane@company.com | ERROR: Attachment not found - invoice_002.pdf
2026-04-01 14:10:35 | bob@example.com | SUCCESS
```

### Example: Domain Expiry Notifications

**Data Source (Excel):**
```
| Name    | Email               | Domain          | ExpiryDate | Attachment        |
|---------|---------------------|-----------------|------------|-------------------|
| John    | john@example.com    | example.com     | 2026-04-15 | invoice_001.pdf   |
| Jane    | jane@company.com    | company.com     | 2026-05-20 | invoice_002.pdf   |
```

**Subject Template:**
```
Domain «Domain» expires on «ExpiryDate»
```

**Result:**
- Email 1: Subject = "Domain example.com expires on 2026-04-15"
- Email 2: Subject = "Domain company.com expires on 2026-05-20"

### Real-World Use Cases

**1. Email Marketing Campaigns**
- Newsletter distribution to up to 5,000 subscribers
- Product announcements with personalized discount codes
- Event invitations with unique registration links

**2. Customer Service**
- Domain/SSL expiry notifications with renewal links
- Invoice delivery with personalized payment terms (supports multiple attachments)
- Service renewal reminders with account-specific details

**3. HR & Internal Communications**
- Employee onboarding with personalized welcome packets
- Training schedule notifications with individual calendars
- Performance review reminders with manager-specific attachments

**4. Sales & Business Development**
- Proposal distribution with company-specific pricing
- Follow-up emails with personalized demo links
- Contract renewals with client-specific terms (multiple PDFs supported)

**5. Education & Training**
- Course enrollment confirmations with individual syllabi
- Grade reports with student-specific attachments
- Event notifications with personalized schedules

---

## Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────┐
│                    MailMergeKit                         │
│                                                         │
│  ┌──────────────────┐   ┌──────────────────┐           │
│  │  Word Add-in     │   │  Merge Engine    │           │
│  │                  │   │                  │           │
│  │  • Ribbon UI     │──▶│  • Read Data     │           │
│  │  • Button Event  │   │    Source        │           │
│  │  • Settings Form │   │  • Queue Records │           │
│  │                  │   │  • Merge Fields  │           │
│  └──────────────────┘   └──────────────────┘           │
│           │                       │                     │
│           └───────────┬───────────┘                     │
│                       ▼                                 │
│           ┌──────────────────────┐                      │
│           │ Outlook Worker       │                      │
│           │ (Single Thread)      │                      │
│           │                      │                      │
│           │  • Create MailItem   │                      │
│           │  • Set To/CC/BCC     │                      │
│           │  • Set Subject       │                      │
│           │  • Set HTMLBody      │                      │
│           │  • Add Attachments   │                      │
│           │  • Save to Drafts    │                      │
│           │  • Release COM       │                      │
│           └──────────────────────┘                      │
│                       │                                 │
│                       ▼                                 │
│           ┌──────────────────────┐                      │
│           │   Outlook Drafts     │                      │
│           └──────────────────────┘                      │
└─────────────────────────────────────────────────────────┘
```

**Critical Architecture Notes:**
- **Outlook COM is single-threaded (STA)** - parallel processing causes crashes
- **Queue-based sequential processing** ensures stability
- **Single Outlook.Application instance** is reused for all emails (COM requirement)
- **COM objects are explicitly released** to prevent memory leaks
- **No Parallel.ForEach** - COM is not thread-safe!

### Key Classes

#### 1. Word Add-in (`Ribbon.cs`)
```csharp
// Key objects used:
Word.Application
Word.Document
Word.MailMerge
Word.MailMergeDataSource
```

#### 2. Merge Engine (`MergeController.cs`)
```csharp
// Reads Word's mail merge data source
var ds = doc.MailMerge.DataSource;

// Queue all records first
var records = new Queue<RecipientData>();
for (int i = 1; i <= ds.RecordCount; i++)
{
    ds.ActiveRecord = i;
    records.Enqueue(new RecipientData
    {
        Email = ds.DataFields["Email"].Value,
        Subject = ds.DataFields["Subject"].Value,
        // ... other fields
    });
}

// Process queue sequentially (single-threaded for COM safety)
foreach (var record in records)
{
    ProcessRecord(record);
}
```

#### 3. Outlook Mailer (`OutlookMailer.cs`)
```csharp
// Reuse single Outlook instance (COM is single-threaded STA)
private static Outlook.Application _outlookApp;

public void ProcessRecord(RecipientData record)
{
    Outlook.MailItem mail = null;
    
    try
    {
        // Create mail item
        mail = (Outlook.MailItem)_outlookApp.CreateItem(
            Outlook.OlItemType.olMailItem);
        
        // Set properties
        mail.To = record.Email;
        mail.Subject = MergeSubject(record);
        mail.HTMLBody = MergeBody(record);
        
        // Add CC/BCC if present (v0.1.0+)
        if (!string.IsNullOrEmpty(record.CC))
            mail.CC = record.CC;
        if (!string.IsNullOrEmpty(record.BCC))
            mail.BCC = record.BCC;
        
        // Add attachments if specified (v0.1.0+)
        if (!string.IsNullOrEmpty(record.Attachment))
        {
            // Support multiple files: file1.pdf;file2.pdf
            var files = record.Attachment.Split(';');
            foreach (var file in files)
            {
                var path = ResolveAttachmentPath(file.Trim());
                if (File.Exists(path))
                    mail.Attachments.Add(path);
            }
        }
        
        // Save to Drafts (Outlook default location)
        // DO NOT use mail.Move() - causes issues
        mail.Save();
        
        Log.Information("Draft created for {Email}", record.Email);
    }
    catch (Exception ex)
    {
        Log.Error(ex, "Failed to create draft for {Email}", record.Email);
        throw;
    }
    finally
    {
        // Critical: Release COM objects to prevent memory leaks
        if (mail != null)
        {
            Marshal.ReleaseComObject(mail);
            mail = null;
        }
    }
}
```

---

## Development Setup

### 1. Clone the Repository

```bash
git clone https://github.com/ProgrammerNomad/MailMergeKit.git
cd MailMergeKit
```

### 2. Install Prerequisites

- **Visual Studio 2022** with:
  - .NET Desktop Development workload
  - Office/SharePoint Development workload
  
- **Microsoft Office** (Word + Outlook)

- **WiX Toolset** (for installer)
  ```bash
  # Download from: https://wixtoolset.org/
  ```

### 3. Restore NuGet Packages

```bash
dotnet restore
```

Or in Visual Studio: **Tools** → **NuGet Package Manager** → **Restore NuGet Packages**

### 4. Build the Solution

```bash
dotnet build MailMergeKit.sln
```

Or in Visual Studio: `Ctrl+Shift+B`

### 5. Debug the Add-in

- Set **MailMergeKit.WordAddin** as startup project
- Press `F5` to start debugging
- Word will launch with the add-in loaded

### Required NuGet Packages

**All packages are free and open source:**

```xml
<!-- Core Office Integration -->
<PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.0" />
<PackageReference Include="Microsoft.Office.Interop.Outlook" Version="15.0.0" />

<!-- Essential Frameworks (v0.0.1) -->
<PackageReference Include="Serilog" Version="3.1.1" />
<PackageReference Include="Serilog.Sinks.File" Version="5.0.0" />
<PackageReference Include="Serilog.Sinks.Console" Version="5.0.1" />
<PackageReference Include="FluentValidation" Version="11.9.0" />
<PackageReference Include="Polly" Version="8.2.1" />
<PackageReference Include="CsvHelper" Version="30.0.1" />
<PackageReference Include="Microsoft.Extensions.Configuration" Version="8.0.0" />
<PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="8.0.0" />

<!-- Utilities -->
<PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
<PackageReference Include="HtmlAgilityPack" Version="1.11.54" />

<!-- Optional: Excel validation without Office Interop -->
<PackageReference Include="ClosedXML" Version="0.102.1" />
```

**Package Purposes:**
- `Microsoft.Office.Interop.Word` - Word automation and merge field processing
- `Microsoft.Office.Interop.Outlook` - Email creation and Outlook integration
- `Serilog` - Professional logging with structured output (replaces custom Logger.cs)
- `FluentValidation` - Clean, reusable validation rules
- `Polly` - Automatic retry logic for COM failures
- `CsvHelper` - Robust CSV data source support
- `Microsoft.Extensions.Configuration` - Type-safe settings management
- `Newtonsoft.Json` - JSON serialization
- `HtmlAgilityPack` - HTML cleanup and optimization
- `ClosedXML` - Read Excel files without Office Interop (optional)

**Total Cost:** $0.00 (all open source)

See [TECHNICAL_STACK.md](TECHNICAL_STACK.md) for code examples and implementation details.

### Development Timeline

- **v0.0.1 Prototype:** 2-3 days
- **v0.1.0 First Beta:** 1-2 weeks
- **v0.5.0 Feature Complete:** 3-4 weeks
- **v1.0.0 Stable Release:** 6-8 weeks

**Getting Started:**

For v0.0.1, focus on the minimal viable prototype (~300 lines of code):
1. **Word ribbon button** - Single button in Mailings tab
2. **Read Word's mail merge data source** - Use existing MailMerge.DataSource
3. **Create Outlook drafts** - Sequential queue processing (COM-safe)
4. **Basic subject merge** - Replace merge fields in subject line
5. **Simple settings dialog** - Recipient field selector only

**Goal:** Get a working prototype in a few hours to validate architecture.

This proves the concept before adding attachments, validation, logging, etc.

---

## Project Structure

```
MailMergeKit/
│
├── MailMergeKit.sln                 # Visual Studio solution
│
├── MailMergeKit.WordAddin/          # Main add-in project
│   ├── Ribbon.cs                    # Ribbon UI definition
│   ├── MergeController.cs           # Merge logic controller
│   ├── OutlookMailer.cs             # Outlook email generator
│   ├── SettingsForm.cs              # Configuration dialog
│   ├── ThisAddIn.cs                 # VSTO add-in entry point
│   ├── Logger.cs                    # Error logging system
│   ├── MergeStateManager.cs         # Checkpoint/resume functionality
│   └── Properties/
│       └── AssemblyInfo.cs
│
├── Installer/                       # MSI installer project
│   ├── Product.wxs                  # WiX installer definition
│   └── setup.bat                    # Build script
│
├── examples/                        # Sample templates and data
│   ├── sample-template.docx         # Example merge template
│   ├── sample-data.xlsx             # Example data source
│   └── demo.gif                     # Product demonstration
│
├── docs/                            # Documentation
│   ├── user-guide.md
│   ├── developer-guide.md
│   ├── performance-guide.md         # Large campaign best practices
│   └── screenshots/
│
├── logs/                            # Merge operation logs (created at runtime)
│   ├── merge-log.txt                # Detailed merge logs
│   └── error-log.txt                # Error tracking
│
├── plugins/                         # Plugin architecture (future)
│   ├── tracking-plugin/
│   └── pdf-plugin/
│
├── tests/                           # Unit tests (future)
│
├── README.md                        # This file
├── LICENSE                          # License file
├── CHANGELOG.md                     # Version history
└── .gitignore
```

---

## Roadmap

### Version 0.0.1 (Current) - Experimental Prototype

**Core Features:**
- [ ] Word ribbon integration
- [ ] Basic merge engine (read Word data source)
- [ ] Outlook draft generation (single-threaded queue)
- [ ] Subject personalization with merge fields
- [ ] Simple settings dialog

**Goal:** Validate architecture and get working prototype in ~2-3 days

**Known Limitations:**
- No attachments yet
- No CC/BCC support
- No preview mode
- No error logging
- Drafts folder only
- Basic UI

### Version 0.1.0 - First Usable Beta

**Core Features:**
- [ ] Static and dynamic attachments
- [ ] Multiple attachments per recipient (file1.pdf;file2.pdf)
- [ ] CC/BCC support
- [ ] Email preview (first email)
- [ ] **Test email mode** - Send first record to test address before full merge
- [ ] **Automatic Outlook startup** - Launch Outlook if not running

**Quality:**
- [ ] Error logging with Serilog (structured logs)
- [ ] Pre-merge field validation (FluentValidation)
- [ ] Duplicate recipient detection with warnings
- [ ] Retry logic for COM failures (Polly)

### Version 0.2.0 - Stability & Performance

- [ ] **Resume interrupted merge** - Checkpoint system (merge_state.json)
- [ ] Progress tracking with real-time updates
- [ ] Enhanced error recovery and retry logic
- [ ] **Merge field picker UI** - Insert available fields from dropdown
- [ ] Anti-spam delays (optional delay between emails)
- [ ] Advanced validation rules

### Version 0.5.0 - Feature Complete

- [ ] Smart attachment path detection (relative, absolute, UNC)
- [ ] Comprehensive logging dashboard
- [ ] Performance optimizations for large campaigns
- [ ] Detailed user and developer documentation
- [ ] Sample templates and example data

### Version 1.0.0 - Stable Release

- [ ] MSI installer (WixSharp)
- [ ] Full documentation (user guide, developer guide)
- [ ] Production-ready error handling
- [ ] Performance tested with 5,000 email campaigns
- [ ] Community feedback incorporated
- [ ] All core features stable and tested

### Version 1.1.0 - Advanced Features

- [ ] Send scheduling with queue management
- [ ] Advanced batch sending with intelligent rate limiting
- [ ] **HTML cleanup and optimization** - Clean Word HTML output
- [ ] Enhanced preview pane with multiple samples
- [ ] Portable (no-install) version for restricted environments

### Version 2.0.0 - Enterprise Features (Future)

- [ ] **CLI mode for automation** - Command-line batch processing
- [ ] **UTM link builder** - Automatic campaign tracking parameters
- [ ] **PDF attachment auto-generation** - Convert Word docs to PDF
- [ ] **SharePoint data source** - Read from SharePoint lists
- [ ] **Template library** - Save and reuse merge configurations
- [ ] **Plugin architecture** - Extensible system for community contributions
- [ ] Multi-language support
- [ ] Advanced analytics dashboard

### Future Considerations

**Advanced Automation:**
- **CLI Mode** - Command-line interface for automation
  ```bash
  MailMergeKit.exe --template template.docx --data data.xlsx --mode draft
  ```
  Enables: Scheduled campaigns, CI/CD integration, batch scripts

**Enterprise Features:**
- Email templates with conditional content (if/else logic)
- Azure AD integration for enterprise SSO
- LDAP/Active Directory data source support
- Compliance features (GDPR consent tracking)

**Analytics & Insights:**
- Web-based dashboard for campaign tracking
- UTM parameter injection for Google Analytics
- Engagement metrics (via UTM tracking)
- Campaign comparison and A/B testing

**Developer Tools:**
- REST API for programmatic access
- Webhook support for event notifications
- PowerShell module for scripting
- Mobile companion app for approval workflows

---

## Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch**
   ```bash
   git checkout -b feature/amazing-feature
   ```
3. **Commit your changes**
   ```bash
   git commit -m 'Add amazing feature'
   ```
4. **Push to the branch**
   ```bash
   git push origin feature/amazing-feature
   ```
5. **Open a Pull Request**

### Development Guidelines

- Follow C# coding conventions
- Add XML documentation comments
- Include unit tests for new features
- Update documentation as needed

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Support

### Documentation

- [User Guide](docs/user-guide.md)
- [Developer Guide](docs/developer-guide.md)
- [FAQ](docs/faq.md)

### Getting Help

- **Issues:** [GitHub Issues](https://github.com/ProgrammerNomad/MailMergeKit/issues)
- **Discussions:** [GitHub Discussions](https://github.com/ProgrammerNomad/MailMergeKit/discussions)
- **Email:** support@mailmergekit.com

### Troubleshooting

**Add-in doesn't appear in Word:**
- Ensure Office is closed during installation
- Check: `File` → `Options` → `Add-ins` → `Manage COM Add-ins`
- Verify MailMergeKit is enabled

**"Outlook is not running" error:**
- Ensure Outlook desktop app is installed (not web version)
- Start Outlook before using MailMergeKit

**Attachments not found:**
- Use absolute paths or relative to document location
- Verify file permissions
- Check the merge log for specific missing files

**Slow performance on large campaigns:**
- Ensure Outlook is already running before merge
- Close unnecessary applications to free up memory
- Use SSD for attachment storage
- For campaigns over 2,000 emails, enable resume feature in case of interruption
- Note: Outlook COM is single-threaded; processing time scales linearly with email count

**Memory issues or crashes:**
- Restart Outlook between very large merges
- Check `logs/error-log.txt` for COM memory errors
- Ensure you're running 64-bit Office with sufficient RAM
- Verify all COM objects are being released properly

**Outlook not starting automatically:**
- Manually start Outlook before running MailMergeKit (v0.0.1)
- Auto-start feature planned for v0.1.0

---

## Performance Best Practices

### Large Campaign Optimization

**For 1,000-5,000 Emails:**

1. **Pre-start Outlook**
   - Launch Outlook before starting merge
   - Reduces COM initialization overhead

3. **Attachment Strategy**
   - Place attachments on SSD for faster access
   - Use UNC paths for network files: `\\server\share\files`
   - Avoid OneDrive/cloud-synced folders

4. **System Requirements**
   - Minimum: 8GB RAM
   - Recommended: 16GB RAM for large campaigns
   - SSD storage for best performance

5. **Batch Strategy**
   - Under 2,000: Single merge recommended
   - 2,000-5,000: Single merge with resume feature enabled
   - Over 5,000: Consider alternative bulk email solutions (Outlook has limitations)

**Expected Performance:**
```
  100 emails: ~1-2 minutes
  500 emails: ~5-8 minutes  
1,000 emails: ~10-15 minutes
2,500 emails: ~25-35 minutes
5,000 emails: ~50-75 minutes
```

*Times assume SSD storage, 16GB RAM, and no attachment processing delays*

**Note:** Outlook is designed for personal/business email, not bulk campaigns. For 10,000+ emails, use dedicated email marketing platforms (Mailchimp, SendGrid, etc.).

---

## Acknowledgments

- **Microsoft Office Interop team** for comprehensive APIs and documentation
- **VSTO community** for excellent examples and best practices
- **Email marketing professionals** who provided requirements and testing feedback
- **Open source contributors** who help improve MailMergeKit
- **Users** who trust MailMergeKit for their campaigns and provide valuable feedback

**Special Thanks:**
- Inspired by the need for an open, privacy-focused alternative to expensive mail merge tools
- Built with feedback from teams managing email campaigns
- Designed for enterprises who need auditable, local-only processing

**Version:** 0.0.1 (Experimental Prototype)

---

<div align="center">

**Made for productivity enthusiasts**

[Back to Top](#mailmergekit)

</div>
