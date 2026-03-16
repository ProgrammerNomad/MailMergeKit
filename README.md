# MailMergeKit

> Inspired by the [mergelook](https://github.com/ilias-sp/mergelook) project.

MailMergeKit is a simple Outlook mail merge tool that generates personalized draft emails from an Excel spreadsheet.

It focuses on a cleaner workflow for creating personalized Outlook emails with attachments using Excel data.

MailMergeKit works entirely on your local machine using Microsoft Excel and Microsoft Outlook, without using any external servers or SMTP configuration.

## Features

- Create personalized Outlook emails from Excel
- Generate multiple draft emails automatically
- Personalize subject and message body
- Attach files per recipient
- Support for To, Cc, Bcc and Reply-To
- Use custom merge fields
- Works with your existing Outlook account

After emails are generated, the user can review them in Outlook and send manually.

## Requirements

- Microsoft Windows
- Microsoft Excel
- Microsoft Outlook

## How It Works

MailMergeKit reads recipient data from Excel and replaces variables inside an Outlook email template.

Example workflow:

```text
Excel list
   ↓
MailMergeKit macro
   ↓
Outlook draft emails generated
   ↓
User reviews and clicks Send
```

No automatic sending is performed.

## Files Included

You only need these two files:

- MailMergeKit.xlsm
- message.oft

### MailMergeKit.xlsm

Excel file containing the VBA macro that performs the mail merge.

### message.oft

Outlook email template used for generating personalized emails.

## Setup

1. Download the repository.
2. Place the following files in the same folder:
   - MailMergeKit.xlsm
   - message.oft
   - attachments
3. Open MailMergeKit.xlsm
4. Click Enable Content to allow macros.

## Excel Data Format

Example spreadsheet:

| To              | NAME | DOMAIN      | ATTACHMENT   |
|-----------------|------|-------------|-------------|
| user@test.com   | John | example.com | invoice.pdf  |

Each row represents one email.

## Template Variables

Inside the Outlook template (message.oft) you can use variables.

Example subject:

```text
Domain ___DOMAIN___ expiration notice
```

Example email body:

```text
Hello ___NAME___,

Your domain ___DOMAIN___ will expire soon.

Regards
Team
```

During mail merge these variables are replaced with values from Excel.

## Example Result

If Excel contains:

| NAME | DOMAIN      |
|------|-------------|
| John | example.com |

The generated email will be:

```text
Subject: Domain example.com expiration notice

Hello John,

Your domain example.com will expire soon.

Regards
Team
```

## Outlook Offline Mode (Recommended)

Before generating emails it is recommended to enable Work Offline in Outlook.

Steps:

1. Open Outlook
2. Go to Send / Receive
3. Click Work Offline

This prevents emails from being sent before reviewing them.

## Editing the Email Template

To modify the template:

1. Open message.oft
2. Make your changes
3. Save again as Outlook Template (.oft)
4. Replace the existing file in the project folder

Do not save it in Outlook's default template directory.

## Custom Fields

You can add your own variables.

Example Excel header:

```text
___PRODUCT___
```

Template usage:

```text
Your product ___PRODUCT___ is ready.
```

## Troubleshooting

### Attachments not found

Ensure the attachment files exist in the same folder as the Excel file.

### Macros disabled

Enable macros when opening MailMergeKit.xlsm.

### Emails not visible in Outlook

Sometimes Outlook does not refresh automatically. Switch folders and return to Outbox or Drafts.

## Disclaimer

MailMergeKit is a lightweight tool designed for simple mail merge tasks inside Outlook. It does not aim to replace full-featured commercial mail merge systems.

All operations are performed locally on your computer.

## License

Open-source project based on the original mergelook concept.
See repository license for details.
