# MailMergeKit v0.0.1 - Getting Started

This is the experimental prototype (v0.0.1) of MailMergeKit.

## What Works

- ✅ Word ribbon integration (button in Mailings tab)
- ✅ Reading Word mail merge data sources
- ✅ Creating Outlook draft emails
- ✅ Subject personalization with merge fields
- ✅ Simple settings dialog
- ✅ Sequential queue processing (COM-safe)

## How to Build

### Prerequisites

1. **Visual Studio 2022** with:
   - .NET Desktop Development workload
   - Office/SharePoint Development workload

2. **Microsoft Office** (Word + Outlook installed locally)

3. **Office Developer Tools for Visual Studio**

### Build Steps

1. Open `MailMergeKit.sln` in Visual Studio 2022

2. Restore NuGet packages (if prompted)

3. Build the solution:
   - Press `Ctrl+Shift+B`
   - Or: **Build** → **Build Solution**

4. Run in debug mode:
   - Press `F5`
   - Word will launch with MailMergeKit loaded

### Testing the Add-in

1. Word should open automatically when you press F5

2. Go to the **Mailings** tab

3. You should see a **MailMergeKit** section with a "Send via MailMergeKit" button

4. Set up a simple mail merge:
   - **Mailings** → **Select Recipients** → **Use Existing List**
   - Choose `examples/sample-data.xlsx` (create this file first)
   - Click **Send via MailMergeKit**

## Current Limitations (v0.0.1)

This is an experimental prototype. The following features are **NOT** implemented yet:

- ❌ No attachment support (planned for v0.1.0)
- ❌ No CC/BCC support (planned for v0.1.0)
- ❌ No preview mode (planned for v0.1.0)
- ❌ No structured logging (console only)
- ❌ No proper error handling
- ❌ No retry logic
- ❌ Body merge uses plain text (HTML in v0.1.0)
- ❌ No progress tracking UI
- ❌ No test email mode

## Known Issues

1. **Body merge not working correctly** - The current implementation reads Word content but doesn't properly merge fields in the body. This will be fixed in v0.1.0.

2. **COM cleanup** - Some COM objects may not be released properly, leading to memory leaks on large campaigns.

3. **Error messages** - Errors are logged to console only (no user-friendly dialogs).

4. **No validation** - Email addresses and merge fields are not validated.

## Next Steps for v0.1.0

1. Fix body HTML merge (use Word's mail merge engine properly)
2. Add attachment support (static and dynamic)
3. Add CC/BCC support
4. Add email preview
5. Add Serilog logging
6. Add FluentValidation
7. Add Polly retry logic
8. Add test email mode
9. Add auto-start Outlook if not running

## Goal of v0.0.1

This prototype validates:
- ✅ VSTO ribbon integration works
- ✅ Can read Word mail merge data source
- ✅ Can create Outlook draft emails
- ✅ Subject merge works
- ✅ Sequential processing is stable

**This proves the architecture is sound before adding advanced features.**

## Development Time

- Setup: 1 hour
- Core implementation: 2-3 hours
- Testing: 30 minutes
- **Total: ~4 hours**

## Feedback

This is a proof-of-concept. Please test and provide feedback before we proceed to v0.1.0.

Focus areas for testing:
1. Does the ribbon button appear?
2. Can you select a data source?
3. Do draft emails appear in Outlook?
4. Does subject personalization work?
5. Any crashes or errors?

Report issues to: https://github.com/ProgrammerNomad/MailMergeKit/issues
