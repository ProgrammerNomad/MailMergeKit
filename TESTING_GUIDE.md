# MailMergeKit v0.0.1 - Testing Guide

## ✅ Project Status: READY FOR TESTING

All files have been created and the project is structurally complete. Follow these steps to test the add-in.

---

## 🚀 How to Open & Build

### Quick Start (Easy Way)

1. **Double-click:** `OPEN_IN_VS.bat` in the project root
   - This will open the solution in Visual Studio 2022

### Manual Start

1. **Open Visual Studio 2022**

2. **Open the solution:**
   - File → Open → Project/Solution
   - Navigate to: `C:\xampp\htdocs\MailMergeKit\MailMergeKit.sln`
   - Click "Open"

3. **Restore NuGet packages** (if prompted):
   - Right-click solution in Solution Explorer
   - Click "Restore NuGet Packages"
   - Wait for restore to complete

4. **Build the solution:**
   - Press `Ctrl+Shift+B`
   - Or: Build → Build Solution
   - Wait for build to complete (should take 5-10 seconds)

5. **Check for errors:**
   - Look at the Output window (View → Output)
   - Should see: "Build: 1 succeeded, 0 failed"

---

## 🧪 Testing Steps

### Step 1: Launch Word with Add-in

1. **In Visual Studio, press F5** (or click Debug → Start Debugging)

2. **Word will launch automatically** (may take 10-20 seconds first time)

3. **Verify the add-in loaded:**
   - Go to the **Mailings** tab in Word
   - Look for a **MailMergeKit** section with a button
   - You should see: "Send via MailMergeKit" button

**✅ If you see the button → Add-in is working!**

**❌ If button doesn't appear:**
- File → Options → Add-ins → Manage COM Add-ins
- Check if "MailMergeKit" is listed and enabled
- If not listed, the add-in failed to register

---

### Step 2: Test Basic Mail Merge

1. **In Word, set up a simple document:**
   ```
   Dear Customer,

   This is a test email for mail merge.

   Best regards
   ```

2. **Connect the sample data source:**
   - Click **Mailings** tab
   - Click **Select Recipients** → **Use Existing List**
   - Navigate to: `C:\xampp\htdocs\MailMergeKit\examples\sample-data.csv`
   - Click "Open"

3. **Verify data loaded:**
   - You should see 5 records
   - Status bar should show: "Record 1 of 5"

---

### Step 3: Run MailMergeKit

1. **Start Outlook** (if not already running)
   - **IMPORTANT:** Outlook MUST be running before the merge

2. **Click the "Send via MailMergeKit" button** in Word

3. **Settings dialog should appear:**
   - **Recipient Email Field:** Should show dropdown with fields
   - Select "**Email**" from the dropdown
   - **Subject Template:** Enter: `Test Email for «FirstName»`
   - Click "**Start Merge**"

4. **Wait for processing:**
   - Should process 5 emails
   - Progress shows in console/debug output

5. **Check results:**
   - Success dialog should appear: "Merge completed successfully!"
   - Should show: "Success: 5, Failed: 0"

---

### Step 4: Verify Outlook Drafts

1. **Open Outlook**

2. **Go to Drafts folder:**
   - Click "Drafts" in the left navigation pane

3. **You should see 5 draft emails:**
   - Email 1: To: john@example.com, Subject: "Test Email for John"
   - Email 2: To: jane@company.com, Subject: "Test Email for Jane"
   - Email 3: To: bob@business.net, Subject: "Test Email for Bob"
   - Email 4: To: alice@enterprise.org, Subject: "Test Email for Alice"
   - Email 5: To: charlie@startup.io, Subject: "Test Email for Charlie"

4. **Open one email to verify:**
   - Subject should be personalized (e.g., "Test Email for John")
   - To field should have the correct email
   - Body should have the Word document content

**✅ If all 5 emails are in Drafts with personalized subjects → SUCCESS!**

---

## 🎯 What to Test

### Basic Functionality (v0.0.1)

| Feature | Test | Expected Result |
|---------|------|-----------------|
| **Ribbon Integration** | Look for button in Mailings tab | ✅ Button appears |
| **Settings Dialog** | Click button | ✅ Dialog opens |
| **Field Detection** | Check dropdown | ✅ Shows: Email, FirstName, LastName, Company, Domain, ExpiryDate |
| **Subject Merge** | Use `«FirstName»` in subject | ✅ Each email has different name |
| **Draft Creation** | Check Outlook Drafts | ✅ 5 emails created |
| **Sequential Processing** | Watch debug output | ✅ Processes one at a time |
| **Error Handling** | No crashes | ✅ Completes without errors |

### Known Limitations (Expected in v0.0.1)

These features are **NOT** implemented yet - don't test these:

- ❌ **Attachments** - Not implemented yet (v0.1.0)
- ❌ **CC/BCC** - Not implemented yet (v0.1.0)
- ❌ **HTML Body** - Uses plain text only (v0.1.0)
- ❌ **Preview Mode** - No preview dialog (v0.1.0)
- ❌ **Logging** - Console output only, no log files
- ❌ **Validation** - No email validation or duplicate checks

---

## 🐛 Common Issues & Solutions

### Issue #1: Button Doesn't Appear

**Problem:** MailMergeKit button not in Mailings tab

**Solutions:**
1. Check if add-in is enabled:
   - File → Options → Add-ins → Manage COM Add-ins
   - Ensure "MailMergeKit" is checked
   
2. Add Trusted Location:
   - File → Options → Trust Center → Trust Center Settings
   - Trusted Locations → Add new location
   - Browse to: `C:\xampp\htdocs\MailMergeKit\src\MailMergeKit.WordAddin\bin\Debug`
   - Check "Subfolders of this location are also trusted"
   
3. Rebuild the solution:
   - In Visual Studio: Build → Rebuild Solution
   - Press F5 again

### Issue #2: "Outlook is not running"

**Problem:** Error message when clicking Start Merge

**Solution:**
- Start Outlook BEFORE running the merge
- Make sure it's the desktop version (not Outlook.com web)

### Issue #3: Build Errors

**Problem:** Build fails with errors

**Solutions:**
1. Check Visual Studio has Office Development workload:
   - Tools → Get Tools and Features
   - Verify "Office/SharePoint development" is installed
   
2. Ensure Office is installed locally:
   - Word and Outlook must be installed on your machine
   
3. Restore NuGet packages:
   - Right-click solution → Restore NuGet Packages

### Issue #4: No Emails Created

**Problem:** Merge completes but no emails in Drafts

**Check:**
1. Verify Outlook is running
2. Check Visual Studio Output window for errors
3. Look for console messages in Debug output
4. Ensure data source has valid email addresses

### Issue #5: Subject Not Personalized

**Problem:** All emails have the same subject

**Solution:**
- Make sure you're using Word's merge field syntax: `«FieldName»`
- Not: `{FieldName}` or `<<FieldName>>` or `<FieldName>`
- Field names are case-sensitive and must match CSV columns

---

## 📊 Test Results Template

Copy this and fill in your results:

```
MailMergeKit v0.0.1 Test Results
Date: _______________
Tester: _______________

✅/❌ Visual Studio build successful
✅/❌ Word launches with F5
✅/❌ Ribbon button appears in Mailings tab
✅/❌ Settings dialog opens
✅/❌ Field dropdown shows all 6 fields
✅/❌ Data source loads (5 records)
✅/❌ Merge completes without crashes
✅/❌ 5 draft emails created in Outlook
✅/❌ Subjects are personalized (different names)
✅/❌ Email addresses are correct

Success Rate: __/10 tests passed

Issues Found:
1. _______________________________________________
2. _______________________________________________

Notes:
_______________________________________________
_______________________________________________
```

---

## 🎓 Advanced Testing (Optional)

If basic tests pass, try these:

### Test 1: Different Subject Templates

Try these subject templates:
- `Hello «FirstName» «LastName»`
- `Domain «Domain» expires on «ExpiryDate»`
- `«Company» - Important Notice`

### Test 2: Larger Data Set

1. Create a CSV with 20-50 records
2. Test merge performance
3. Verify all emails created

### Test 3: Error Handling

1. Use invalid data (missing email addresses)
2. Disconnect Outlook mid-merge
3. Use invalid field names in subject

---

## 📝 Next Steps After Testing

### If Testing Succeeds ✅

**Report:**
- "v0.0.1 works! Ready for v0.1.0 development"
- Share any observations or suggestions

**Next development:**
1. Add attachment support
2. Add CC/BCC fields
3. Add email preview
4. Add Serilog logging
5. Add FluentValidation

### If Testing Fails ❌

**Report the issue with:**
1. What you were doing (exact steps)
2. What happened (error message, screenshot)
3. What you expected to happen
4. Your environment:
   - Windows version
   - Office version (Word/Outlook)
   - Visual Studio version

**Debug steps:**
1. Check Output window in Visual Studio (View → Output)
2. Look for red error messages
3. Copy full error text
4. Check if it's a known issue in BUILD_SUMMARY.md

---

## 🆘 Getting Help

1. **Check BUILD_SUMMARY.md** - Complete troubleshooting guide
2. **Check Output window** - Visual Studio shows detailed errors
3. **Check docs/developer-guide.md** - Development setup help
4. **GitHub Issues** - Report bugs or ask questions

---

## ✨ Success Criteria

The v0.0.1 prototype is successful if:

✅ Visual Studio builds without errors
✅ Word launches with add-in loaded
✅ Ribbon button appears in Mailings tab
✅ Settings dialog opens and shows fields
✅ Can select data source (CSV works)
✅ Merge completes without crashes
✅ Draft emails appear in Outlook
✅ Subjects are personalized with merge fields
✅ No memory leaks or COM crashes
✅ Sequential processing is stable

**If 8/10 or more pass → v0.0.1 is successful!**

---

## 🎯 Your Testing Checklist

Before you start:
- [ ] Visual Studio 2022 installed with Office workload
- [ ] Word 2016+ installed locally
- [ ] Outlook 2016+ installed locally
- [ ] Outlook is running
- [ ] You have admin rights (first-time VSTO setup)

Testing:
- [ ] Open solution in Visual Studio
- [ ] Build succeeds (Ctrl+Shift+B)
- [ ] Press F5 to launch Word
- [ ] Ribbon button appears
- [ ] Settings dialog works
- [ ] Merge creates 5 draft emails
- [ ] Subjects are personalized
- [ ] No crashes or errors

---

**Ready to test? Double-click `OPEN_IN_VS.bat` to get started!**

---

*Good luck! Let me know the results.* 🚀
