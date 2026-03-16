# MailMergeKit Examples

This folder contains sample files to help you get started with MailMergeKit.

## Files

### sample-data.csv
A simple CSV file with sample recipient data demonstrating basic mail merge fields.

**Fields:**
- `Email` - Recipient email address (required)
- `FirstName` - Recipient first name
- `LastName` - Recipient last name
- `Company` - Company name
- `Domain` - Domain name (for domain expiry example)
- `ExpiryDate` - Expiration date example

### sample-template.txt
A template showing how to structure your Word document for mail merge.

## How to Use

1. **Create a new Word document**

2. **Set up mail merge data source:**
   - Open Word
   - Go to **Mailings** tab
   - Click **Select Recipients** → **Use Existing List**
   - Choose `sample-data.csv`

3. **Insert merge fields in your document:**
   - Place cursor where you want a field
   - Click **Mailings** → **Insert Merge Field**
   - Select the field (e.g., FirstName, Domain)

4. **Example template:**
   ```
   Dear «FirstName» «LastName»,

   This is a reminder that your domain «Domain» will expire on «ExpiryDate».

   Please renew your domain to avoid service interruption.

   Best regards,
   «Company» Support Team
   ```

5. **Configure subject in MailMergeKit:**
   - Click **Send via MailMergeKit**
   - Enter subject: `Domain «Domain» expires on «ExpiryDate»`
   - Select recipient field: `Email`
   - Click **Start Merge**

6. **Review drafts in Outlook:**
   - Open Outlook
   - Go to **Drafts** folder
   - Review each email
   - Click **Send** when ready

## Advanced Examples (v0.1.0+)

### With Attachments

Add an `Attachment` column to your CSV:
```csv
Email,FirstName,Domain,Attachment
john@example.com,John,example.com,invoices/invoice_001.pdf
jane@company.com,Jane,company.com,invoices/invoice_002.pdf
```

### Multiple Attachments

Separate multiple files with semicolon:
```csv
Email,FirstName,Attachment
john@example.com,John,invoice.pdf;receipt.pdf;terms.pdf
```

### With CC/BCC

Add CC and BCC columns:
```csv
Email,FirstName,CC,BCC
john@example.com,John,manager@example.com,admin@example.com
```

## Tips

1. **Test first** - Start with 2-3 records to test your template
2. **Valid emails** - Ensure all email addresses are valid
3. **Field names** - Use simple field names without spaces or special characters
4. **UTF-8 encoding** - Save CSV files as UTF-8 for special characters
5. **Merge field syntax** - Use «FieldName» (Word's merge field format)

## Troubleshooting

**Merge fields not appearing:**
- Make sure you've set up the data source first
- Field names in CSV must match exactly (case-sensitive)

**Drafts not created:**
- Check if Outlook is running
- Verify email addresses are valid
- Check console for error messages (v0.0.1)

**Subject not personalized:**
- Use «FieldName» syntax (not {FieldName} or <<FieldName>>)
- Field names must match your data source columns

For more help, see the main [README.md](../README.md)
