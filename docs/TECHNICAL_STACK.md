# MailMergeKit Technical Stack Plan

> **All Free & Open Source** - Zero paid licenses required

## Executive Summary

This document outlines the recommended frameworks and libraries for MailMergeKit development. **Every tool listed is 100% free for commercial use** with no licensing costs.

**Estimated Code Reduction:** ~40-50% fewer manual lines compared to raw implementations

---

## Recommended Stack by Priority

### Tier 1: Essential (Implement in v0.0.1)

These provide maximum value with minimal learning curve:

#### 1. Serilog - Professional Logging
**License:** Apache 2.0 (Free)  
**NuGet:** `Serilog`, `Serilog.Sinks.File`, `Serilog.Sinks.Console`

**Benefits:**
- Replaces custom `Logger.cs` (~50 lines saved)
- Structured logging with JSON support
- Automatic timestamps and log levels
- Multiple output targets (file, console, debug)

**Before (Manual):**
```csharp
// Custom Logger.cs
public static class Logger
{
    public static void Log(string message)
    {
        File.AppendAllText("logs/merge-log.txt", 
            $"{DateTime.Now} | {message}\n");
    }
}
```

**After (Serilog):**
```csharp
// Setup once in ThisAddIn.cs
Log.Logger = new LoggerConfiguration()
    .WriteTo.File("logs/merge-log-.txt", rollingInterval: RollingInterval.Day)
    .WriteTo.Console()
    .CreateLogger();

// Usage everywhere
Log.Information("Processing {Email} | Status: {Status}", email, "SUCCESS");
Log.Error(ex, "Failed to create draft for {Email}", email);
```

**Lines Saved:** ~50-70 lines of custom logging code

---

#### 2. FluentValidation - Data Validation
**License:** Apache 2.0 (Free)  
**NuGet:** `FluentValidation`

**Benefits:**
- Clean, reusable validation rules
- Replaces messy if/else validation (~100+ lines saved)
- Better error messages
- Easy to test

**Before (Manual):**
```csharp
public bool ValidateRecipient(RecipientData data)
{
    if (string.IsNullOrEmpty(data.Email))
        throw new Exception("Email required");
    
    if (!data.Email.Contains("@"))
        throw new Exception("Invalid email");
    
    if (!string.IsNullOrEmpty(data.Attachment))
    {
        if (!File.Exists(data.Attachment))
            throw new Exception($"Attachment not found: {data.Attachment}");
    }
    // ...more validation
}
```

**After (FluentValidation):**
```csharp
public class RecipientValidator : AbstractValidator<RecipientData>
{
    public RecipientValidator()
    {
        RuleFor(x => x.Email)
            .NotEmpty().WithMessage("Email is required")
            .EmailAddress().WithMessage("Invalid email format");
        
        RuleFor(x => x.Attachment)
            .Must(File.Exists)
            .When(x => !string.IsNullOrEmpty(x.Attachment))
            .WithMessage("Attachment file not found: {PropertyValue}");
        
        RuleFor(x => x.CC)
            .Must(BeValidEmailList)
            .When(x => !string.IsNullOrEmpty(x.CC))
            .WithMessage("CC contains invalid email addresses");
    }
    
    private bool BeValidEmailList(string emails)
    {
        return emails.Split(';').All(e => IsValidEmail(e.Trim()));
    }
}

// Usage
var validator = new RecipientValidator();
var result = validator.Validate(recipient);
if (!result.IsValid)
{
    foreach (var error in result.Errors)
        Log.Error("Validation failed: {Error}", error.ErrorMessage);
}
```

**Lines Saved:** ~100-150 lines of manual validation

---

#### 3. Polly - Resilience & Retry Logic
**License:** BSD 3-Clause (Free)  
**NuGet:** `Polly`

**Benefits:**
- Automatic retry for COM failures
- Circuit breaker for repeated failures
- Reduces error handling boilerplate (~50 lines saved)

**Before (Manual):**
```csharp
public void CreateDraft(RecipientData data)
{
    int retries = 0;
    Exception lastException = null;
    
    while (retries < 3)
    {
        try
        {
            var mail = outlook.CreateItem(OlItemType.olMailItem);
            mail.To = data.Email;
            mail.Save();
            return;
        }
        catch (COMException ex)
        {
            lastException = ex;
            retries++;
            Thread.Sleep(2000);
        }
    }
    
    throw new Exception("Failed after 3 retries", lastException);
}
```

**After (Polly):**
```csharp
private static readonly AsyncPolicy _retryPolicy = Policy
    .Handle<COMException>()
    .WaitAndRetryAsync(3, 
        attempt => TimeSpan.FromSeconds(Math.Pow(2, attempt)),
        onRetry: (ex, timeSpan, retryCount, context) =>
        {
            Log.Warning("COM retry {RetryCount} after {Delay}s", 
                retryCount, timeSpan.TotalSeconds);
        });

public async Task CreateDraft(RecipientData data)
{
    await _retryPolicy.ExecuteAsync(() => 
    {
        var mail = outlook.CreateItem(OlItemType.olMailItem);
        mail.To = data.Email;
        mail.Save();
        return Task.CompletedTask;
    });
}
```

**Lines Saved:** ~50-80 lines of retry logic

---

#### 4. CsvHelper - CSV Data Source Support
**License:** Apache 2.0 / MS-PL (Free)  
**NuGet:** `CsvHelper`

**Benefits:**
- Robust CSV reading (handles quoted fields, line breaks, etc.)
- Type-safe mapping to objects
- Better than manual string.Split() parsing

**Before (Manual):**
```csharp
var lines = File.ReadAllLines("data.csv");
var headers = lines[0].Split(',');
foreach (var line in lines.Skip(1))
{
    var values = line.Split(',');
    // Error-prone, doesn't handle quotes, commas in fields, etc.
}
```

**After (CsvHelper):**
```csharp
using (var reader = new StreamReader("data.csv"))
using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
{
    var records = csv.GetRecords<RecipientData>();
    foreach (var record in records)
    {
        // Strongly-typed access
        ProcessRecipient(record);
    }
}
```

**Lines Saved:** ~30-50 lines of CSV parsing

---

#### 5. Microsoft.Extensions.Configuration
**License:** MIT (Free)  
**NuGet:** `Microsoft.Extensions.Configuration`, `Microsoft.Extensions.Configuration.Json`

**Benefits:**
- Replaces manual JSON config handling
- Type-safe settings objects
- Supports multiple config sources (JSON, XML, environment)

**Before (Manual with Newtonsoft.Json):**
```csharp
public class ConfigManager
{
    public static AppSettings Load()
    {
        var json = File.ReadAllText("config.json");
        return JsonConvert.DeserializeObject<AppSettings>(json);
    }
    
    public static void Save(AppSettings settings)
    {
        var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
        File.WriteAllText("config.json", json);
    }
}
```

**After (Microsoft.Extensions.Configuration):**
```csharp
var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false)
    .Build();

var settings = config.GetSection("MailMerge").Get<MergeSettings>();
```

**Lines Saved:** ~20-40 lines of config management

---

### Tier 2: High Value (Implement in v0.1.0)

#### 6. ClosedXML - Excel Without Office Interop
**License:** MIT (Free)  
**NuGet:** `ClosedXML`

**Benefits:**
- Read/write Excel files without Office installed
- Faster than Office Interop for data reading
- No COM cleanup needed
- Great for validation and preview

**Important Note:** EPPlus v5+ requires commercial license. **Use ClosedXML instead** - fully free!

**Usage:**
```csharp
using (var workbook = new XLWorkbook("data.xlsx"))
{
    var worksheet = workbook.Worksheet(1);
    var range = worksheet.RangeUsed();
    
    foreach (var row in range.RowsUsed().Skip(1)) // Skip header
    {
        var email = row.Cell(1).Value.ToString();
        var domain = row.Cell(2).Value.ToString();
        // Process...
    }
}
```

**Use Case:** Pre-validate Excel data source before merge, or read data without Office Interop for CLI mode.

**Lines Saved:** ~60-100 lines when used instead of Office Interop

---

#### 7. WixSharp - C# Installer Definition
**License:** MIT (Free)  
**NuGet:** `WixSharp`

**Benefits:**
- Write installer in C# instead of XML
- Less verbose than raw WiX
- Type-safe, IntelliSense support

**Before (WiX XML):**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="MailMergeKit" Language="1033" 
           Version="0.0.1" Manufacturer="YourCompany">
    <Package InstallerVersion="200" Compressed="yes" />
    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLFOLDER" Name="MailMergeKit">
          <Component Id="MainExecutable">
            <File Id="AddInDll" Source="MailMergeKit.dll" />
            <!-- 100+ more lines of XML -->
          </Component>
        </Directory>
      </Directory>
    </Directory>
  </Product>
</Wix>
```

**After (WixSharp C#):**
```csharp
var project = new Project("MailMergeKit",
    new Dir(@"%ProgramFiles%\MailMergeKit",
        new File(@"bin\Release\MailMergeKit.dll"),
        new File(@"bin\Release\MailMergeKit.vsto")))
{
    GUID = new Guid("YOUR-GUID-HERE"),
    Version = new Version("0.0.1"),
    ControlPanelInfo = new ProductInfo
    {
        Contact = "support@mailmergekit.com",
        Manufacturer = "MailMergeKit",
        ProductIcon = "icon.ico"
    }
};

project.BuildMsi();
```

**Lines Saved:** ~200-300 lines of XML converted to ~50 lines C#

---

#### 8. AngleSharp - Modern HTML Processing
**License:** MIT (Free)  
**NuGet:** `AngleSharp`

**Benefits:**
- Modern HTML parser (better API than HtmlAgilityPack)
- CSS selector support
- Clean Word HTML output

**Usage:**
```csharp
var context = BrowsingContext.New(Configuration.Default);
var document = await context.OpenAsync(req => req.Content(htmlContent));

// Remove Word-specific tags
var wordTags = document.QuerySelectorAll("[class^='Mso']");
foreach (var tag in wordTags)
    tag.Remove();

// Clean inline styles
var styledElements = document.QuerySelectorAll("[style]");
foreach (var element in styledElements)
    element.RemoveAttribute("style");

var cleanHtml = document.Body.InnerHtml;
```

**When to Use:** v0.2.0+ for HTML cleanup feature

---

#### 9. Handlebars.Net - Template Engine
**License:** MIT (Free)  
**NuGet:** `Handlebars.Net`

**Benefits:**
- Better template syntax than string replacement
- Conditional logic in templates
- Reusable helpers

**Usage:**
```csharp
// Template
var template = Handlebars.Compile("Domain {{Domain}} expires on {{ExpiryDate}}");

// Render
var result = template(new 
{ 
    Domain = "example.com", 
    ExpiryDate = "2026-04-15" 
});
```

**When to Use:** v2.0+ for advanced template features

---

### Tier 3: Nice to Have (v0.5.0+)

#### 10. MediatR - CQRS Pattern
**License:** Apache 2.0 (Free)  
**NuGet:** `MediatR`

**Benefits:**
- Organize code by commands/queries
- Reduces coupling
- Better for larger projects

**When to Use:** v1.0+ when codebase grows complex

---

#### 11. xUnit + Moq - Testing
**License:** Apache 2.0 (Free)  
**NuGet:** `xunit`, `Moq`

**Benefits:**
- Modern testing framework
- Mock Office Interop interfaces
- Test without COM dependencies

**Usage:**
```csharp
[Fact]
public void MergeEngine_ValidatesEmail()
{
    // Arrange
    var mockOutlook = new Mock<Outlook.Application>();
    var engine = new MergeEngine(mockOutlook.Object);
    
    // Act
    var result = engine.ProcessRecipient(new RecipientData 
    { 
        Email = "invalid" 
    });
    
    // Assert
    Assert.False(result.Success);
}
```

**When to Use:** v0.1.0+ for quality assurance

---

## NOT Recommended (Paid/Overkill)

### ❌ Add-in Express
- **Cost:** $349+/developer
- **Why Skip:** VSTO is free and sufficient
- **Verdict:** Unnecessary expense

### ❌ EPPlus v5+
- **Cost:** $499+ for commercial use
- **Why Skip:** ClosedXML is free and equivalent
- **Verdict:** Use ClosedXML instead

### ❌ DevExpress / Telerik
- **Cost:** $1000+/year
- **Why Skip:** WinForms built-in controls are adequate
- **Verdict:** Overkill for simple dialog

---

## Recommended NuGet Packages for v0.0.1

```xml
<!-- TIER 1: Essential -->
<PackageReference Include="Serilog" Version="3.1.1" />
<PackageReference Include="Serilog.Sinks.File" Version="5.0.0" />
<PackageReference Include="Serilog.Sinks.Console" Version="5.0.1" />
<PackageReference Include="FluentValidation" Version="11.9.0" />
<PackageReference Include="Polly" Version="8.2.1" />
<PackageReference Include="CsvHelper" Version="30.0.1" />
<PackageReference Include="Microsoft.Extensions.Configuration" Version="8.0.0" />
<PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="8.0.0" />

<!-- EXISTING -->
<PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.0" />
<PackageReference Include="Microsoft.Office.Interop.Outlook" Version="15.0.0" />
<PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
<PackageReference Include="HtmlAgilityPack" Version="1.11.54" />
```

**Total Added Packages:** 8  
**Total Cost:** $0.00 (All free & open source)

---

## Implementation Plan

### Phase 1: v0.0.1 Prototype (2-3 days)
1. Install Serilog → Replace Logger.cs
2. Install FluentValidation → Add RecipientValidator.cs
3. Install Polly → Add retry logic to OutlookMailer.cs
4. Keep existing: Office Interop, Newtonsoft.Json

**Estimated Development Time:** Same as manual (~2-3 days)  
**Code Quality:** Much higher  
**Maintenance:** Much easier

### Phase 2: v0.1.0 Beta (1 week)
5. Add CsvHelper → Support CSV data sources
6. Add Microsoft.Extensions.Configuration → Better config management
7. Add ClosedXML → Pre-validate Excel files
8. Add WixSharp → Easier installer maintenance

### Phase 3: v0.5.0+ (As needed)
9. Add AngleSharp → HTML cleanup
10. Add xUnit + Moq → Unit testing
11. Consider MediatR if complexity grows

---

## Cost Comparison

| Approach | Libraries Cost | Development Time | Maintenance |
|----------|---------------|------------------|-------------|
| **Manual Code** | $0 | 6-8 weeks | High (bugs, boilerplate) |
| **Free Frameworks** | $0 | 4-6 weeks | Low (tested libraries) |
| **Paid Tools** | $1000-2000+ | 3-5 weeks | Low (vendor lock-in) |

**Recommendation:** Use free frameworks - same cost as manual but 30% faster and more maintainable.

---

## Architecture with Frameworks

```
MailMergeKit v0.0.1 with Free Frameworks
├── Core
│   ├── .NET Framework 4.8
│   ├── VSTO (Word/Outlook add-in)
│   └── C# 7.3+
├── Logging
│   └── Serilog (replaces custom Logger.cs)
├── Validation
│   └── FluentValidation (replaces manual checks)
├── Resilience
│   └── Polly (auto-retry COM failures)
├── Data Access
│   ├── Microsoft.Office.Interop.Word (mail merge data)
│   ├── CsvHelper (CSV support)
│   └── ClosedXML (optional: Excel validation)
├── Configuration
│   └── Microsoft.Extensions.Configuration (replaces manual JSON)
├── Email Generation
│   └── Microsoft.Office.Interop.Outlook (COM)
├── Installer
│   └── WixSharp (C# instead of XML)
└── Testing (v0.1.0+)
    ├── xUnit
    └── Moq
```

---

## Lines of Code Estimate

| Component | Manual Code | With Frameworks | Savings |
|-----------|-------------|-----------------|---------|
| Logger.cs | ~80 lines | ~10 lines (Serilog setup) | 70 lines |
| Validation | ~150 lines | ~30 lines (FluentValidation) | 120 lines |
| Retry Logic | ~80 lines | ~15 lines (Polly) | 65 lines |
| CSV Parsing | ~60 lines | ~20 lines (CsvHelper) | 40 lines |
| Config Management | ~50 lines | ~15 lines (MS.Extensions) | 35 lines |
| Installer XML | ~300 lines | ~60 lines (WixSharp) | 240 lines |
| **TOTAL** | **~720 lines** | **~150 lines** | **~570 lines (79%)** |

**Additional Benefits:**
- Better error messages
- More robust error handling
- Easier to test
- Industry-standard patterns
- Active community support

---

## Final Recommendation

### Use These (100% Free):
1. ✅ **Serilog** - Professional logging
2. ✅ **FluentValidation** - Clean validation
3. ✅ **Polly** - Automatic retry
4. ✅ **CsvHelper** - CSV support
5. ✅ **Microsoft.Extensions.Configuration** - Settings management
6. ✅ **ClosedXML** - Excel without Interop (optional)
7. ✅ **WixSharp** - Easier installer (v0.1.0+)

### Skip These:
- ❌ EPPlus v5+ (use ClosedXML)
- ❌ Add-in Express (VSTO is sufficient)
- ❌ Any paid UI frameworks (WinForms is adequate)

### Total Cost: $0.00
### Development Time Saved: ~30-40%
### Code Quality: Professional-grade

---

**Next Step:** Should I update the README.md Technology Stack section to reflect these free frameworks?
