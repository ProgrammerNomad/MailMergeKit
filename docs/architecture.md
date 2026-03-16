# MailMergeKit Architecture

> Clean Architecture implementation for enterprise-grade Word/Outlook add-in

## Architecture Overview

MailMergeKit follows **Clean Architecture** principles with clear separation of concerns across multiple layers.

### Architectural Principles

1. **Dependency Inversion** - Core has no dependencies on infrastructure
2. **Single Responsibility** - Each layer has one clear purpose
3. **Interface Segregation** - Small, focused interfaces
4. **Separation of Concerns** - UI, business logic, and data access are isolated

---

## Layer Architecture

```
┌─────────────────────────────────────────────────────────┐
│                   Presentation Layer                    │
│              (MailMergeKit.WordAddin)                   │
│                                                         │
│  • Ribbon UI (Word integration)                         │
│  • WinForms dialogs (Settings, Preview, Progress)      │
│  • VSTO bootstrap and initialization                   │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│                  Application Layer                      │
│             (MailMergeKit.Application)                  │
│                                                         │
│  • Controllers (MergeController)                        │
│  • Use Cases (GenerateDraftEmails, PreviewMerge)       │
│  • Application workflows                               │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│                    Core/Domain Layer                    │
│               (MailMergeKit.Core)                       │
│                                                         │
│  • Business logic (MergeEngine)                         │
│  • Domain models (RecipientData, MergeSettings)        │
│  • Interfaces (IEmailService, IMergeEngine)            │
│  • Validation rules (FluentValidation)                 │
│  • NO DEPENDENCIES on UI or Infrastructure             │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│                Infrastructure Layer                     │
│            (MailMergeKit.Infrastructure)                │
│                                                         │
│  • Outlook COM service (OutlookService)                 │
│  • Word COM service (WordMergeReader)                   │
│  • Data source readers (CSV, Excel)                     │
│  • Logging (Serilog setup)                              │
│  • File system operations                               │
└─────────────────────────────────────────────────────────┘
```

---

## Dependency Flow Rules

### ✅ Allowed Dependencies

```
UI → Application → Core
Infrastructure → Core
```

### ❌ Forbidden Dependencies

```
Core → Infrastructure  (NEVER!)
Core → Outlook COM     (NEVER!)
Core → Word COM        (NEVER!)
Core → UI              (NEVER!)
```

**Why?** Core must remain pure business logic with no external dependencies. This enables:
- Unit testing without COM
- Platform independence
- Easy mocking
- Plugin architecture

---

## Project Structure (v1.0 Target)

```
MailMergeKit/
│
├── src/
│   │
│   ├── MailMergeKit.WordAddin/        # Presentation Layer
│   │   ├── Addin/
│   │   │   ├── ThisAddIn.cs           # VSTO entry point
│   │   │   └── AddinBootstrap.cs      # DI container setup
│   │   ├── Ribbon/
│   │   │   └── MailMergeRibbon.cs     # Word ribbon UI
│   │   └── UI/
│   │       ├── SettingsForm.cs        # Configuration dialog
│   │       ├── PreviewForm.cs         # Email preview
│   │       └── ProgressForm.cs        # Merge progress UI
│   │
│   ├── MailMergeKit.Application/      # Application Layer
│   │   ├── Controllers/
│   │   │   └── MergeController.cs     # Main workflow controller
│   │   └── UseCases/
│   │       ├── GenerateDraftEmails.cs # Primary use case
│   │       └── PreviewMerge.cs        # Preview use case
│   │
│   ├── MailMergeKit.Core/             # Core/Domain Layer
│   │   ├── Models/
│   │   │   ├── RecipientData.cs       # Recipient model
│   │   │   ├── MergeSettings.cs       # Configuration model
│   │   │   └── AttachmentInfo.cs      # Attachment metadata
│   │   ├── Services/
│   │   │   ├── MergeEngine.cs         # Core merge logic
│   │   │   └── AttachmentResolver.cs  # Path resolution
│   │   ├── Validation/
│   │   │   └── RecipientValidator.cs  # FluentValidation rules
│   │   └── Interfaces/
│   │       ├── IEmailService.cs       # Email abstraction
│   │       ├── IDataSourceReader.cs   # Data source abstraction
│   │       └── IMergeEngine.cs        # Merge engine contract
│   │
│   └── MailMergeKit.Infrastructure/   # Infrastructure Layer
│       ├── Outlook/
│       │   ├── OutlookService.cs      # Outlook COM implementation
│       │   └── OutlookFactory.cs      # COM object factory
│       ├── Word/
│       │   └── WordMergeReader.cs     # Word data source reader
│       ├── DataSources/
│       │   ├── CsvReaderService.cs    # CSV reader
│       │   └── ExcelReaderService.cs  # Excel reader (ClosedXML)
│       └── Logging/
│           └── SerilogSetup.cs        # Logging configuration
│
├── installer/
│   └── MailMergeKit.Installer/        # WixSharp installer
│
├── tests/
│   └── MailMergeKit.Tests/            # Unit tests
│
├── examples/                          # Sample data and templates
└── docs/                              # Documentation
```

---

## Component Responsibilities

### 1. WordAddin (Presentation)

**Responsibility:** User interface only

**Contains:**
- Ribbon buttons
- WinForms dialogs
- VSTO integration

**Does NOT contain:**
- Business logic
- Data access
- COM operations (delegates to Infrastructure)

**Example:**
```csharp
// Ribbon.cs
private void btnGenerateDrafts_Click(object sender, RibbonControlEventArgs e)
{
    var controller = new MergeController();
    controller.GenerateDraftEmails();
}
```

### 2. Application (Workflows)

**Responsibility:** Application workflows and coordination

**Contains:**
- Controllers
- Use cases
- Workflow orchestration

**Does NOT contain:**
- UI code
- COM operations
- Pure business logic (that's in Core)

**Example:**
```csharp
// MergeController.cs
public class MergeController
{
    private readonly IMergeEngine _mergeEngine;
    private readonly IEmailService _emailService;
    
    public void GenerateDraftEmails()
    {
        var recipients = _mergeEngine.GetRecipients();
        
        foreach (var recipient in recipients)
        {
            _emailService.CreateDraft(recipient);
        }
    }
}
```

### 3. Core (Business Logic)

**Responsibility:** Pure business logic

**Contains:**
- Domain models
- Business rules
- Interfaces (contracts)
- Validation logic

**NEVER contains:**
- UI references
- Outlook/Word COM
- File I/O
- External dependencies

**Example:**
```csharp
// MergeEngine.cs (Core)
public class MergeEngine : IMergeEngine
{
    public List<RecipientData> GetRecipients(IDataSourceReader reader)
    {
        var records = reader.ReadAll();
        
        // Pure business logic - no COM!
        return records
            .Where(r => !string.IsNullOrEmpty(r.Email))
            .Select(r => new RecipientData
            {
                Email = r.Email,
                Subject = MergeSubjectFields(r),
                // ...
            })
            .ToList();
    }
}
```

### 4. Infrastructure (External Systems)

**Responsibility:** External system integration

**Contains:**
- Outlook COM wrapper
- Word COM wrapper
- File system access
- Logging implementation

**Example:**
```csharp
// OutlookService.cs (Infrastructure)
public class OutlookService : IEmailService
{
    private Outlook.Application _outlookApp;
    
    public void CreateDraft(RecipientData recipient)
    {
        var mail = (Outlook.MailItem)_outlookApp.CreateItem(
            Outlook.OlItemType.olMailItem);
        
        mail.To = recipient.Email;
        mail.Subject = recipient.Subject;
        mail.Save();
        
        Marshal.ReleaseComObject(mail);
    }
}
```

---

## Data Flow Example

### Use Case: Generate Draft Emails

```
User clicks ribbon button
        │
        ▼
MailMergeRibbon.cs (WordAddin)
        │
        │ creates
        ▼
MergeController (Application)
        │
        │ calls
        ▼
MergeEngine.GetRecipients() (Core)
        │
        │ uses
        ▼
IDataSourceReader (interface in Core, implemented in Infrastructure)
        │
        │ returns
        ▼
List<RecipientData> (Core model)
        │
        │ forEach
        ▼
IEmailService.CreateDraft() (interface in Core, implemented in Infrastructure)
        │
        │ calls
        ▼
OutlookService (Infrastructure)
        │
        │ interacts with
        ▼
Outlook COM
        │
        ▼
Draft Email Created
```

---

## Testing Strategy

### Unit Tests (Core Layer)

**Test without COM dependencies:**

```csharp
[Fact]
public void MergeEngine_RemovesInvalidEmails()
{
    // Arrange
    var mockReader = new Mock<IDataSourceReader>();
    mockReader.Setup(r => r.ReadAll()).Returns(new[]
    {
        new { Email = "valid@example.com", Name = "John" },
        new { Email = "", Name = "Invalid" }
    });
    
    var engine = new MergeEngine();
    
    // Act
    var recipients = engine.GetRecipients(mockReader.Object);
    
    // Assert
    Assert.Single(recipients);
    Assert.Equal("valid@example.com", recipients[0].Email);
}
```

### Integration Tests (Infrastructure Layer)

**Test COM integration with mocks:**

```csharp
[Fact]
public void OutlookService_CreatesMailItem()
{
    var mockOutlook = new Mock<Outlook.Application>();
    var mockMailItem = new Mock<Outlook.MailItem>();
    
    mockOutlook.Setup(o => o.CreateItem(It.IsAny<OlItemType>()))
               .Returns(mockMailItem.Object);
    
    var service = new OutlookService(mockOutlook.Object);
    
    service.CreateDraft(new RecipientData 
    { 
        Email = "test@example.com" 
    });
    
    mockMailItem.Verify(m => m.Save(), Times.Once);
}
```

---

## Versioning Strategy

### v0.0.1 - v0.1.0: Monolithic Structure
- Single `MailMergeKit.WordAddin` project
- All code in one assembly
- Organized by folders (Ribbon/, Services/, Models/)

### v0.2.0 - v0.5.0: Transition Phase
- Introduce interfaces
- Separate concerns within single project
- Prepare for multi-project split

### v1.0.0: Full Clean Architecture
- 4 projects (WordAddin, Core, Infrastructure, Application)
- Complete separation of concerns
- Full unit test coverage
- Plugin-ready architecture

---

## Plugin Architecture (v2.0+)

With Clean Architecture, plugins can:

1. Implement `IEmailService` for alternative email providers (e.g., SendGrid)
2. Implement `IDataSourceReader` for custom data sources (e.g., CRM APIs)
3. Extend `IMergeEngine` for custom merge logic

**Example Plugin:**
```csharp
// Plugin project
public class SendGridService : IEmailService
{
    public void CreateDraft(RecipientData recipient)
    {
        // Send via SendGrid API instead of Outlook
    }
}
```

Register via DI:
```csharp
container.RegisterType<IEmailService, SendGridService>();
```

---

## Best Practices

### 1. Keep Core Pure
- No COM references
- No UI references
- Only business logic and models

### 2. Use Interfaces
- Define contracts in Core
- Implement in Infrastructure
- Easy to mock for testing

### 3. Dependency Injection
- Use constructor injection
- Register dependencies in `AddinBootstrap.cs`
- Makes testing easier

### 4. COM Safety
- All COM operations in Infrastructure only
- Proper cleanup with `Marshal.ReleaseComObject`
- Single-threaded queue processing

---

## Migration Path

**From v0.0.1 → v1.0.0:**

1. **v0.1.0** - Introduce interfaces and folder organization
2. **v0.2.0** - Extract Models and Interfaces to separate folders
3. **v0.5.0** - Create `Core` and `Infrastructure` projects
4. **v1.0.0** - Full 4-layer architecture with complete separation

**Time estimate:** 6-8 weeks gradual refactoring alongside feature development

---

## References

- [Clean Architecture by Robert C. Martin](https://blog.cleancoder.com/uncle-bob/2012/08/13/the-clean-architecture.html)
- [SOLID Principles](https://en.wikipedia.org/wiki/SOLID)
- [Domain-Driven Design](https://martinfowler.com/bliki/DomainDrivenDesign.html)
- [VSTO Best Practices](https://docs.microsoft.com/en-us/visualstudio/vsto/)

---

**Version:** Target for v1.0.0 (Full implementation)  
**Current:** v0.0.1 (Monolithic prototype)
