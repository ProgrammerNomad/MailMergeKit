# MailMergeKit Developer Guide

> Guide for contributors and developers

## Table of Contents

- [Development Environment](#development-environment)
- [Project Structure](#project-structure)
- [Building from Source](#building-from-source)
- [Running Tests](#running-tests)
- [Contributing](#contributing)
- [Code Style](#code-style)

---

## Development Environment

### Required Tools

- **Visual Studio 2022** (Community, Professional, or Enterprise)
- **Workloads:**
  - .NET Desktop Development
  - Office/SharePoint Development
- **Microsoft Office** (Word + Outlook desktop)
- **Git** for version control

### Optional Tools

- **WiX Toolset** v3.11+ (for installer development)
- **ReSharper** or **Rider** (code quality)

---

## Project Structure

See [architecture.md](architecture.md) for complete architecture documentation.

### v0.0.1 Structure (Current)

Simple monolithic structure for rapid prototyping:

```
MailMergeKit/
├── src/
│   └── MailMergeKit.WordAddin/
│       ├── ThisAddIn.cs
│       ├── Ribbon.cs
│       ├── OutlookMailer.cs
│       └── SettingsForm.cs
├── docs/
├── examples/
└── tests/
```

### v1.0 Structure (Target)

Clean Architecture with 4 layers:

```
src/
├── MailMergeKit.WordAddin/      # Presentation
├── MailMergeKit.Application/    # Workflows
├── MailMergeKit.Core/           # Business logic
└── MailMergeKit.Infrastructure/ # External systems
```

---

## Building from Source

### 1. Clone Repository

```bash
git clone https://github.com/ProgrammerNomad/MailMergeKit.git
cd MailMergeKit
```

### 2. Restore NuGet Packages

```bash
dotnet restore
```

Or in Visual Studio: `Tools` → `NuGet Package Manager` → `Restore NuGet Packages`

### 3. Build Solution

```bash
dotnet build MailMergeKit.sln --configuration Debug
```

Or in Visual Studio: `Ctrl+Shift+B`

### 4. Debug the Add-in

- Set `MailMergeKit.WordAddin` as startup project
- Press `F5`
- Word will launch with the add-in loaded

---

## Running Tests

[v0.1.0+]

### Unit Tests

```bash
dotnet test tests/MailMergeKit.Tests
```

### With Coverage

```bash
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
```

---

## Contributing

### Development Workflow

1. **Fork** the repository
2. **Create branch** from `develop`
   ```bash
   git checkout -b feature/your-feature
   ```
3. **Make changes** following code style guidelines
4. **Write tests** for new features
5. **Update documentation** if needed
6. **Commit** with clear messages
   ```bash
   git commit -m "Add: Subject field validation"
   ```
7. **Push** to your fork
   ```bash
   git push origin feature/your-feature
   ```
8. **Create Pull Request** to `develop` branch

### Branch Strategy

- `main` - Stable releases only
- `develop` - Active development
- `feature/*` - New features
- `bugfix/*` - Bug fixes
- `hotfix/*` - Critical fixes to main

---

## Code Style

### C# Conventions

Follow [Microsoft C# Coding Conventions](https://docs.microsoft.com/en-us/dotnet/csharp/fundamentals/coding-style/coding-conventions).

**Key rules:**
- Use `PascalCase` for public members
- Use `camelCase` for private fields with `_` prefix
- Use meaningful variable names
- Add XML documentation for public APIs

**Example:**
```csharp
/// <summary>
/// Creates a draft email for the specified recipient.
/// </summary>
/// <param name="recipient">The recipient data.</param>
/// <exception cref="ArgumentNullException">If recipient is null.</exception>
public void CreateDraft(RecipientData recipient)
{
    if (recipient == null)
        throw new ArgumentNullException(nameof(recipient));
    
    var mail = _outlookApp.CreateItem(OlItemType.olMailItem);
    // ...
}
```

### Project Organization

```
MailMergeKit.Core/
├── Models/          # Data models
├── Services/        # Business services
├── Interfaces/      # Contracts
└── Validation/      # Validation rules
```

### Dependency Rules

**Allowed:**
```
UI → Application → Core
Infrastructure → Core
```

**Forbidden:**
```
Core → Infrastructure ❌
Core → UI ❌
```

---

## COM Development

### Critical Rules

1. **Single-threaded only** - No `Parallel.ForEach` with COM
2. **Always release COM objects:**
   ```csharp
   try
   {
       // Use COM object
   }
   finally
   {
       if (comObject != null)
           Marshal.ReleaseComObject(comObject);
   }
   ```
3. **Reuse single Outlook instance** - Don't create multiple COM instances
4. **Check for null** before releasing

### Debugging COM Issues

- Enable via Visual Studio: `Tools` → `Options` → `Debugging` → `Enable COM Interop Debugging`
- Monitor with: Process Explorer (check handle counts)
- Log all COM operations

---

## Logging

Use **Serilog** for all logging:

```csharp
Log.Information("Processing {Email}", recipient.Email);
Log.Warning("Attachment not found: {Path}", attachmentPath);
Log.Error(ex, "Failed to create draft for {Email}", recipient.Email);
```

**Never use:**
- `Console.WriteLine()`
- `Debug.WriteLine()`
- Custom logging classes

---

## Testing Guidelines

### Unit Tests

Test business logic without COM dependencies:

```csharp
[Fact]
public void MergeEngine_ValidatesEmailFormat()
{
    // Arrange
    var validator = new RecipientValidator();
    var recipient = new RecipientData { Email = "invalid" };
    
    // Act
    var result = validator.Validate(recipient);
    
    // Assert
    Assert.False(result.IsValid);
}
```

### Integration Tests

Mock COM interfaces:

```csharp
var mockOutlook = new Mock<Outlook.Application>();
var sut = new OutlookService(mockOutlook.Object);
```

---

## Release Process

[v1.0+]

1. Update `CHANGELOG.md`
2. Bump version in `AssemblyInfo.cs`
3. Create git tag: `v0.1.0`
4. Build installer
5. Create GitHub release
6. Update documentation

---

## Resources

### Documentation

- [Architecture Guide](architecture.md)
- [Technical Stack](TECHNICAL_STACK.md)
- [Changelog](CHANGELOG.md)

### External References

- [VSTO Documentation](https://docs.microsoft.com/en-us/visualstudio/vsto/)
- [Outlook Object Model](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook)
- [Word Object Model](https://docs.microsoft.com/en-us/office/vba/api/overview/word)

---

## Getting Help

- **Issues:** [GitHub Issues](https://github.com/ProgrammerNomad/MailMergeKit/issues)
- **Discussions:** [GitHub Discussions](https://github.com/ProgrammerNomad/MailMergeKit/discussions)
- **Email:** dev@mailmergekit.com

---

**Status:** Living document - updated with each release
