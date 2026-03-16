# Changelog

All notable changes to MailMergeKit will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned for v0.1.0
- Static and dynamic attachments
- Multiple attachments per recipient
- CC/BCC support
- Email preview functionality
- Test email mode
- Automatic Outlook startup
- Error logging with timestamps
- Pre-merge field validation
- Duplicate recipient detection

## [0.0.1] - 2026-03-16

### Added
- Initial prototype release
- Word Add-in with custom ribbon button
- Basic mail merge data source reader
- Outlook draft email generation
- Subject personalization with merge fields
- Simple settings dialog
- Queue-based processing (COM-safe architecture)
- Single Outlook.Application instance reuse
- Basic COM object cleanup

### Architecture
- **Single-threaded Outlook worker (STA-safe)** - No Parallel.ForEach
- **Queue-based sequential processing** - COM is not thread-safe
- **Word.MailMerge.DataSource integration** - Uses existing Word data
- **Draft-first workflow** - No auto-sending, always save to Drafts
- **Single Outlook.Application instance** - Reused for all emails (COM requirement)

### Known Limitations
- No attachment support yet (v0.1.0)
- No CC/BCC fields (v0.1.0)
- Manual Outlook startup required (auto-start in v0.1.0)
- No resume capability (v0.2.0)
- Basic error handling only
- No logging to file (v0.1.0)

### Notes
- This is an experimental prototype to validate core architecture
- Designed for campaigns up to ~5,000 emails (Outlook COM limitations)
- All processing is 100% local - no data leaves your computer
- Open source and free forever

---

## Version Strategy

- **0.x.x** - Experimental / early development
- **0.1.0** - First usable beta
- **0.5.0** - Feature complete
- **1.0.0** - Stable release
- **1.x.x** - New features
- **x.x.1** - Bug fixes

[Unreleased]: https://github.com/ProgrammerNomad/MailMergeKit/compare/v0.0.1...HEAD
[0.0.1]: https://github.com/ProgrammerNomad/MailMergeKit/releases/tag/v0.0.1
