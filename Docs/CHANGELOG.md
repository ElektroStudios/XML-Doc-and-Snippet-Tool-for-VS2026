# Snippet Tool for Visual Studio 2026 Change Log 📋

## v1.0.1 *(current)* 🆕

### Fixed
- Close the `Code` element correctly in generated VB.NET and C# snippets.
- Escape snippet author metadata, split embedded CDATA terminators, and preserve literal template markers in selected source.
- Write snippet files as UTF-8, matching their XML declaration, without leaving an unused `.tmp` file behind.
- Calculate indentation safely for short, blank, and whitespace-only selections.
- Escape XML-sensitive characters in reference attributes and literal code operations.
- Preserve valid nested XML and entities when applying bold, italic, underline, paragraph, and remarks formatting.
- Handle CRLF, LF, and CR line endings in generated code examples and remarks; avoid duplicating documentation prefixes in remarks.
- Guard empty selections before calculating the last selected character.
- Register menu handlers after deferred DTE initialization and recheck registration state after asynchronous service acquisition.
- Remove the obsolete `InstalledProductRegistration` attribute and use the existing VSIX manifest for extension metadata.
- Align assembly and VSIX version metadata.

### Tests and build
- Add xUnit v3 behavior and command-contract tests, including real DTE listener tests with a mocked shell boundary.

## v1.0
Initial Release.
