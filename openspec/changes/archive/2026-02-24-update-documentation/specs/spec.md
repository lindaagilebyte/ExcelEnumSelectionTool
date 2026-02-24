# Specs: update-documentation

## document-undo-limitations

- **WHEN** user views the documentation for the Undo feature
- **THEN** it explicitly states VBA macros destroy the native Excel Undo stack
- **AND** it explicitly states the custom `Ctrl+Z` patch only undoes the single most recent selection made by the tool.
