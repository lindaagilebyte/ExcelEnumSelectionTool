# Proposal: update-documentation

## Goal
To update the project `README.md` to clearly communicate the limitations and custom behavior of the Undo functionality when using the Excel Enum Selection Tool. Specifically, we need to explain that because VBA macros normally destroy the Excel Undo stack, we have implemented a custom `Ctrl+Z` patch. However, this patch only remembers the *single* cell modification made by the tool itself, and any previous history is still lost.

## New Capabilities
- `document-undo-limitations`: Add a clear warning and explanation section to the documentation regarding the VBA Undo stack destruction and the custom single-step undo patch.

## Impacted Capabilities
- `documentation`: The `README.md` will be expanded to include this new information.

## Affected Systems
- `README.md`
