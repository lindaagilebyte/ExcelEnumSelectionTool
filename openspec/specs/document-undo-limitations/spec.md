# document-undo-limitations

## Context
When a VBA macro modifies an Excel worksheet (e.g., `ActiveCell.Value = "..."`), Excel clears the entire Undo stack. To mitigate this frustration for users of the Enum Selection Tool, a custom Undo handler was implemented using `Application.OnUndo`. However, this custom handler only remembers the single cell modification made by the form, not the previous history. We need to document this behavior clearly in the `README.md` so users understand the limitations.

## Requirements

### Requirement: Explain VBA Undo limitation
The documentation MUST clearly state that Excel natively destroys the Undo stack when a VBA macro alters the worksheet. This sets the context for why the tool behaves differently than normal Excel operations.

#### Scenario: User wonders why they can't undo past actions
- **WHEN** a user looks at the Troubleshooting or Usage section of the README
- **THEN** they find an explicit warning that running the macro clears prior Undo history.

### Requirement: Explain custom single-step Undo patch
The documentation MUST explain that the tool provides a custom `Ctrl+Z` (Undo) function specifically to revert the choice made in the dropdown menu. 

#### Scenario: User makes a mistake in the dropdown
- **WHEN** a user selects the wrong item and clicks Confirm
- **THEN** the documentation tells them they can press `Ctrl + Z` to revert that specific cell change.

### Requirement: Clarify the single-step limitation
The documentation MUST clarify that this custom Undo patch *only* reverts the most recent cell modification made by the tool itself, and cannot restore any actions taken before the tool was launched.

#### Scenario: User tries to undo multiple times
- **WHEN** a user presses `Ctrl + Z` multiple times after using the tool
- **THEN** the documentation has prepared them to expect that only the tool's action is reverted, and further undos will not work.
