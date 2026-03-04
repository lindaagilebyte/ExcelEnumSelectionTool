# Specs: Add User Manual for Enum Selection Tool

This change introduces documentation for the `user-manual` capability. 
Because this is purely documentation, no existing codebase behaviors are modified.

## ADDED Requirements

### Requirement: Document the Anchor Text Rule
- **Description**: The manual MUST explicitly state that the tool searches for the exact text `定義(巨集顯示)` to locate enum definitions.
#### Scenario: Planner needs to understand how definitions are found
- **WHEN** a planner reads the manual to understand the tool's mechanics
- **THEN** they learn that `定義(巨集顯示)` acts as the anchor point for all extractions.

### Requirement: Document Spatial Positioning
- **Description**: The manual MUST explain the relative positioning of the Enum Key and Enum Values relative to the anchor text.
#### Scenario: Planner adds a new Enum Key
- **WHEN** a planner wants to define a new dropdown list named "ItemRarity"
- **THEN** the manual instructs them to place the key exactly one row above and one column to the left of the anchor text (e.g., Key at `A10`, Anchor at `B11`).
#### Scenario: Planner adds values to an Enum
- **WHEN** a planner has defined the key and anchor
- **THEN** the manual instructs them to list the dropdown values starting directly below the anchor text (e.g., starting at `B12` downwards without blank spaces).

### Requirement: Document Data Format Constraints
- **Description**: The manual MUST explain that the list of values must be continuous, stopping at the first empty cell.
#### Scenario: Planner enters non-continuous data
- **WHEN** a planner inadvertently leaves a blank row inside an enum list
- **THEN** the manual warns them that the tool stops reading at the first empty cell, and values after the blank will be ignored.

### Requirement: Bilingual Documentation
- **Description**: The manual MUST include an identical Traditional Chinese (Taiwan) translation following the English version.
#### Scenario: Taiwanese Planner needs documentation
- **WHEN** a planner who prefers Traditional Chinese reads the manual
- **THEN** they find a fully translated version of all rules and guides at the end of the document.
