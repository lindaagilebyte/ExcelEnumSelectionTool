# Design: Add User Manual for Enum Selection Tool

## Context
The Enum Selection Tool allows game planners to select predefined string values from a dropdown menu in their data configuration Excel files. This relies on a reference file named `列舉定義(企劃用).xlsx`. However, the tool's behavior regarding how it extracts data from this reference file is implicit in the VBA code and not documented for the end-users (planners).

## Goals
- Provide a clear, step-by-step guide for planners on how the tool reads `列舉定義(企劃用).xlsx`.
- Explain the spatial requirements (relative positioning of the anchor text, enum key, and values) needed for the tool to function.
- Document how to add, modify, or remove enum definitions.

## Non-Goals
- Changing the existing VBA macro code (`Module_EnumSelector.bas`).
- Changing the layout or structure of the existing `列舉定義(企劃用).xlsx` file.

## Decisions
- **Format**: The manual will be written in Markdown format (`USER_MANUAL.md`) and placed in the root directory for easy access, or alongside the `README.md`.
- **Content Structure**: The manual will include both an English version and an identical Traditional Chinese (Taiwan) version at the end. Each language section will contain:
  1.  **Overview**: What the tool is and its purpose.
  2.  **How it Works (The Magic)**: A visual explanation of the spatial scanning logic.
  3.  **Step-by-Step Guide**: How to add a new enum.
  4.  **Troubleshooting**: Common issues (e.g., misspelled anchor text, wrong relative positioning).
- **Visual Aids**: We will use ASCII diagrams or clear table structures in the markdown to represent the Excel layout, as planners are primarily visual/spreadsheet-oriented.

## Risks / Trade-offs
- **Risk**: Planners might ignore the manual and still break the file.
  - **Mitigation**: Make the manual as concise and visual as possible. Highlight the critical "Anchor Text" rule prominently.
