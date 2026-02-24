# Design: update-documentation

## Overview
This design outlines how we will update the `README.md` to communicate the technical limitations of Excel VBA's Undo stack and our custom patch for the Enum Selection Tool.

## Context
Excel natively clears the entire Undo history whenever a VBA macro modifies a worksheet. Because our tool relies on a macro to inject the selected enum value into a cell, users lose their ability to undo prior actions. To mitigate this, we implemented a custom `Application.OnUndo` handler. However, this handler only remembers the single cell change made by the UserForm. 

## Goals / Non-Goals

**Goals:**
- Provide clear, upfront warnings in the documentation so users are not surprised when their Undo history is erased.
- Explain the exact capabilities and limitations of the custom `Ctrl+Z` patch.
- Keep the explanation accessible to non-technical Excel users while still being technically accurate.

**Non-Goals:**
- We are *not* attempting to build a comprehensive, multi-level custom Undo stack in VBA (this is notoriously difficult and unstable). This change is strictly about documentation.

## Decisions

### 1. Update "Part 2: Using the Excel Dropdown Menu -> 3. Undo Support"
Currently, the README says:
> If you click Confirm by mistake, you can immediately hit `Ctrl + Z` (or click Undo in Excel's top bar) to instantly revert the cell back to its original value.

**Decision:** We will expand this section significantly.
- Add a bold **Warning** explaining the native Excel VBA limitation (that running the tool clears the standard Undo stack).
- Clarify that the supported `Ctrl + Z` is a custom patch that *only* works for the immediate cell change made by the Tool, and it cannot restore any history prior to opening the menu.

### 2. Update "Part 1: Injecting the Macro" or "Introduction" (Optional but Recommended)
**Decision:** We should add a very brief note in the introduction or overview emphasizing that this tool includes macros that behave differently from standard Excel typing, specifically regarding Undo behavior. 
*Reasoning: Setting expectations early is better than hiding warnings deep in the usage section.*

## Risks / Trade-offs

- **Risk:** Users might still not read the documentation and get frustrated.
- **Trade-off:** Making the documentation longer to explain technical nuances might make it slightly less "quick-start" friendly, but it's necessary for setting accurate expectations for a destructive action like clearing the Undo stack.
