# Spreadsheet GUI

A Lisp-powered spreadsheet application built with Common Lisp and LTK

![Version](https://img.shields.io/badge/version-0.5.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[日本語版 README](README-JP.md)

## Overview

A unique spreadsheet where formulas are written in Lisp S-expressions. Cells can hold any Lisp value including numbers, lists, symbols, and strings. Supports lambda expressions, auto-recalculation, Undo/Redo, file save/load, and CSV import/export.

## Screenshot

```
┌─────────────────────────────────────────────────────────────────────┐
│ [File] [Edit]                                                       │
├─────────────────────────────────────────────────────────────────────┤
│ =(+ A1 5)                                                           │
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │         10 │ (1 2 3 4 5)            │          │               │
│  2 │         15 │ ← Auto-updated!        │          │               │
│  3 │         30 │ Hello                  │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
         ↑ Numbers right-aligned    ↑ Text left-aligned
```

## Features

### Smart Formula References (v0.5)
- Auto-update references on row/column insert/delete
- Supports absolute references (A1, B2), relative references (rel), and relative ranges (rel-range)
- #REF! error for references to deleted cells
- Confirmation dialog when data loss may occur

### User Interface
- Context menu on headers (right-click)
- Resizable columns and rows (drag header borders)
- Column width / Row height setting dialogs
- Insert/Delete rows and columns
- Number right-alignment, text left-alignment
- Direct input mode (start typing to edit)

### Formula System
- S-expression formulas with full Lisp syntax
- Lambda expressions for custom transformations
- Position references: `this-row`, `this-col`, `rel`, `rel-range`
- 80+ whitelisted pure Lisp functions
- Auto-recalculation with dependency tracking

### Data Management
- Undo/Redo with up to 100 operations
- File save/load in native `.ssp` format
- CSV import/export for Excel/Google Sheets compatibility
- Range selection with drag or Shift+Arrow
- System clipboard support (TSV format)

## Requirements

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5+

## Installation

```bash
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# Install Quicklisp
curl -O https://beta.quicklisp.org/quicklisp.lisp
sbcl --load quicklisp.lisp
```

```lisp
(quicklisp-quickstart:install)
(ql:add-to-init-file)
(quit)
```

## Usage

```lisp
(load "spreadsheet-gui.lisp")
(ss-gui:start)
```

### Startup Parameters

```lisp
(ss-gui:start)                           ; Default: 26 rows, 14 columns
(ss-gui:start :rows 30 :cols 10)         ; Custom size
(ss-gui:start :input-lines 5)            ; Larger input area
```

## Keyboard Shortcuts

### Navigation & Selection

| Key | Action |
|-----|--------|
| Arrow keys | Move cursor |
| Shift+Arrow | Extend selection |
| Click/Drag | Select cells |
| Any character | Start typing (clears cell) |
| F2 | Edit mode (keep content) |
| Escape | Cancel editing |
| Delete/BackSpace | Clear cells |

### Input Confirmation

| Key | Action |
|-----|--------|
| Enter | Confirm → Move down |
| Ctrl+Enter | Confirm → Stay |
| Shift+Enter | Newline in input |
| Alt+Enter | Confirm → Move right |
| Shift+Alt+Enter | Confirm → Move left |

### Edit Operations

| Key | Action |
|-----|--------|
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+X | Cut |
| Ctrl+C | Copy |
| Ctrl+V | Paste |

### File Operations

| Key | Action |
|-----|--------|
| Ctrl+N | New sheet |
| Ctrl+O | Open file |
| Ctrl+S | Save file |

## Menus

### File Menu

| Item | Description |
|------|-------------|
| New | Create new empty sheet |
| Open... | Open .ssp file |
| Save | Save to current file |
| Save As... | Save as new file |
| Import CSV... | Import from CSV |
| Export CSV... | Export to CSV |
| Exit | Exit application |

### Edit Menu

| Item | Description |
|------|-------------|
| Undo | Undo (Ctrl+Z) |
| Redo | Redo (Ctrl+Y) |
| Cut | Cut (Ctrl+X) |
| Copy | Copy (Ctrl+C) |
| Paste | Paste (Ctrl+V) |
| Delete | Delete selection |
| Insert Row | Insert row at cursor |
| Insert Column | Insert column at cursor |
| Delete Row | Delete row at cursor |
| Delete Column | Delete column at cursor |

### Context Menu (Right-Click on Headers)

**On Column Header (A, B, C...):**

| Item | Description |
|------|-------------|
| Insert Column | Insert column at clicked position |
| Delete Column | Delete clicked column |
| Column Width... | Set column width (dialog) |

**On Row Header (1, 2, 3...):**

| Item | Description |
|------|-------------|
| Insert Row | Insert row at clicked position |
| Delete Row | Delete clicked row |
| Row Height... | Set row height (dialog) |

## Formula Examples

### Basic Arithmetic

```lisp
=(+ A1 B1)           ; Add
=(* A1 2)            ; Multiply
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
```

### Range Operations

```lisp
=(sum (range A1 A10))   ; Sum of A1:A10
=(avg (range A1 A26))   ; Average
=(max (range B1 B26))   ; Maximum
```

### Relative References

```lisp
=(rel -1 0)          ; One row up
=(rel 0 -1)          ; One column left
=(sum (rel-range -4 0 -1 0))  ; Sum of 4 rows above
```

When rows/columns are inserted or deleted, relative references are automatically adjusted to maintain the same logical reference.

### Lambda Expressions

```lisp
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)
```

### Conditionals

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## Smart Formula Reference Updates

When inserting or deleting rows/columns, formula references are automatically updated:

### Absolute References

| Operation | Before | After |
|-----------|--------|-------|
| Insert row 2 | =A3 | =A4 |
| Delete row 2 | =A3 | =A2 |
| Delete row 2 | =A2 | #REF! |
| Insert col B | =C1 | =D1 |

### Relative References

| Operation | Cell | Before | After |
|-----------|------|--------|-------|
| Insert row 2 | B3→B4 | =(rel -1 0) | =(rel -2 0) |
| Insert col B | C1→D1 | =(rel 0 -1) | =(rel 0 -2) |

The relative offset is adjusted so the formula continues to reference the same original cell.

## File Operations

### Save and Load (.ssp format)

```lisp
;; Save current spreadsheet
(ss-gui:save "mydata.ssp")

;; Load from file
(ss-gui:load-file "mydata.ssp")

;; Create new empty sheet
(ss-gui:new-sheet)

;; Check current file
ss-gui:*current-file*
```

### .ssp File Format

The `.ssp` format is a human-readable S-expression format:

```lisp
(:spreadsheet
 :format-version 1
 :metadata (:created "2025-01-15T10:30:00"
            :app-version "0.5")
 :grid (:rows 26 :cols 14)
 :cells (("A1" 10)
         ("A2" 15 (+ a1 5))
         ("A3" 30 (* a2 2))
         ("B1" (1 2 3 4 5))))
```

### CSV Export/Import

```lisp
;; Export to CSV (data only)
(ss-gui:export-csv "data.csv")

;; Export with column header (A, B, C...)
(ss-gui:export-csv "data.csv" :include-header t)

;; Export with formulas
(ss-gui:export-csv "data.csv" :include-formulas t)

;; Import from CSV
(ss-gui:import-csv "data.csv")

;; Import skipping header row
(ss-gui:import-csv "data.csv" :has-header t)
```

## Error Messages

| Error | Meaning |
|-------|---------|
| #構文: ... | Syntax error in formula |
| #評価: ... | Evaluation error |
| #循環: A1 | Circular reference detected |
| #REF! | Reference to deleted cell |

## Version History

### v0.5.1 (Current)
- Resizable input area (drag the border between input and spreadsheet)
- Fixed-width font for input area
- Tab width set to 4 characters

### v0.5
- **Smart formula reference updates** on row/column insert/delete
  - Absolute references (A1, B2, etc.)
  - Relative references (rel row-offset col-offset)
  - Relative ranges (rel-range start-row start-col end-row end-col)
- **#REF! error** for references to deleted cells
- **Confirmation dialog** when insert would cause data loss
- Context menu on headers (right-click)
- Column width / Row height dialogs
- Resizable columns and rows (drag header borders)
- Insert/Delete rows and columns (Edit menu)

### v0.4
- File menu with save/load (.ssp format)
- Edit menu with Undo/Redo/Cut/Copy/Paste
- Direct input mode (type to start editing)
- F2 for edit mode
- Multiple Enter key behaviors
- CSV import/export
- Keyboard shortcuts (Ctrl+N/O/S/Z/Y/X/C/V)

### v0.3
- Auto-recalculation with dependency tracking
- Range selection (drag, Shift+Arrow)
- Copy & Paste with system clipboard
- Position reference functions (rel, rel-range)
- Lambda expressions
- 80+ Lisp functions

### v0.2
- Basic spreadsheet functionality
- S-expression formulas

## License

MIT License

## Author

Fukuyori
