# vscode-bms-editor

VS Code extension for editing Mainframe BMS maps with both source and visual rendering.

## What the Extension Does

This extension provides:

- BMS language recognition for `.bms` files
- Syntax highlighting for common BMS macros and attributes
- A visual renderer for BMS maps
- Direct editing of fields on a grid
- BMS source generation from the rendered layout

## How to Open the Renderer

Open any `.bms` file, then use one of these entry points:

1. Click `Open BMS Renderer` in the editor title bar.
2. Right-click a `.bms` file in the Explorer and choose `Open BMS Renderer`.
3. Right-click inside a BMS editor and choose `Open BMS Renderer`.

The renderer opens beside the source editor.

## Renderer Layout

The renderer is divided into four main areas:

1. Top bar
	Contains view buttons, theme toggle, save button, fill mode, grid size, sync, and auto-resize.
2. Grid area
	Shows the BMS map as a terminal-style layout.
3. Field panel
	Shows the selected field properties and lets you edit them.
4. Palette
	Lets you drag new fields and arrays onto the grid.

## Basic Workflow

1. Open a `.bms` file in VS Code.
2. Open the renderer.
3. Drag fields from the palette onto the grid, or select existing fields from the parsed source.
4. Change properties in the field panel.
5. Save the generated BMS with the `Save` button.

If `Sync` is enabled, edits in the renderer also update the source file automatically.

## Views

The renderer has two views:

- `Render`: visual grid editor
- `Source`: generated BMS source preview

Use the buttons in the top bar to switch between them.

## Adding Fields

Use the palette at the bottom to drag fields onto the grid.

Available items include:

- Labels
- Input Text
- Input Num
- Horizontal arrays
- Vertical arrays
- Outlined labels
- Password fields

When you drop a field, it is created at that grid position if space is available.

## Selecting and Editing Fields

Click a field to select it.

The field panel lets you edit:

- ID
- Row
- Column
- Length
- Type
- Color
- Highlight
- Brightness
- Numeric flag
- Cursor flag
- Field Set flag
- Outline flags
- Stopper

For ASKIP text fields, the panel also shows a `Text` input.

### ID Rules

Field IDs are restricted to:

- maximum 25 characters
- letters
- numbers
- hyphen `-`

Invalid characters are removed automatically while typing.

## Inline Text Editing

Double-click an ASKIP field on the grid to edit its text directly.

If `Auto-resize` is enabled, ASKIP label fields grow automatically while you type, unless another field blocks the expansion.

## Arrays

Use the horizontal or vertical array items from the palette to create arrays.

How arrays behave:

- The number input beside the array icon controls how many elements you request.
- If the full array does not fit near the grid edge, the extension creates only the elements that actually fit.
- The resulting array metadata matches the number of elements placed on screen.

When an array is selected, the field panel shows:

- Array Cols
- Array Rows
- H Step
- V Step

You can also resize arrays from the bounding box handles shown around the selected array.

## Groups

You can group and ungroup non-array fields from the context menu.

Right-click a field or selection to access:

- Go to Definition
- Group
- Ungroup
- Delete

Arrays are treated as structured groups and cannot be ungrouped through the normal ungroup command.

## Stopper Fields

UNPROT fields can have a stopper.

When enabled, the generator emits an `ASKIP LENGTH=0` stopper immediately to the right of the field.

You can:

- enable or disable it from the field panel
- click a rendered stopper marker on the grid to select it
- press `Delete` to remove it

For array members, stopper changes propagate to the whole array.

## Save and Sync

### Save

Click `Save` to write the generated BMS back to the file.

### Sync

Enable `Sync` to keep the renderer and the source file connected:

- edits in the file update the renderer
- edits in the renderer update the file

### Persisted UI Settings

The renderer remembers these settings between sessions:

- Fill mode
- Sync
- Auto-resize
- Theme

## Generated BMS Format

The generator outputs:

- `DFHMSD` header macros
- `DFHMDI` map definition
- split `DFHMDF` field definitions with continuation lines
- blank labels for ASKIP fields
- final trailer with `DFHMSD TYPE=FINAL` and `END`

## Keyboard and Mouse Shortcuts

Supported shortcuts include:

- `Delete` or `Backspace`: delete selected field or stopper
- `Ctrl/Cmd + Z`: undo
- `Ctrl/Cmd + Y`: redo
- `Ctrl/Cmd + C`: copy selected fields
- `Ctrl/Cmd + V`: paste copied fields
- `Ctrl/Cmd + Click`: reveal the field in the source editor
- `Double-click` on ASKIP field: inline text edit

You can also use lasso selection and multi-selection with modifier keys.

## Grid Options

The toolbar also lets you control:

- `Fill`: how empty field space is rendered
- `Lines`: number of BMS rows in the grid
- `Cols`: number of BMS columns in the grid
- `Theme`: light or dark grid theme

## Development Notes

If you are running the extension from source:

1. Install dependencies with `npm install`.
2. Start the watcher with `npm run watch`.
3. Press `F5` in VS Code to open the Extension Development Host.
4. Open a `.bms` file and launch the renderer.

## Limitations

- The renderer is designed around standard BMS map editing workflows and may not preserve unusual hand-formatted source exactly as written.
- Arrays are generated as structured repeated fields, not arbitrary grouped layouts.
- Syntax highlighting is grammar-based and may not cover every custom BMS dialect nuance.
