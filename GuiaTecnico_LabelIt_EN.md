# LabelIt Project Technical Guide

## 1. Introduction

**LabelIt** is a label design and printing application developed in **Lazarus (Free Pascal)**. The main goal is to provide an intuitive tool for creating custom labels, allowing the inclusion of text, images, barcodes, and tables, with support for external data binding.

The development philosophy follows the **KISS (Keep It Simple and Stable)** principle, prioritizing ease of use and code stability.

---

## 2. System Architecture

### 2.1. Technologies Used

- **Language:** Object Pascal (Delphi Mode).
- **IDE:** Lazarus.
- **Compiler:** Free Pascal (FPC).
- **Main Libraries:**
  - `lazbarcodes`: For generating barcodes (EAN13, Code128, QR Code).
  - `Zipper`: For compression and persistence of `.lit` files.
  - `Graphics`: For visual rendering in the designer.

### 2.2. File Structure

- `uMain.pas`: Main application form and menu/file management.
- `uEtiqueta.pas`: The "heart" of the editor, containing the visual designer logic, event management, property grid, and window definitions (`TTabSettings`).
- `uLabelObjects.pas`: Definition of label object classes (`TLabelItem` and derivatives).
- `uLabelIt.pas`: Global data structures and persistence logic (Save/Load).
- `uDadosExternos.pas`: Management of database connections (SQL Server, MySQL, SQLite) and data files.

### 2.3. Code Standards

The project strictly follows Pascal programming best practices:

- **Naming:** `T` prefix for types, `F` for private fields, `L` for local variables, `A` for arguments, and `C` for constants.
- **Encapsulation:** Extensive use of properties and methods for state manipulation.
- **Maintainability:** Clear separation between interface (`interface`) and implementation (`implementation`).

---

## 3. Object Model (`uLabelObjects.pas`)

All visible elements on the label derive from the `TLabelItem` base class.

### 3.1. Class Hierarchy

- **`TLabelItem`**: Abstract class defining common properties such as `Left`, `Top`, `Width`, `Height`, `Visible`, `FieldName`, and `RecordOffset`.
  - **`TLabelText`**: Represents text blocks with support for fonts and alignment.
  - **`TLabelImage`**: Allows inserting images (static or via Data Binding).
  - **`TLabelBarcode`**: Generates dynamic barcodes.
  - **`TLabelGrid`**: A complex table structure with support for:
    - Merging adjacent cells (Merge/Span).
    - Individual background color per cell.
    - Border formatting (Top, Right, Bottom, Left) with independent visibility and colors.

### 3.2. Object Management (`TLabelManager`)

The `TLabelManager` class is responsible for maintaining the list of active objects, managing multi-selection, duplication, and sorting operations (Z-Order).

---

## 4. Data Persistence

Labels are saved in the proprietary `.lit` format, which is a compressed (ZIP) file. The application ensures data integrity through:

- **Encoding:** Full UTF-8 support to ensure the correct representation of special characters (ã, ç, á, etc.).
- **Corruption Detection:** Intelligent mechanism that detects inconsistencies between `TipoDadosExternos` and the configured files (`SQL.r`, `SQLQuery.sql`), proceeding with automatic correction when necessary.

### 4.1. Internal .lit File Structure

The `.lit` file is a ZIP archive that organizes data at different levels: the root for global connection data and subdirectories for each model/label type present in the project.

| Path / File               | Description              | Main Content
|---------------------------|--------------------------|-------------------
| **`E.0/`**, **`E.1/`**... | **Model Directories**   | Containers for different label definitions in the same file.
| `E.N/Main.r`              | Model Settings           | Printer, Paper, Dimensions, Margins, Orientation, and Layout.
| `E.N/Obj.0.r`...          | Label Objects            | Binary serialized data for each component within that model.
| `SQL.r`                   | Connection Settings     | Data type, Version (MySQL), Server, Database, User, Password, and Parameters.
| `SQLQuery.sql`            | Data Query               | SQL command or spreadsheet metadata.

---

## 5. Advanced Features

### 5.1. Visual Designer

- **Designer Per Tab:** Each open tab has its own zoom settings, orientation, and visual state (`TTabSettings`).
- **Multi-Selection:** Batch editing in the property grid for multiple selected objects.
- **Clipboard:** Cut/Copy/Paste system that preserves all object properties, including data bindings.

### 5.2. Data Binding

Objects can be dynamically linked:

- **`FieldName`**: Field name in the database.
- **`RecordOffset`**: Allows different objects to point to subsequent records on the same printed page.
- **Intelligence:** Automatic update of the designer view when navigating through external data.

### 5.3. Property Grid (`sgProperties`)

The grid has been optimized for productivity:

- **Grouping:** Properties organized into expandable sections (➖Font, ➖Cell, ➖Borders).
- **Priority Hierarchy:** When editing grid cells, the display priority is: Child Object > Cell > Table (Grid).
- **Specialized Editors:** Integrated dialogs for Colors, Fonts, and Files.

### 5.4. Layout Automation

The system automatically detects dimensions and layout based on the printer's paper name:

- Identification of continuous paper vs. fixed labels.
- Automatic calculation of the number of columns (e.g., "54mm x 3" sets 3 columns).
- Dynamic adjustment of margins and orientation.

---

## 6. Printing Flow

Printing uses precise definitions in millimeters, converting logical coordinates to the printer's DPI. The rendering engine ensures that what is seen in the designer is exactly what is printed (WYSIWYG).

---

## 7. Maintenance Notes and Philosophy
>
> *"It's simple to look complicated, the complicated thing is to look simple."*

When performing maintenance, observe:

1. **Consistency:** New fields must be integrated into `SaveToWriter`, `LoadFromReader`, and `Clone`.
2. **Compiler Alerts:** String conversion warnings are monitored but allowed in safe UTF-8 contexts. For new hints, use `{%H-}` if the parameter is not used.
3. **Performance:** Grid drawing and barcode operations must be efficient to maintain designer fluidity at high zoom levels.

---
*Report updated on January 22, 2026.*
