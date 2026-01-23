# LabelIt 🏷️

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Lazarus](https://img.shields.io/badge/Lazarus-Free_Pascal-blue.svg)](https://www.lazarus-ide.org/)
[![Antigravity](https://img.shields.io/badge/IDE-Antigravity-darkblue.svg)](https://antigravity.google/)

> *"It's simple to look complicated; the complicated thing is to look simple."*

> [!IMPORTANT]
> **Project Status:** This application is currently in the active development phase of its **first Beta version**. Some features may undergo significant changes.

**LabelIt** is a robust and intuitive label design and printing platform, developed in **Lazarus (Free Pascal)**. Focused on productivity and ease of use, it allows for the creation of complex labels with simplified dynamic data integration.

---

## ✨ Features

- **🚀 Drag-and-Drop Visual Designer:** Fluid interface for real-time object positioning and resizing.
- **📦 Supported Objects:**
  - **Text:** Full support for fonts, alignments, and colors.
  - **Images:** Insertion of static graphics or via data binding.
  - **Barcodes:** Integration with `lazbarcodes` (EAN13, Code128, QR Code, etc.).
  - **Dynamic Grids:** Advanced table editing with cell merging.
- **🔗 Comprehensive Data Binding:** Link your label fields to various data sources, including:
  - **SQL Databases:** SQL Server, MySQL, SQLite, and others.
  - **Spreadsheets:** Excel files and other structured data formats.
- **📄 .lit Format:** Optimized file system (ZIP) that groups definitions, queries, and metadata.
- **🛠️ Property Grid:** Precise editing of all object attributes, including batch editing for multi-selection.

---

## 🛠️ Technologies and Dependencies

The application was built using the **Object Pascal** ecosystem:

- **IDE:** [Lazarus](https://www.lazarus-ide.org/) (2.2.0 or higher recommended)
- **Compiler:** Free Pascal (FPC)
- **Required Libraries:**
  - `lazbarcodes` (available via Online Package Manager)

---

## 🚀 How to Compile

1. **Install Lazarus:** Download and install the latest version from [lazarus-ide.org](https://www.lazarus-ide.org/).
2. **Install Dependencies:** In Lazarus, go to `Package` -> `Online Package Manager` and search for/install `lazbarcodes`.
3. **Open the Project:** In the `Project` menu, select `Open Project` and choose the `.lpi` file in the `Source` folder.
4. **Compile:** Press `F9` to compile and run.

---

## 🎯 Design Philosophy (KISS)

**LabelIt** strictly follows the **KISS (Keep It Simple and Stable)** philosophy. We believe that a powerful tool does not need to be complicated. Every line of code is written with stability and clarity for the end-user in mind.

---

## 📄 License

This project is licensed under the MIT License - see the `LICENSE` file for more details.

---

<p align="center">
  Developed with ❤️ by José Tomás Rodrigues
</p>


