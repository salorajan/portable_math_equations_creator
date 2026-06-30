# Portable Equation Converter (PMC)

A high-performance document accessibility utility that parses mathematical equations typed in multiple notation dialects and compiles source files into highly accessible MS Word, HTML, PDF, or Markdown formats conforming to **WCAG 2.2 (Level A/AA/AAA)** standards.

Designed to assist blind and visually impaired students in creating and reading accessible technical documents.

---

## Key Features

- **Multi-Dialect Math Parsing**: Supports **LaTeX**, **Typst typeset**, and **SymPy** formulas. Auto-detects dialects inside standard math delimiters (`$` and `$$`) and translates them into uniform MathML or OMML targets.
- **Standalone Accessibility HTML**: Outputs custom-styled responsive HTML web pages utilizing MathML 3.0, high-contrast colors (slate gray text on off-white background), relative sizing, text spacing, clear visual focus rings, skip-to-content links, and keyboard-navigable scrollable block math containers. Passes `Axe-Core` validation without errors.
- **Accessible MS Word (DOCX)**: Converts formulas into native Microsoft OMML math equations. Post-processes files to set table headers to repeat across pages (`tblHeader`) and prevents rows from splitting across pages (`cantSplit`).
- **Tagged PDFs**: Compiles tagged, accessible PDFs with nested outlines via MS Word COM translation.
- **Drag-and-Drop Interactive CLI**: Batch script wrapper `run_converter.bat` allows users to convert documents interactively or via simple drag-and-drop.

---

## Project Directory Structure

```
equation_converter/
│
├── run_converter.bat      # Master interactive batch utility for Windows (supports Drag-and-Drop)
├── run_converter.sh       # Master interactive bash utility for macOS
├── requirements.txt       # Python package dependencies
├── .gitignore             # Configured git ignore patterns
│
├── bin/                   # Portable binaries folder
│   ├── convert_md.exe     # Standalone compiled application executable for Windows
│   ├── convert_md_mac_arm64 # Standalone compiled application executable for macOS (Apple Silicon)
│   ├── convert_md_mac_x86_64 # Standalone compiled application executable for macOS (Intel)
│   ├── convert_md_linux   # Standalone compiled application executable for Linux
│   ├── pandoc.exe         # Portable Pandoc document converter (Windows)
│   ├── tidy.exe           # Portable HTML Tidy structure validator (Windows)
│   ├── typst.exe          # Portable Typst PDF compiler (Windows)
│   └── (Mac Binaries)     # Equivalent binaries for macOS platforms (pandoc-mac-*, typst-mac-*)
```

---

## Installation & Setup

All required engines (Pandoc, Typst, HTML Tidy) are bundled directly inside the executables or in the `bin/` folder. The application runs immediately as a standalone portable package on Windows, macOS, and Linux without any prerequisite installations.

### Installing Python Dependencies (For Development)
If you wish to edit the source code or run via Python, install dependencies using:
```cmd
pip install -r requirements.txt
```

---

## How to Run

### 1. Using the Master Interactive Script (Recommended)
#### Windows:
Simply double-click `run_converter.bat` in the root folder, or **drag and drop** your source document (Markdown, DOCX, or HTML) directly onto the `run_converter.bat` file icon. 

#### macOS & Linux:
Open a terminal in the root folder and run:
```bash
chmod +x run_converter.sh
./run_converter.sh [input_file]
```

The script will automatically detect the path (or prompt you to enter it) and prompt you to choose one of the core output formats:
1. **MS Word (.docx)**
2. **HTML (.html)**
3. **PDF (.pdf)**
4. **Markdown (.md)**

### 2. Using the Command Line Executable
You can run the compiled binary directly from your terminal:

**Windows:**
```cmd
bin\convert_md.exe <input_file> <word|html|pdf|md> [output_file]
```

**macOS (Apple Silicon):**
```bash
bin/convert_md_mac_arm64 <input_file> <word|html|pdf|md> [output_file]
```

**macOS (Intel):**
```bash
bin/convert_md_mac_x86_64 <input_file> <word|html|pdf|md> [output_file]
```

**Linux:**
```bash
bin/convert_md_linux <input_file> <word|html|pdf|md> [output_file]
```
*If `output_file` is omitted, the utility automatically appends `_out` (e.g. `test2.md` -> `test2_out.html`).*

**Examples (Windows):**
```cmd
bin\convert_md.exe test2.md word
bin\convert_md.exe test2.md html test_webpage.html
bin\convert_md.exe test2.md pdf
bin\convert_md.exe test2.md md test_remediated.md
```

---

## Mathematical Notation Cheatsheet

Write equations inside your source Markdown document using your preferred notation:

| Math Notation | Syntax Style | Examples |
| :--- | :--- | :--- |
| **LaTeX** | Standard delimiter blocks | `$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$` |
| **Typst** | Python-like syntax, no backslashes | `t:{x = (-b +- sqrt(b^2 - 4 a c)) / (2 a)}` <br> `$$ sum_(k=0)^n binom(n, k) a^(n-k) b^k $$` |
| **SymPy** | Evaluatable symbolic math code | `py:{a**2 - b**2 = (a - b)*(a + b)}` <br> `py:{det(Matrix([[a, b], [c, d]])) = a*d - b*c}` |

---

## Source Code & Rebuilding

The source code and building scripts reside on the `dev` branch of the repository. If you need to modify the source code or compile new binaries, switch to the `dev` branch:
1. Checkout the `dev` branch: `git checkout dev`.
2. Locate the source code in `references/convert_md.py`.
3. To rebuild the executables, push your changes to the `dev` branch. The GitHub Actions workflow will automatically run testing, rebuild the binaries for all three platforms, and deploy them directly back to the `main` branch.

---

## WCAG 2.2 Accessibility Enforcement

Every converted document is automatically structured to meet Web Content Accessibility Guidelines (WCAG) 2.2:
- **Heading Outline Integrity (WCAG 1.3.1)**: Audits header syntax and warns the user if header levels are skipped (e.g., `#` directly to `###`), preventing navigation disorientation.
- **Image Alternative Text (WCAG 1.1.1)**: Checks all image markdown tags. Flags empty alt texts and injects placeholder warnings specifying the exact line number requiring remediation.
- **Repeating Headers & Split Prevention (WCAG 1.3.1)**: Alters DOCX tables to ensure the first row repeats on page boundaries and individual rows never break across pages.
- **Screen Reader Math Fallbacks**: HTML outputs include MathML tags enriched with standard `alttext` and `aria-label` fields containing the LaTeX string, allowing talking web browsers to speak the formula.
- **Contrast & Font Spacing (WCAG 1.4.3 & 1.4.12)**: Embedded CSS style blocks override browser layouts to guarantee readable line spacing, letter/word spacing, clear focus indicators, and high contrast.
