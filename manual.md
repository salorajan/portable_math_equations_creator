# Portable Equation Converter (PMC) - User Manual

This manual explains how to use the Portable Equation Converter to create and compile highly accessible technical documents containing math equations. 

The tool is specifically designed to help blind and visually impaired students type math formulas inside simple Markdown documents using their preferred notation, and compile them into fully accessible **MS Word (DOCX)**, **HTML**, or **PDF** documents conforming to **WCAG 2.2** standards.

---

## 1. How to Write Math Equations

You can write math equations inside your Markdown document using three different notation styles (dialects). The converter will auto-detect the style or use explicit prefixes to translate them into accessible MathML or OMML targets.

### A. LaTeX Notation
LaTeX is the most common mathematical typesetting language. Write LaTeX equations using standard dollar sign delimiters:
*   **Inline Math**: Wrap equation in single `$` signs.
    *   *Example*: The equation `$E = mc^2$` represents mass-energy equivalence.
*   **Block Math**: Wrap equation in double `$$` signs on separate lines.
    *   *Example*:
        ```latex
        $$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$
        ```

### B. Typst Typeset Notation
Typst is a modern, python-like typesetting syntax that avoids backslashes, making it easier to read and type on screen readers. To ensure it is parsed as Typst, prefix the formula with `t:{` and close it with `}`:
*   **Inline/Block Typst**: Wrap your formula in `t:{ ... }`.
    *   *Example 1 (Quadratic Formula)*: `t:{x = (-b +- sqrt(b^2 - 4 a c)) / (2 a)}`
    *   *Example 2 (Basel Sum)*: `t:{sum_(n=1)^oo 1/n^2 = pi^2/6}`
    *   *Example 3 (Integral)*: `t:{integral_a^b f(x) dif x = F(b) - F(a)}`

### C. SymPy Notation
SymPy notation uses standard Python code math syntax. This is highly beneficial for students who write code or want to evaluate symbolic expressions. Wrap the formula in `py:{ ... }` or `sympy:{ ... }`:
*   **Inline/Block SymPy**: Wrap your formula in `py:{ ... }`.
    *   *Example 1 (Derivative)*: `py:{diff(sin(x), x) = cos(x)}`
    *   *Example 2 (Factoring)*: `py:{a**2 - b**2 = (a - b)*(a + b)}`
    *   *Example 3 (Determinant)*: `py:{det(Matrix([[a, b], [c, d]])) = a*d - b*c}`

---

## 2. Document Accessibility Guidelines (WCAG 2.2)

To ensure your document is fully accessible for screen readers and assistive technology, please follow these guidelines when writing:

1.  **Logical Headings**: Do not skip heading levels (e.g., do not place a level 3 heading `###` directly after a level 1 heading `#`). Always nest headings sequentially (`# H1` -> `## H2` -> `### H3`).
2.  **Image Alternative Text**: Never leave alt text empty on images. Always provide a description:
    *   *Accessible*: `![Graph plotting y equals sine of x showing one full period](sine_wave.png)`
    *   *Inaccessible*: `![](sine_wave.png)`
3.  **Accessible Tables**: Do not leave header cells empty or use uneven cell counts. Ensure headers clearly label all columns.

---

## 3. How to Run the Converter

All necessary engines are portable and bundled inside the `bin/` directory. You do not need to install python or external utilities to convert documents.

### Method 1: Drag-and-Drop (Windows - Recommended)
1.  Locate your source `.md` file.
2.  Drag the file icon and drop it directly onto [run_converter.bat](file:///C:/salo/acb/quill/equation_converter/run_converter.bat).
3.  A terminal window will open and prompt you to choose the target format:
    *   `1` for MS Word (`.docx`)
    *   `2` for HTML (`.html`)
    *   `3` for PDF (`.pdf`)
4.  Press Enter, and the converted file will be generated in the same directory.

### Method 2: Interactive CLI
1.  Double-click [run_converter.bat](file:///C:/salo/acb/quill/equation_converter/run_converter.bat).
2.  Type or paste the path to your document (e.g., `test0.md`) and press Enter.
3.  Select your target output format and press Enter.

### Method 3: Command Line (Advanced)
Open your terminal in the project directory and run the binary [convert_md.exe](file:///C:/salo/acb/quill/equation_converter/bin/convert_md.exe) directly:
```cmd
bin\convert_md.exe <input_file> <word|html|pdf|md> [output_file]
```
*Example*:
```cmd
bin\convert_md.exe test0.md html my_accessible_page.html
```
