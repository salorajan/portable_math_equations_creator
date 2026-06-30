# Equation Converter: Accessible Mathematics Guide & Manual

This manual is designed for students, particularly those who are blind or visually impaired, to create and convert documents containing complex mathematical equations. By typing equations in formats you are familiar with—such as LaTeX, Typst (typeset), or SymPy—the Equation Converter app compiles your Markdown files into fully accessible MS Word (`.docx`), PDF (`.pdf`), and HTML (`.html`) files.

---

## 1. Overview of the Application

For visually impaired students using screen readers, writing and reading mathematics has historically been a significant barrier. Standard math images lack text equivalents, and LaTeX markup can be dense and difficult to read when spoken aloud. 

The Equation Converter app bridges this gap by allowing you to write equations in the syntax that is easiest for you to type and read.

*   **Write**: You create a Markdown file typing mathematical equations inside standard delimiters (`$ ... $` for inline math or `$$ ... $$` for block math) using your choice of LaTeX, Typst, or SymPy syntax.
*   **Compile**: You run a simple command to compile the file into your chosen target format. The app automatically detects your math dialect, audits the document structure for accessibility, and outputs:
    *   **HTML**: Web pages containing accessible **MathML** (Mathematical Markup Language) with fallback text descriptions.
    *   **MS Word**: Documents containing native Microsoft math equations (**OMML** format) that screen readers can read token-by-token.
    *   **PDF**: Tagged, high-contrast accessible documents with integrated math content descriptions.

---

## 2. Document Accessibility Features Followed

Every document compiled through the app is automatically audited and structured to conform to **WCAG 2.2 (Level A and AA)** accessibility standards:

1. **Info and Relationships (WCAG 1.3.1)**:
    *   **Heading Hierarchy**: The app audits your heading levels (`#`, `##`, `###`) to ensure they form a nested outline. Level skipping (e.g., `#` directly to `###`) is flagged, ensuring screen readers can build an accurate document outline for navigation.
    *   **Repeating Table Headers**: Word documents are structured so that the first row of every table is marked as a repeating header, meaning screen readers will repeat column names when reading table cells on subsequent pages.
    *   **Row Split Prevention**: Tables are configured to prevent rows from breaking awkwardly across page breaks.
2. **Non-Text Content (WCAG 1.1.1)**:
    *   **Alt Text Auditing & Remediation**: All images are checked for alternative text descriptions. Missing alt tags are flagged, and the system provides automated placeholder alerts pointing you to the exact line number requiring remediation.
3. **Reflow (WCAG 1.4.10)**:
    *   **Scrollable Math Containers**: In generated HTML web pages, block equations are wrapped in keyboard-focusable, horizontally scrollable containers. This prevents equations from clipping or overlapping on mobile devices or when zoomed.
4. **Contrast (WCAG 1.4.3) & Text Spacing (WCAG 1.4.12)**:
    *   **Accessible Web Styling**: HTML files include a premium, high-contrast responsive layout (slate gray text on a warm off-white background) with relative text sizing and optimized line/paragraph heights for reading comfort.
5. **Mathematical Accessibility**:
    *   **MathML Alt Text & ARIA Labels**: HTML files enrich mathematical formulas with `alttext` and `aria-label` attributes containing the LaTeX code. If a screen reader cannot parse the MathML structure, it will read aloud the LaTeX equivalent (e.g., "Formula: x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}").

---

## 3. Math Syntax Guide: LaTeX, Typst, and SymPy

This section compares the three supported mathematical syntax formats. We have selected 5 algebraic and 5 calculus equations to demonstrate how you can write them.

### Formatting Conventions
*   **LaTeX Math**: Uses standard LaTeX syntax (using backslashes like `\frac`).
*   **Typst Typeset Math**: A simpler, python-like mathematics markup language. It has no backslashes. Division is `/`, multiplication is represented by space, and symbols are written as plain words (e.g., `alpha`, `partial`, `Psi`).
*   **SymPy Math**: Evaluatable Python symbolic math code. Exponents are `**`, fractions are divisions, and equations are structured using relation constructors like `Eq(LHS, RHS)`.

---

### Algebra Equations

#### 1. Quadratic Formula
Used to find the roots of a quadratic equation of the form $ax^2 + bx + c = 0$.

*   **LaTeX Syntax**:
    ```latex
    $$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$
    ```
*   **Typst Syntax** *(simplifies fractions using parenthesis and `/`)*:
    ```typst
    $$ x = (-b +- sqrt(b^2 - 4 a c)) / (2 a) $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{x = (-b + sqrt(b**2 - 4*a*c)) / (2*a)}
    ```

#### 2. Difference of Squares
A standard algebraic identity for factoring polynomials.

*   **LaTeX Syntax**:
    ```latex
    $$a^2 - b^2 = (a - b)(a + b)$$
    ```
*   **Typst Syntax**:
    ```typst
    $$ a^2 - b^2 = (a - b) (a + b) $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{a**2 - b**2 = (a - b)*(a + b)}
    ```

#### 3. Binomial Theorem
Describes the algebraic expansion of powers of a binomial.

*   **LaTeX Syntax**:
    ```latex
    $$(a + b)^n = \sum_{k=0}^{n} \binom{n}{k} a^{n-k} b^k$$
    ```
*   **Typst Syntax** *(uses `sum_` for summation subscripts and `binom(n, k)`)*:
    ```typst
    $$ (a + b)^n = sum_(k=0)^n binom(n, k) a^(n-k) b^k $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{Eq((a + b)**n, Sum(binomial(n, k) * a**(n-k) * b**k, (k, 0, n)))}
    ```

#### 4. Logarithm Change of Base
Used to convert logarithms from one base to another.

*   **LaTeX Syntax**:
    ```latex
    $$\log_a(x) = \frac{\log_b(x)}{\log_b(a)}$$
    ```
*   **Typst Syntax**:
    ```typst
    $$ log_a(x) = log_b(x) / log_b(a) $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{log(x, a) = log(x, b) / log(a, b)}
    ```

#### 5. Matrix Determinant
The formula for the determinant of a 2x2 matrix.

*   **LaTeX Syntax**:
    ```latex
    $$\det\begin{pmatrix} a & b \\ c & d \end{pmatrix} = ad - bc$$
    ```
*   **Typst Syntax** *(defines matrices using `mat(...)` with semicolons separating rows)*:
    ```typst
    $$ det(mat(a, b; c, d)) = a d - b c $$
    ```
*   **SymPy Syntax** *(represents nested lists as matrix rows)*:
    ```python
    py:{det(Matrix([[a, b], [c, d]])) = a*d - b*c}
    ```

---

### Calculus Equations

#### 6. Limit Definition of Derivative
Represents the instantaneous rate of change of a function.

*   **LaTeX Syntax**:
    ```latex
    $$f'(x) = \lim_{h \to 0} \frac{f(x+h) - f(x)}{h}$$
    ```
*   **Typst Syntax** *(uses `lim_` for subscripts and `->` for the limit arrow)*:
    ```typst
    $$ f'(x) = lim_(h -> 0) (f(x+h) - f(x)) / h $$
    ```
*   **SymPy Syntax** *(uses `Derivative` and `Limit` constructors)*:
    ```python
    py:{Derivative(f(x), x) = Limit((f(x+h) - f(x))/h, h, 0)}
    ```

#### 7. Fundamental Theorem of Calculus
Connects differentiation and integration.

*   **LaTeX Syntax**:
    ```latex
    $$\int_{a}^{b} f(x) \, dx = F(b) - F(a)$$
    ```
*   **Typst Syntax** *(integrals defined with `integral_sub^super` and `dif x`)*:
    ```typst
    $$ integral_a^b f(x dif x) = F(b) - F(a) $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{Integral(f(x), (x, a, b)) = F(b) - F(a)}
    ```

#### 8. Integration by Parts
A rule that transforms the integral of a product of functions.

*   **LaTeX Syntax**:
    ```latex
    $$\int u \frac{dv}{dx} \, dx = uv - \int v \frac{du}{dx} \, dx$$
    ```
*   **Typst Syntax**:
    ```typst
    $$ integral u (partial v) / (partial x) d x = u v - integral v (partial u) / (partial x) d x $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{Integral(u * Derivative(v, x), x) = u * v - Integral(v * Derivative(u, x), x)}
    ```

#### 9. Taylor Series Expansion
Represents a function as an infinite sum calculated from its derivatives at a point.

*   **LaTeX Syntax**:
    ```latex
    $$f(x) = \sum_{n=0}^{\infty} \frac{f^{(n)}(a)}{n!} (x-a)^n$$
    ```
*   **Typst Syntax** *(sums to `infinity` and utilizes double-parenthesis for higher-order derivatives)*:
    ```typst
    $$ f(x) = sum_(n=0)^infinity (f^((n))(a)) / (n!) (x-a)^n $$
    ```
*   **SymPy Syntax**:
    ```python
    py:{Eq(f(x), Sum(Derivative(f(x), (x, n)).subs(x, a) / factorial(n) * (x - a)**n, (n, 0, oo)))}
    ```

#### 10. Gaussian Integral
The integral of the 1D Gaussian function over the entire real line.

*   **LaTeX Syntax**:
    ```latex
    $$\int_{-\infty}^{\infty} e^{-x^2} \, dx = \sqrt{\pi}$$
    ```
*   **Typst Syntax** *(supports negative boundaries `-infinity` and symbols like `pi`)*:
    ```typst
    $$ integral_(-infinity)^infinity e^(-x^2) d x = sqrt(pi) $$
    ```
*   **SymPy Syntax** *(uses `oo` for infinity)*:
    ```python
    py:{Integral(exp(-x**2), (x, -oo, oo)) = sqrt(pi)}
    ```

---

## 4. How to Run the App

To compile your Markdown file into your chosen document formats, open your terminal (Command Prompt or PowerShell) and run the following commands.

### 1. Compile to a Converted LaTeX Markdown File
Converts all Typst and SymPy equations into standard LaTeX:
```cmd
python convert_md.py input.md md output.md
```

### 2. Convert to an Accessible MS Word Document
Creates an MS Word document where equations are written in native Microsoft math equations (OMML format):
```cmd
python convert_md.py input.md word output.docx
```

### 3. Convert to an Accessible Web Page
Creates an accessible HTML page containing semantic MathML formulas:
```cmd
python convert_md.py input.md html output.html
```

### 4. Convert to a Tagged Accessible PDF Document
Generates a tagged, accessible PDF document:
```cmd
python convert_md.py input.md pdf output.pdf
```

---

## 5. Conclusion & Future Directions

The development of this utility is part of an ongoing, concerted effort to make math learning and scientific documentation fully accessible to blind and visually impaired students. 

### Future Roadmap
1. **DAISY Target Format Integration**:
    *   Integrating compiling options for the DAISY eBook standard (synchronizing digital talking books with semantic MathML), allowing players like Dolphin EasyReader to read and navigate math dynamically.
2. **MathCat Math-to-Speech Translation**:
    *   Integrating **MathCat** (Mathematics Cognitive Assistance Technology) as the primary audio speech rendering backend. MathCat translates MathML/LaTeX code directly into customizable, highly natural spoken descriptions, providing options for different speech rates, details, and braille outputs.
3. **AI-Driven Math Tutoring & MathChat**:
    *   Introducing interactive tutoring trees where students can converse with an LLM in math dialogue frameworks, receiving step-by-step guidance on algebra and calculus problems using standard `\boxed{}` formats.

---

## Acknowledgement

We would like to express our deepest gratitude and appreciation to **Vocational Rehabilitation Arizona (Voc Rehab Arizona)** for all of their help, invaluable guidance, and persistent support throughout this project. Their dedication to creating educational opportunities and technology tools for visually impaired students has been the driving force behind this utility's success.
