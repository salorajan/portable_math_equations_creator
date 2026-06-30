# Mathematical Equations Audit Document (Algebra & Calculus)

This document contains 10 algebra and 10 calculus equations using mixed mathematical formats (LaTeX, Typst typeset, and SymPy) to test the converter pipeline.

---

## Algebra Equations

### 1. Quadratic Formula (LaTeX)
Used to find the roots of a quadratic equation of the form $ax^2 + bx + c = 0$:
$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

### 2. Difference of Squares (Explicit SymPy Wrapper)
A standard algebraic identity for factoring polynomials:
py:{a**2 - b**2 = (a - b)*(a + b)}

### 3. Binomial Theorem (Auto-Detected Typst Block)
Describes the algebraic expansion of powers of a binomial:
$$ (a + b)^n = sum_(k=0)^n binom(n, k) a^(n-k) b^k $$

### 4. Logarithm Change of Base (LaTeX)
Used to convert logarithms from one base to another:
$$\log_a(x) = \frac{\log_b(x)}{\log_b(a)}$$

### 5. Cubic Equation Cardan's Formula - Part (Explicit Typst Wrapper)
The algebraic solution for one of the roots of a reduced cubic equation:
t:{x = root(3, -q/2 + sqrt(q^2/4 + p^3/27)) + root(3, -q/2 - sqrt(q^2/4 + p^3/27))}

### 6. Matrix Determinant (Explicit SymPy Wrapper)
The formula for the determinant of a 2x2 matrix:
py:{det(Matrix([[a, b], [c, d]])) = a*d - b*c}

### 7. System of Linear Equations (LaTeX)
A system of two linear equations in matrix form:
$$\begin{pmatrix} a_{11} & a_{12} \\ a_{21} & a_{22} \end{pmatrix} \begin{pmatrix} x_1 \\ x_2 \end{pmatrix} = \begin{pmatrix} b_1 \\ b_2 \end{pmatrix}$$

### 8. Geometric Series Sum (Auto-Detected Typst Block)
The sum of an infinite geometric series where $|r| < 1$:
$$ sum_(k=0)^infinity a r^k = a / (1 - r) $$

### 9. Rational Function Decomposition (Auto-Detected SymPy Block)
Partial fraction decomposition of a simple rational expression:
$$ 1 / (x**2 - 1) = 1 / (2*(x - 1)) - 1 / (2*(x + 1)) $$

### 10. Discriminant of Cubic Equation (LaTeX)
The algebraic discriminant that determines the nature of the roots of $ax^3 + bx^2 + cx + d = 0$:
$$\Delta = 18abcd - 4b^3d + b^2c^2 - 4ac^3 - 27a^2d^2$$

---

## Calculus Equations

### 11. Limit Definition of Derivative (Auto-Detected Typst Block)
Represents the instantaneous rate of change of a function:
$$ f'(x) = lim_(h -> 0) (f(x+h) - f(x)) / h $$

### 12. Fundamental Theorem of Calculus (LaTeX)
Connects differentiation and integration:
$$\int_{a}^{b} f(x) \, dx = F(b) - F(a)$$

### 13. Integration by Parts (Explicit SymPy Wrapper)
A rule that transforms the integral of a product of functions:
py:{Integral(u * Derivative(v, x), x) = u * v - Integral(v * Derivative(u, x), x)}

### 14. Taylor Series Expansion (Auto-Detected Typst Block)
Represents a function as an infinite sum of terms calculated from its derivatives at a point:
$$ f(x) = sum_(n=0)^infinity (f^((n))(a)) / (n!) (x - a)^n $$

### 15. Gaussian Integral (LaTeX)
The integral of the 1D Gaussian function over the entire real line:
$$\int_{-\infty}^{\infty} e^{-x^2} \, dx = \sqrt{\pi}$$

### 16. Laplace Transform Definition (Explicit SymPy Wrapper)
An integral transform that converts a function of a real variable (time) to a function of a complex variable:
py:{L(f(t)) = Integral(exp(-s*t) * f(t), (t, 0, oo))}

### 17. Gradient of a Scalar Field (Explicit Typst Wrapper)
The vector of partial derivatives of a scalar function in Cartesian coordinates:
t:{grad phi = (partial phi) / (partial x) bold(i) + (partial phi) / (partial y) bold(j) + (partial phi) / (partial z) bold(k)}

### 18. Fourier Series Expansion (LaTeX)
Represents a periodic function as a sum of sine and cosine terms:
$$f(x) = \frac{a_0}{2} + \sum_{n=1}^{\infty} \left( a_n \cos\left(\frac{2\pi n x}{T}\right) + b_n \sin\left(\frac{2\pi n x}{T}\right) \right)$$

### 19. Chain Rule for Multi-Variable Functions (Auto-Detected Typst Block)
The partial derivative chain rule for a function $z(x, y)$ where $x$ and $y$ are functions of $u$:
$$ (partial z) / (partial u) = (partial z) / (partial x) (partial x) / (partial u) + (partial z) / (partial y) (partial y) / (partial u) $$

### 20. First-Order Linear Differential Equation (Auto-Detected SymPy Block)
General solution for the differential equation $y' + P(x)y = Q(x)$:
$$ y(x) = exp(-Integral(P(x), x)) * (Integral(Q(x) * exp(Integral(P(x), x)), x) + C) $$
