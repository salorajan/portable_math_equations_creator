# Copyright (c) 2026 MyCompany LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
Markdown to MS Word and HTML converter with equation parsing and conversion support.
Converts SymPy and MathML formulas in a Markdown file into standard LaTeX,
then uses Pandoc to compile to MS Word (.docx) or HTML (with MathML formatting).

Features:
- Portable Pandoc detection: Looks for a local pandoc.exe file in the script's folder,
  a "bin" subfolder, or a "pandoc" subfolder, before falling back to system PATH.
"""

import os
import sys
import re
import subprocess
import tempfile
import sympy
import html
import docx
import json
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def convert_docx_to_pdf_via_word(docx_path: str, pdf_path: str) -> bool:
    """Converts a DOCX file to PDF using Microsoft Word COM automation if available."""
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        
        docx_abs = os.path.abspath(docx_path)
        pdf_abs = os.path.abspath(pdf_path)
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(docx_abs)
            doc.SaveAs(pdf_abs, FileFormat=17)
            doc.Close()
            return True
        except Exception as e:
            print(f"Word COM Conversion Error: {e}", file=sys.stderr)
            return False
        finally:
            word.Quit()
            pythoncom.CoUninitialize()
    except Exception as e:
        return False


def read_file_with_fallback(file_path: str) -> str:
    """Read a file trying UTF-8, CP1252, and Latin-1 encodings successively."""
    encodings = ["utf-8", "cp1252", "latin-1"]
    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    # If all fail, fallback to default reading with error replacement
    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def make_mathml_accessible(html_content: str) -> str:
    """
    Finds <math> elements in HTML, wraps block math in scrollable containers, 
    but leaves the math element clean and navigable for screen readers.
    Strips redundant aria-label and alttext to prevent double-reading.
    """
    # Regex to find <math> tags and their contents
    math_pattern = re.compile(r'(<math([^>]*)>)(.*?)(</math>)', re.DOTALL)
    
    def replace_math(match):
        full_open_tag = match.group(1)
        tag_attrs = match.group(2)
        inner_content = match.group(3)
        close_tag = match.group(4)
        
        display_type = "inline"
        if 'display="block"' in tag_attrs:
            display_type = "block"
            
        # Strip redundant aria-label and alttext to prevent double-reading
        clean_attrs = re.sub(r'\b(aria-label|alttext)="[^"]*"', '', tag_attrs)
        # Normalize double spaces
        clean_attrs = re.sub(r'\s+', ' ', clean_attrs).strip()
        
        new_open_tag = f'<math {clean_attrs}>' if clean_attrs else '<math>'
        math_block = f'{new_open_tag}{inner_content}{close_tag}'
        
        if display_type == "block":
            return f'<div class="math-container">{math_block}</div>'
        return math_block
        
    return math_pattern.sub(replace_math, html_content)


def strip_semantics_and_annotations(html_content: str) -> str:
    """
    Strips <semantics> and <annotation> tags from MathML in HTML content,
    replacing them with their clean Presentation MathML child (usually <mrow>).
    """
    from lxml import etree
    try:
        parser = etree.HTMLParser()
        tree = etree.fromstring(html_content.encode('utf-8'), parser=parser)
        
        # Find all <semantics> elements
        semantics_elements = tree.xpath('//semantics') or tree.xpath('//*[local-name()="semantics"]')
        
        for semantics in semantics_elements:
            parent = semantics.getparent()
            if parent is not None:
                if len(semantics) > 0:
                    presentation_math = semantics[0]
                    idx = parent.index(semantics)
                    parent.insert(idx, presentation_math)
                    parent.remove(semantics)
                    
        result = etree.tostring(tree, encoding='unicode', method='html')
        return result
    except Exception as e:
        print(f"Warning: Failed to strip semantics/annotations: {e}", file=sys.stderr)
        return html_content


def ensure_mathjax_installed() -> str:
    """
    Checks if MathJax is installed locally in references/mathjax/.
    If not, prompts the user to download it.
    Returns the script source path (local path if installed, CDN URL if not installed/declined).
    """
    local_dir = os.path.join("references", "mathjax")
    local_path = os.path.join(local_dir, "tex-mml-chtml.js")
    
    if os.path.exists(local_path):
        return "references/mathjax/tex-mml-chtml.js"
        
    cdn_url = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"
    
    # Try to prompt the user if in an interactive terminal
    if sys.stdin.isatty():
        try:
            print("\nMathJax is recommended for rendering equations beautifully in all browsers.")
            print("Would you like to download and install MathJax locally for offline use? [y/N]: ", end="", flush=True)
            choice = input().strip().lower()
            if choice in ('y', 'yes'):
                print("Downloading MathJax from CDN...", flush=True)
                os.makedirs(local_dir, exist_ok=True)
                import urllib.request
                urllib.request.urlretrieve(cdn_url, local_path)
                print("Successfully installed MathJax locally!", flush=True)
                return "references/mathjax/tex-mml-chtml.js"
        except Exception as e:
            print(f"Warning: Failed to download MathJax: {e}. Falling back to CDN link.", file=sys.stderr)
            
    return cdn_url


def post_process_html_accessibility(html_file_path: str, title: str):
    """
    Post-processes the generated standalone HTML file to add accessibility features:
    - CSS styling for contrast, text spacing, focus ring, skip link.
    - Skip-to-content links.
    - Wrap body content in a <main id="main-content"> tag.
    - Enrich MathML elements with alttext and aria-labels.
    - Merges table columns (Stage 2/3 table merging).
    - Integrates MathJax for high-quality mathematical rendering.
    """
    content = read_file_with_fallback(html_file_path)
    
    # Enrich MathML
    content = make_mathml_accessible(content)
    
    # Strip semantics and annotation tags to leave strictly clean MathML 3.0 elements
    content = strip_semantics_and_annotations(content)
    
    # Merge table columns
    content = merge_html_table_columns(content)
    
    accessibility_css = """
    /* Screen reader only utility class */
    .sr-only {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    /* Accessible, responsive, and high-contrast typography */
    :root {
      --bg-color: #fcfcf9; /* Warm off-white for reading comfort */
      --text-color: #1e293b; /* Deep slate gray for contrast (exceeds 7:1 contrast) */
      --link-color: #1d4ed8; /* High-contrast blue */
      --focus-ring-color: #2563eb; /* High-contrast blue focus ring */
    }
    
    body {
      background-color: var(--bg-color);
      color: var(--text-color);
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      line-height: 1.65;
      letter-spacing: 0.012em;
      word-spacing: 0.03em;
      max-width: 48rem;
      margin: 0 auto;
      padding: 2rem 1.5rem;
    }

    /* Skip link styles */
    .skip-link {
      position: absolute;
      top: -100px;
      left: 10px;
      background: var(--focus-ring-color);
      color: white;
      padding: 10px 15px;
      z-index: 1000;
      transition: top 0.2s;
      text-decoration: none;
      font-weight: bold;
      border-radius: 4px;
    }
    .skip-link:focus {
      top: 10px;
      outline: none;
    }
    
    /* Headings */
    h1, h2, h3, h4, h5, h6 {
      color: #0f172a;
      margin-top: 1.8em;
      margin-bottom: 0.8em;
      line-height: 1.3;
    }
    
    p, ul, ol, dl, table, blockquote {
      margin-bottom: 1.6em; /* Paragraph spacing */
    }
    
    a {
      color: var(--link-color);
      text-decoration: underline;
      text-underline-offset: 3px;
    }
    
    a:hover {
      color: #1e40af;
    }
    
    /* Clear Visual Keyboard Focus Indicators (WCAG 2.4.7 / 2.2 Focus Appearance) */
    a:focus-visible, button:focus-visible, [tabindex="0"]:focus-visible {
      outline: 3px solid var(--focus-ring-color) !important;
      outline-offset: 2px !important;
      border-radius: 2px;
    }
    
    /* Responsive Math Tables & Scrollable containers (WCAG 1.4.10 Reflow) */
    .math-container {
      overflow-x: auto;
      margin: 1.6em 0;
      padding: 0.8em;
      border-radius: 4px;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
    }
    .math-container:focus-visible {
      outline: 3px solid var(--focus-ring-color);
      outline-offset: 0;
    }
    
    /* Tables */
    table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 2em;
    }
    th, td {
      padding: 0.75rem;
      text-align: left;
      border-bottom: 1px solid #e2e8f0;
    }
    th {
      background-color: #f1f5f9;
      font-weight: 600;
      color: #0f172a;
    }
    
    /* Blockquotes */
    blockquote {
      border-left: 4px solid #cbd5e1;
      padding-left: 1rem;
      color: #475569;
      margin-left: 0;
      margin-right: 0;
    }
    """
    
    mathjax_src = ensure_mathjax_installed()
    mathjax_script = f"""
    <script>
    window.MathJax = {{
      options: {{
        enableMenu: true,
        renderAtStart: true
      }},
      chtml: {{
        displayAlign: 'center'
      }},
      loader: {{
        load: ['a11y/semantic-enrich', 'a11y/explorer']
      }}
    }};
    </script>
    <script src="{mathjax_src}" id="MathJax-script" async></script>
    """
    
    style_block = f"<style>\n{accessibility_css}\n</style>\n{mathjax_script}\n"
    if "</head>" in content:
        content = content.replace("</head>", f"{style_block}</head>")
    else:
        content = style_block + content
        
    body_match = re.search(r'(<body[^>]*>)', content)
    if body_match:
        body_open_tag = body_match.group(1)
        skip_link_html = f'\n  <a href="#main-content" class="skip-link">Skip to main content</a>\n'
        content = content.replace(body_open_tag, f'{body_open_tag}{skip_link_html}<main id="main-content">', 1)
        
        if "</body>" in content:
            content = content.replace("</body>", "</main>\n</body>", 1)
        else:
            content = content + "\n</main>"
            
    with open(html_file_path, "w", encoding="utf-8") as f:
        f.write(content)


def normalize_markdown_tables(md_content: str) -> str:
    """
    Scans markdown content for pipe tables and ensures every row in a table
    has the exact same number of columns (cells) as the header row.
    Pads shorter rows with empty cells.
    Also ensures there is a blank line before the table start.
    """
    lines = md_content.splitlines()
    in_table = False
    table_lines = []
    output_lines = []
    
    def flush_table(t_lines):
        if not t_lines:
            return []
            
        processed_rows = []
        max_cols = 0
        
        for line in t_lines:
            stripped = line.strip()
            if stripped.startswith("|") and stripped.endswith("|"):
                cells = [c.strip() for c in stripped.split("|")[1:-1]]
                processed_rows.append(cells)
                max_cols = max(max_cols, len(cells))
            else:
                processed_rows.append(None)
                
        new_table_lines = []
        for idx, cells in enumerate(processed_rows):
            original_line = t_lines[idx]
            if cells is None:
                new_table_lines.append(original_line)
                continue
                
            if len(cells) < max_cols:
                is_separator = all(re.match(r'^:?-+:?$', c) for c in cells if c)
                padding_cell = "---" if is_separator else ""
                cells.extend([padding_cell] * (max_cols - len(cells)))
                
            new_line = "| " + " | ".join(cells) + " |"
            new_table_lines.append(new_line)
            
        return new_table_lines

    for idx, line in enumerate(lines):
        stripped = line.strip()
        is_table_line = stripped.startswith("|") and stripped.endswith("|")
        
        if is_table_line:
            if not in_table:
                # Table is starting! Check if the previous line in output_lines is non-empty
                if output_lines and output_lines[-1].strip() != "":
                    # Insert a blank line to separate table from paragraph (WCAG/Markdown parse fix)
                    output_lines.append("")
                in_table = True
            table_lines.append(line)
        else:
            if in_table:
                output_lines.extend(flush_table(table_lines))
                table_lines = []
                in_table = False
            output_lines.append(line)
            
    if in_table:
        output_lines.extend(flush_table(table_lines))
        
    return "\n".join(output_lines)


def merge_html_table_columns(html_content: str) -> str:
    """
    Finds HTML tables and merges consecutive columns based on content tags.
    - In header rows: Merges if cell is empty, has spaces/&nbsp;, or contains ".." or "colspan".
    - In data rows: Merges ONLY if cell contains ".." or "colspan" (to avoid merging empty data cells).
    """
    table_pattern = re.compile(r'(<table[^>]*>.*?</table>)', re.DOTALL)
    
    def process_table(table_match):
        table_html = table_match.group(1)
        
        has_thead = "<thead>" in table_html
        row_pattern = re.compile(r'(<tr[^>]*>.*?</tr>)', re.DOTALL)
        rows = list(row_pattern.finditer(table_html))
        if not rows:
            return table_html
            
        new_rows = []
        
        for idx, row_match in enumerate(rows):
            row_html = row_match.group(1)
            is_header = False
            if has_thead:
                is_header = "<th>" in row_html
            else:
                is_header = (idx == 0)
                
            cell_pattern = re.compile(r'(<(th|td)([^>]*)>)(.*?)(</\2>)', re.DOTALL)
            cells = list(cell_pattern.finditer(row_html))
            if not cells:
                new_rows.append(row_html)
                continue
                
            new_cells = []
            skip_count = 0
            
            for i in range(len(cells)):
                if skip_count > 0:
                    skip_count -= 1
                    continue
                    
                cell_match = cells[i]
                open_tag = cell_match.group(1)
                tag_name = cell_match.group(2)
                attrs_str = cell_match.group(3)
                content = cell_match.group(4).strip()
                close_tag = cell_match.group(5)
                
                colspan = 1
                colspan_match = re.search(r'colspan="(\d+)"', attrs_str)
                if colspan_match:
                    colspan = int(colspan_match.group(1))
                    
                j = i + 1
                while j < len(cells):
                    next_cell = cells[j]
                    next_content = next_cell.group(4).strip()
                    
                    if is_header:
                        should_merge = not next_content or next_content in ("&nbsp;", "..", "colspan")
                    else:
                        should_merge = next_content in ("..", "colspan")
                        
                    if should_merge:
                        colspan += 1
                        j += 1
                        skip_count += 1
                    else:
                        break
                        
                if colspan > 1:
                    if 'colspan=' in attrs_str:
                        new_attrs = re.sub(r'colspan="\d+"', f'colspan="{colspan}"', attrs_str)
                    else:
                        new_attrs = f'{attrs_str} colspan="{colspan}"'
                    new_open_tag = f'<{tag_name}{new_attrs}>'
                    new_cells.append(f'{new_open_tag}{content}{close_tag}')
                else:
                    new_cells.append(cell_match.group(0))
                    
            tr_attrs = re.match(r'^(<tr[^>]*>)', row_html).group(1)
            new_row_html = f'{tr_attrs}\n' + '\n'.join(new_cells) + '\n</tr>'
            new_rows.append(new_row_html)
            
        rebuilt_table = table_html
        for orig_row, new_row in zip(rows, new_rows):
            rebuilt_table = rebuilt_table.replace(orig_row.group(0), new_row, 1)
            
        return rebuilt_table
        
    return table_pattern.sub(process_table, html_content)


def post_process_docx_accessibility(docx_file_path: str, lang: str):
    """
    Applies accessibility enhancements to the generated DOCX file:
    - Merges table columns (Stage 2/3 table merging).
    - Sets table headers to repeat across pages (tblHeader).
    - Prevents table rows from splitting across page breaks (cantSplit).
    """
    doc = docx.Document(docx_file_path)
    
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            is_header = (row_idx == 0)
            
            trPr = row._tr.get_or_add_trPr()
            
            # Repeat Header
            if is_header:
                if not row._tr.xpath('w:trPr/w:tblHeader'):
                    tblHeader = OxmlElement('w:tblHeader')
                    trPr.append(tblHeader)
            
            # Prevent Row Splits
            if not row._tr.xpath('w:trPr/w:cantSplit'):
                cantSplit = OxmlElement('w:cantSplit')
                trPr.append(cantSplit)
                
            # Merge Columns
            i = 0
            while i < len(row.cells):
                cell = row.cells[i]
                content = cell.text.strip()
                
                j = i + 1
                colspan = 1
                
                while j < len(row.cells):
                    next_cell = row.cells[j]
                    next_content = next_cell.text.strip()
                    
                    if is_header:
                        should_merge = not next_content or next_content in ("..", "colspan")
                    else:
                        should_merge = next_content in ("..", "colspan")
                        
                    if should_merge:
                        colspan += 1
                        j += 1
                    else:
                        break
                        
                if colspan > 1:
                    # Clear cells text before merging to prevent duplicate text concatenation
                    for k in range(i + 1, j):
                        row.cells[k].text = ""
                    target_cell = row.cells[i + colspan - 1]
                    cell.merge(target_cell)
                    i = j
                else:
                    i += 1
                    
    doc.save(docx_file_path)


def parse_yaml_frontmatter(content: str) -> tuple:
    """
    Parses YAML frontmatter at the beginning of the markdown content.
    Returns a tuple of (metadata_dict, remaining_content_str).
    """
    metadata = {}
    lines = content.splitlines()
    if not lines:
        return metadata, content
        
    first_line_idx = 0
    while first_line_idx < len(lines) and not lines[first_line_idx].strip():
        first_line_idx += 1
        
    if first_line_idx < len(lines) and lines[first_line_idx].strip() == "---":
        end_idx = -1
        for i in range(first_line_idx + 1, len(lines)):
            if lines[i].strip() == "---":
                end_idx = i
                break
        if end_idx != -1:
            for i in range(first_line_idx + 1, end_idx):
                line = lines[i]
                if ":" in line:
                    key, val = line.split(":", 1)
                    key = key.strip().lower()
                    val = val.strip().strip('"').strip("'")
                    metadata[key] = val
            remaining_content = "\n".join(lines[end_idx + 1:])
            return metadata, remaining_content
            
    return metadata, content


def validate_markdown_accessibility(content: str) -> bool:
    """
    Validates markdown content against WCAG 2.2 accessibility rules.
    Checks for:
    - Proper heading hierarchy (no skipped levels, e.g. H1 directly to H3)
    - Image alt text presence (no empty alt text unless marked decorative)
    - Descriptive link labels (avoid "click here", raw URLs)
    Returns True if no warnings, False if warnings were found.
    """
    warnings = []
    lines = content.splitlines()
    
    last_heading_level = 0
    in_code_block = False
    
    # Regexes
    image_pattern = re.compile(r'!\[(.*?)\]\((.*?)\)')
    link_pattern = re.compile(r'(?<!\!)\[(.*?)\]\((.*?)\)')
    
    generic_link_words = {"click here", "read more", "link", "url", "here", "website", "page", "more", "go"}
    
    for idx, line in enumerate(lines, 1):
        stripped = line.strip()
        
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
            
        if in_code_block:
            continue
            
        # 1. Heading hierarchy check
        if stripped.startswith("#"):
            m = re.match(r'^(#+)\s+(.*)', stripped)
            if m:
                level = len(m.group(1))
                heading_text = m.group(2)
                if last_heading_level > 0 and level > last_heading_level + 1:
                    warnings.append(
                        f"Line {idx}: Heading level skipped. Heading '{heading_text}' is level {level} "
                        f"but previous heading was level {last_heading_level}. Headings must form a nested structure."
                    )
                last_heading_level = level
                
        # 2. Image Alt Text Check
        for match in image_pattern.finditer(line):
            alt_text = match.group(1).strip()
            image_url = match.group(2).strip()
            if not alt_text or alt_text.lower() in ("image", "picture", "photo", "graphic"):
                warnings.append(
                    f"Line {idx}: Image '{image_url}' is missing a descriptive alt text. "
                    f"Empty alt text or generic labels like '{alt_text}' are not accessible for screen readers."
                )
                
        # 3. Link Text Check
        for match in link_pattern.finditer(line):
            link_text = match.group(1).strip()
            link_url = match.group(2).strip()
            
            # Clean up link text from nested bracket formatting and Pandoc attributes (like {.underline})
            link_text_clean = link_text.strip('[]')
            link_text_clean = re.sub(r'\{\.[a-zA-Z0-9_-]+\}', '', link_text_clean)
            
            text_lower = link_text_clean.lower()
            cleaned_text = re.sub(r'[^\w\s]', '', text_lower).strip()
            
            if not cleaned_text or cleaned_text in generic_link_words or re.match(r'^https?://', link_text_clean):
                warnings.append(
                    f"Line {idx}: Link to '{link_url}' uses non-descriptive text '{link_text}'. "
                    f"Use descriptive words that explain where the link goes instead of '{link_text}'."
                )

    if warnings:
        print("\n" + "="*70)
        print(" ACCESSIBILITY WARNINGS (WCAG 2.2 AUDIT)")
        print("="*70)
        for warn in warnings:
            print(f"  [!] {warn}")
        print("="*70 + "\n")
        return False
        
    return True


def fix_missing_alt_texts(content: str) -> str:
    """
    Finds images in the markdown content and, if alt text is missing or generic,
    replaces it with 'this image in this {line} is not alt text . put it'.
    """
    lines = content.splitlines()
    new_lines = []
    image_pattern = re.compile(r'!\[(.*?)\]\((.*?)\)')
    
    for idx, line in enumerate(lines, 1):
        def replace_img(match):
            alt_text = match.group(1).strip()
            image_url = match.group(2).strip()
            if not alt_text or alt_text.lower() in ("image", "picture", "photo", "graphic", "logo"):
                new_alt = f"this image in this {idx} is not alt text . put it"
                return f"![{new_alt}]({image_url})"
            return match.group(0)
            
        new_line = image_pattern.sub(replace_img, line)
        new_lines.append(new_line)
        
    return "\n".join(new_lines)


def unescape_escaped_math(content: str) -> str:
    """
    Finds escaped math symbols (like \\$\\$ or \\$) in intermediate markdown
    created by Pandoc, and unescapes them and their internal characters
    (like \\_ and \\^) so the equation parser can detect them.
    """
    # 1. Match escaped block math: \$\$(.*?)\$\$
    block_pattern = re.compile(r'\\\$\\\$(.*?)\\\$\\\$', re.DOTALL)
    
    def replace_block(match):
        inner = match.group(1)
        # Unescape common markdown escapes
        inner = inner.replace(r'\_', '_').replace(r'\^', '^')
        inner = inner.replace(r'\{', '{').replace(r'\}', '}')
        inner = inner.replace(r'\*', '*').replace(r'\[', '[').replace(r'\]', ']')
        inner = inner.replace('\\\\', '\\')
        return f"$$\n{inner.strip()}\n$$"
        
    content = block_pattern.sub(replace_block, content)
    
    # 2. Match escaped inline math: \$(.*?)\$
    inline_pattern = re.compile(r'\\\$(.*?)\\\$', re.DOTALL)
    
    def replace_inline(match):
        inner = match.group(1)
        inner = inner.replace(r'\_', '_').replace(r'\^', '^')
        inner = inner.replace(r'\{', '{').replace(r'\}', '}')
        inner = inner.replace(r'\*', '*').replace(r'\[', '[').replace(r'\]', ']')
        inner = inner.replace('\\\\', '\\')
        return f"${inner.strip()}$"
        
    content = inline_pattern.sub(replace_inline, content)
    
    return content


# --- Portable Pandoc Path Resolution ---

def get_pandoc_command() -> str:
    """
    Returns the path to Pandoc.
    Checks for a local 'pandoc.exe' (on Windows) or 'pandoc' in:
      1. The temporary extraction folder (sys._MEIPASS) if running inside a PyInstaller bundle
      2. The directory of this script (or executable)
      3. A 'bin' subdirectory
      4. A 'pandoc' subdirectory
    Falls back to 'pandoc' (which searches the system PATH) if not found locally.
    """
    search_dirs = []
    
    # If running inside a PyInstaller bundle, check sys._MEIPASS first
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        search_dirs.append(sys._MEIPASS)
        search_dirs.append(os.path.join(sys._MEIPASS, "bin"))
        search_dirs.append(os.path.join(sys._MEIPASS, "pandoc"))
        
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Standard executables names to scan
    exe_names = ["pandoc.exe", "pandoc"]
    
    # Locations to search
    search_dirs.extend([
        script_dir,
        os.path.join(script_dir, "bin"),
        os.path.join(script_dir, "pandoc")
    ])
    
    for search_dir in search_dirs:
        for name in exe_names:
            candidate_path = os.path.join(search_dir, name)
            if os.path.exists(candidate_path) and os.path.isfile(candidate_path):
                return candidate_path
                
    # Fallback to system command
    return "pandoc"


def get_typst_command() -> str:
    """
    Returns the path to Typst.
    Checks for a local 'typst.exe' (on Windows) or 'typst' in:
      1. The temporary extraction folder (sys._MEIPASS) if running inside a PyInstaller bundle
      2. The directory of this script (or executable)
      3. A 'bin' subdirectory
    Falls back to 'typst' (which searches the system PATH) if not found locally.
    """
    search_dirs = []
    
    # If running inside a PyInstaller bundle, check sys._MEIPASS first
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        search_dirs.append(sys._MEIPASS)
        search_dirs.append(os.path.join(sys._MEIPASS, "bin"))
        
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Standard executables names to scan
    exe_names = ["typst.exe", "typst"]
    
    # Locations to search
    search_dirs.extend([
        script_dir,
        os.path.join(script_dir, "bin")
    ])
    
    for search_dir in search_dirs:
        for name in exe_names:
            candidate_path = os.path.join(search_dir, name)
            if os.path.exists(candidate_path) and os.path.isfile(candidate_path):
                return candidate_path
                
    # Fallback to system command
    return "typst"


def check_typst_installed():
    """Verify that Typst is installed and accessible."""
    cmd = get_typst_command()
    try:
        subprocess.run([cmd, "--version"], capture_output=True)
        return True
    except FileNotFoundError:
        return False


def get_tidy_command() -> str:
    """
    Returns the path to HTML Tidy.
    Checks for a local 'tidy.exe' (on Windows) or 'tidy' in:
      1. The temporary extraction folder (sys._MEIPASS) if running inside a PyInstaller bundle
      2. The directory of this script (or executable)
      3. A 'bin' subdirectory
    Falls back to 'tidy' (which searches the system PATH) if not found locally.
    """
    search_dirs = []
    
    # If running inside a PyInstaller bundle, check sys._MEIPASS first
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        search_dirs.append(sys._MEIPASS)
        search_dirs.append(os.path.join(sys._MEIPASS, "bin"))
        
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
    # Standard executables names to scan
    exe_names = ["tidy.exe", "tidy"]
    
    # Locations to search
    search_dirs.extend([
        script_dir,
        os.path.join(script_dir, "bin")
    ])
    
    for search_dir in search_dirs:
        for name in exe_names:
            candidate_path = os.path.join(search_dir, name)
            if os.path.exists(candidate_path) and os.path.isfile(candidate_path):
                return candidate_path
                
    # Fallback to system command
    return "tidy"


def find_browser_driver_manager_paths() -> tuple:
    """
    Search in ~/.browser-driver-manager to find Chrome and ChromeDriver executables.
    Returns a tuple of (chrome_path, chromedriver_path) or (None, None).
    """
    home = os.path.expanduser("~")
    base_dir = os.path.join(home, ".browser-driver-manager")
    if not os.path.exists(base_dir):
        return None, None
        
    chrome_path = None
    chromedriver_path = None
    
    # Walk the directory to find chrome.exe and chromedriver.exe
    for root, dirs, files in os.walk(base_dir):
        if "chrome.exe" in files and not chrome_path:
            chrome_path = os.path.join(root, "chrome.exe")
        if "chromedriver.exe" in files and not chromedriver_path:
            chromedriver_path = os.path.join(root, "chromedriver.exe")
            
    return chrome_path, chromedriver_path


def parse_and_print_axe_report(report_path: str):
    """Parses and prints the Axe accessibility violations report."""
    if not os.path.exists(report_path):
        return
        
    try:
        with open(report_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"Warning: Failed to parse Axe report JSON: {e}")
        return
        
    if not data or not isinstance(data, list):
        return
        
    violations = data[0].get("violations", [])
    if not violations:
        print("[OK] Axe Accessibility Audit: No violations detected!")
        return
        
    print("\n" + "="*70)
    print(" AXE ACCESSIBILITY VIOLATIONS (WCAG 2.2 AUDIT)")
    print("="*70)
    
    total_occurrences = 0
    for idx, v in enumerate(violations, 1):
        v_id = v.get("id", "unknown")
        impact = v.get("impact", "unknown")
        help_text = v.get("help", "")
        help_url = v.get("helpUrl", "")
        nodes = v.get("nodes", [])
        
        occurrences = len(nodes)
        total_occurrences += occurrences
        
        print(f"  [{idx}] Violation: {v_id} (Impact: {impact})")
        print(f"      Help: {help_text}")
        if help_url:
            print(f"      Rule Info: {help_url}")
        print("      Affects:")
        for node in nodes:
            target = ", ".join(node.get("target", []))
            html_snippet = node.get("html", "").strip()
            # Truncate html snippet if too long
            if len(html_snippet) > 120:
                html_snippet = html_snippet[:117] + "..."
            print(f"        - Selector: {target}")
            print(f"          HTML:     {html_snippet}")
        print("-" * 70)
        
    print(f"  Total: {len(violations)} rule violation(s) with {total_occurrences} occurrence(s) detected.")
    print("="*70 + "\n")


def run_automated_accessibility_checks(html_file_path: str):
    """
    Runs automated validation checks on the output HTML:
    1. HTML Tidy for structural validation.
    2. Axe CLI for WCAG 2.2 compliance checking.
    """
    print("\nStarting automated accessibility and structural validation...")
    
    # 1. HTML Tidy Check
    tidy_bin = get_tidy_command()
    try:
        res = subprocess.run([tidy_bin, "-errors", "-quiet", html_file_path], capture_output=True, text=True, encoding="utf-8")
        output = (res.stdout or "") + (res.stderr or "")
        if output.strip():
            print("\n" + "="*70)
            print(" HTML TIDY VALIDATION ISSUES")
            print("="*70)
            print(output.strip())
            print("="*70 + "\n")
        else:
            print("[OK] HTML Tidy: No structural errors found.")
    except FileNotFoundError:
        print("Warning: HTML Tidy ('tidy.exe') is not installed locally or globally. Skipping structure validation.")
    except Exception as e:
        print(f"Warning: HTML Tidy validation failed: {e}")
        
    # 2. Axe Core Check
    chrome_path, chromedriver_path = find_browser_driver_manager_paths()
    
    # Generate relative temporary filename to avoid absolute path issue in Axe CLI --save option on Windows
    report_filename = "_axe_report_temp.json"
    html_dir = os.path.dirname(os.path.abspath(html_file_path))
    report_path = os.path.join(html_dir, report_filename)
        
    try:
        from urllib.request import pathname2url
        file_url = "file:" + pathname2url(os.path.abspath(html_file_path))
        
        # Build command. Since Axe CLI resolves the --save value relative to CWD,
        # we run it with CWD set to html_dir and pass the relative report filename.
        npx_cmd = "npx.cmd" if os.name == "nt" else "npx"
        cmd = [npx_cmd, "--yes", "@axe-core/cli", file_url, "--save", report_filename]
        if chrome_path and chromedriver_path:
            cmd += [
                "--chrome-path", chrome_path,
                "--chromedriver-path", chromedriver_path
            ]
            
        print("Running Axe CLI accessibility audit...")
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", cwd=html_dir)
        if os.path.exists(report_path) and os.path.getsize(report_path) > 0:
            parse_and_print_axe_report(report_path)
        else:
            print("Warning: Axe CLI did not generate a report.")
            if res.stderr:
                print(f"Axe CLI stderr: {res.stderr.strip()}")
            if res.stdout:
                print(f"Axe CLI stdout: {res.stdout.strip()}")
    except Exception as e:
        print(f"Warning: Axe CLI accessibility audit failed to run: {e}")
    finally:
        if os.path.exists(report_path):
            try:
                os.remove(report_path)
            except Exception:
                pass


# --- Tokenizer and Equation parsing functions ---

def tokenize_text(text: str) -> list:
    """Tokenize text into plain text runs and equation markup spans."""
    tokens = []
    pos = 0
    n = len(text)
    while pos < n:
        # Check for code blocks and code spans to avoid tokenizing equations inside them
        if text.startswith("```", pos):
            end = text.find("```", pos + 3)
            if end != -1:
                end += 3
                val = text[pos:end]
                if tokens and tokens[-1][0] == "text":
                    tokens[-1] = ("text", tokens[-1][1] + val)
                else:
                    tokens.append(("text", val))
                pos = end
                continue
        elif text.startswith("`", pos):
            end = text.find("`", pos + 1)
            if end != -1:
                end += 1
                val = text[pos:end]
                if tokens and tokens[-1][0] == "text":
                    tokens[-1] = ("text", tokens[-1][1] + val)
                else:
                    tokens.append(("text", val))
                pos = end
                continue

        # 1. Try to match MathML
        if text.startswith("<math", pos):
            end = text.find("</math>", pos)
            if end != -1:
                end += len("</math>")
                tokens.append(("mathml", text[pos:end]))
                pos = end
                continue
                
        # 2. Try to match LaTeX display/block $$ ... $$
        if text.startswith("$$", pos):
            end = text.find("$$", pos + 2)
            if end != -1:
                end += 2
                tokens.append(("latex_block", text[pos:end]))
                pos = end
                continue
                
        # 3. Try to match LaTeX inline $ ... $
        if text[pos] == "$":
            # Ensure it's not escaped by a backslash
            if pos > 0 and text[pos-1] == "\\":
                # Escaped dollar sign, treat as normal text
                pass
            else:
                end = text.find("$", pos + 1)
                if end != -1:
                    tokens.append(("latex_inline", text[pos:end+1]))
                    pos = end + 1
                    continue
                    
        # 4. Try to match SymPy sympy:{ ... } or py:{ ... }
        if text.startswith("sympy:{", pos):
            end = text.find("}", pos + 7)
            if end != -1:
                tokens.append(("sympy", text[pos:end+1]))
                pos = end + 1
                continue
        if text.startswith("py:{", pos):
            end = text.find("}", pos + 4)
            if end != -1:
                tokens.append(("sympy", text[pos:end+1]))
                pos = end + 1
                continue
                
        # Try to match Typst math typst:{ ... } or t:{ ... }
        if text.startswith("typst:{", pos):
            end = text.find("}", pos + 7)
            if end != -1:
                tokens.append(("typst", text[pos:end+1]))
                pos = end + 1
                continue
        if text.startswith("t:{", pos):
            end = text.find("}", pos + 3)
            if end != -1:
                tokens.append(("typst", text[pos:end+1]))
                pos = end + 1
                continue
                
        # 5. Try to match LaTeX/math explicit blocks latex:{ ... } or math:{ ... }
        if text.startswith("latex:{", pos):
            end = text.find("}", pos + 7)
            if end != -1:
                tokens.append(("latex_inline", text[pos:end+1]))
                pos = end + 1
                continue
        if text.startswith("math:{", pos):
            end = text.find("}", pos + 6)
            if end != -1:
                tokens.append(("latex_inline", text[pos:end+1]))
                pos = end + 1
                continue
                
        # If no math matches, consume 1 character as text
        char = text[pos]
        if tokens and tokens[-1][0] == "text":
            tokens[-1] = ("text", tokens[-1][1] + char)
        else:
            tokens.append(("text", char))
        pos += 1
        
    return tokens


def str_to_sympy(sympy_str: str):
    """Parse standard SymPy text string into a SymPy expression object or relation."""
    sympy_str = sympy_str.strip()
    
    # Custom local dictionary to resolve conflicting names like Q, P, y
    local_dict = {
        'Q': sympy.Function('Q'),
        'P': sympy.Function('P'),
        'y': sympy.Function('y')
    }
    
    # Define relational operators and their corresponding SymPy classes
    operators = [
        (">=", sympy.Ge),
        ("<=", sympy.Le),
        ("=", sympy.Eq),
        (">", sympy.Gt),
        ("<", sympy.Lt)
    ]
    for op_str, op_class in operators:
        if op_str in sympy_str:
            parts = sympy_str.split(op_str, 1)
            try:
                lhs = sympy.sympify(parts[0].strip(), locals=local_dict)
                rhs = sympy.sympify(parts[1].strip(), locals=local_dict)
                return op_class(lhs, rhs, evaluate=False)
            except Exception:
                pass
                
    return sympy.sympify(sympy_str, locals=local_dict)


def run_pandoc_conversion(input_data: str, from_fmt: str, to_fmt: str, extra_args: list = []) -> str:
    """Run pandoc with stdin and stdout."""
    cmd = [get_pandoc_command(), "-f", from_fmt, "-t", to_fmt] + extra_args
    res = subprocess.run(cmd, input=input_data, capture_output=True, text=True, encoding="utf-8")
    if res.returncode != 0:
        raise RuntimeError(f"Pandoc execution failed: {res.stderr}")
    return res.stdout


def mathml_to_latex(mathml_str: str) -> str:
    """Convert MathML to LaTeX using Pandoc."""
    if "xmlns" not in mathml_str:
        mathml_str = mathml_str.replace("<math", '<math xmlns="http://www.w3.org/1998/Math/MathML"')
    try:
        # Pandoc parses mathml in html format and outputs markdown math
        markdown_out = run_pandoc_conversion(mathml_str, "html", "markdown").strip()
        if markdown_out.startswith("$$") and markdown_out.endswith("$$"):
            return markdown_out[2:-2].strip()
        elif markdown_out.startswith("$") and markdown_out.endswith("$"):
            return markdown_out[1:-1].strip()
        return markdown_out
    except Exception as e:
        raise RuntimeError(f"Failed to convert MathML to LaTeX: {e}")


def convert_sympy_to_latex(sympy_str: str) -> str:
    """Convert a SymPy expression string to standard LaTeX representation."""
    expr = str_to_sympy(sympy_str)
    return sympy.latex(expr)


def typst_to_latex(typst_math: str) -> str:
    """Convert Typst math syntax to LaTeX math syntax using Pandoc."""
    # Preprocess Typst math to replace shorthands that Pandoc's Typst reader fails to parse correctly
    # Replace common user shorthands like hbar and oint
    typst_math = re.sub(r'\bhbar\b|hbar(?=_)', ' planck.reduce ', typst_math)
    typst_math = re.sub(r'\boint\b|oint(?=_)', ' integral.cont ', typst_math)
    typst_math = re.sub(r'\bgrad\b', ' nabla ', typst_math)
    
    # Replace '+-' with 'plus.minus' and '-+' with 'minus.plus'
    typst_math = typst_math.replace("+-", " plus.minus ").replace("-+", " minus.plus ")
    
    if not (typst_math.startswith("$") and typst_math.endswith("$")):
        typst_math_input = f"${typst_math}$"
    else:
        typst_math_input = typst_math
        
    try:
        markdown_out = run_pandoc_conversion(typst_math_input, "typst", "markdown").strip()
        s = markdown_out
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            s = s[1:-1].strip()
            
        if s.startswith("$$") and s.endswith("$$"):
            return s[2:-2].strip()
        elif s.startswith("$") and s.endswith("$"):
            return s[1:-1].strip()
        elif s.startswith("\\(") and s.endswith("\\)"):
            return s[2:-2].strip()
        elif s.startswith("\\[") and s.endswith("\\]"):
            return s[2:-2].strip()
        return s
    except Exception as e:
        raise RuntimeError(f"Failed to convert Typst math to LaTeX: {e}")
 
 
def detect_math_format(formula: str) -> str:
    """
    Detects the mathematical notation dialect of the formula.
    Returns "latex", "typst", or "sympy".
    """
    formula = formula.strip()
    if not formula:
        return "latex"
        
    # If it contains LaTeX backslashes, it is definitely LaTeX
    if "\\" in formula:
        return "latex"
        
    # If it contains LaTeX subscript/superscript grouping, it is LaTeX
    if "_{" in formula or "^{" in formula:
        return "latex"
        
    # If it contains SymPy characteristics:
    # - ** (exponent in Python/SymPy)
    # - Eq(...) (relation in SymPy)
    # - symbols(...)
    # - Common functions: Matrix(, Integral(, Derivative(, Limit(, exp(
    sympy_indicators = [
        "**", "Eq(", "symbols(", "Matrix(", "Integral(", "Derivative(", "Limit(", "exp(", "oo"
    ]
    if any(ind in formula for ind in sympy_indicators):
        return "sympy"
        
    # If it contains common Typst/typeset syntax
    # - +- (plus-minus)
    # - ->, <-, =>, <= (arrows, comparisons)
    # - Common functions: bold(), vec(), mat(), op()
    # - Key Typst constants/symbols: planck, reduce, nabla, partial, infinity, integral
    typst_indicators = [
        "+-", "->", "<-", "=>", "bold(", "vec(", "mat(", "op(", "dif ", 
        "nabla", "partial", "infinity", "planck", "reduce", "sum_", "lim_", "integral_", "binom(", "root("
    ]
    # Check for common Greek letter words
    greek_letters = {
        "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta", "iota", "kappa",
        "lambda", "mu", "nu", "xi", "omicron", "pi", "rho", "sigma", "tau", "upsilon", "phi",
        "chi", "psi", "omega", "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma",
        "Upsilon", "Phi", "Psi", "Omega"
    }
    
    # Split by non-alphanumeric to find words
    words = set(re.findall(r'[a-zA-Z]+', formula))
    
    if any(ind in formula for ind in typst_indicators) or not words.isdisjoint(greek_letters):
        return "typst"
        
    # If it has simple multiplication like * (but not **), check if it's SymPy or Typst
    if "*" in formula and "**" not in formula:
        # Check if it has python variables (e.g. 4*a*c or 2*a)
        if re.search(r'[a-zA-Z0-9]\*[a-zA-Z]', formula) or re.search(r'[a-zA-Z]\*[0-9]', formula):
            return "sympy"
            
    # Default to typst if it uses standard typst operators like / without backslashes
    if "/" in formula:
        return "typst"
        
    return "latex"


def process_markdown_equations(content: str, metadata: dict = {}) -> str:
    """Find and replace all SymPy, MathML, and Typst formulas in markdown with standard LaTeX math."""
    math_format = metadata.get("math", "auto").strip().lower()
    tokens = tokenize_text(content)
    result = []
    current_line = 1
    
    strip_leading_spaces = False
    
    for i, (token_type, val) in enumerate(tokens):
        token_lines = val.count("\n")
        
        # Apply leading space stripping from block math styling on preceding tokens
        if strip_leading_spaces:
            if token_type == "text":
                val = re.sub(r'^[ \t]*', '', val)
            strip_leading_spaces = False
            
        if token_type == "text":
            result.append(val)
        elif token_type in ("latex_inline", "latex_block"):
            # Extract raw formula
            pure_formula = val
            if val.startswith("$$") and val.endswith("$$"):
                pure_formula = val[2:-2].strip()
            elif val.startswith("$") and val.endswith("$"):
                pure_formula = val[1:-1].strip()
                
            # Classify formula
            if math_format == "typst":
                detected_format = "typst"
            elif math_format == "sympy":
                detected_format = "sympy"
            elif math_format == "latex":
                detected_format = "latex"
            else:  # "auto" or unrecognized
                detected_format = detect_math_format(pure_formula)
                
            is_block = token_type == "latex_block"
            
            # If it is determined to be Typst or SymPy, convert it to LaTeX
            if detected_format == "typst":
                try:
                    latex_eq = typst_to_latex(pure_formula)
                    if is_block:
                        result.append(f"$$\n{latex_eq}\n$$")
                    else:
                        result.append(f"${latex_eq}$")
                except Exception as e:
                    print(f"Warning: Failed to convert Typst formula '{pure_formula[:30]}' at line {current_line}: {e}", file=sys.stderr)
                    result.append(val)
            elif detected_format == "sympy":
                try:
                    latex_eq = convert_sympy_to_latex(pure_formula)
                    if is_block:
                        result.append(f"$$\n{latex_eq}\n$$")
                    else:
                        result.append(f"${latex_eq}$")
                except Exception as e:
                    print(f"Warning: Failed to convert SymPy formula '{pure_formula[:30]}' at line {current_line}: {e}", file=sys.stderr)
                    result.append(val)
            else:
                # Keep as LaTeX
                result.append(val)
                
        elif token_type in ("sympy", "mathml", "typst"):
            # Determine if this token is on its own line (block equation)
            is_start = False
            if i == 0:
                is_start = True
            elif tokens[i-1][0] == "text":
                m = re.search(r'(?:\r?\n|^)([ \t]*)$', tokens[i-1][1])
                if m:
                    is_start = True
                    
            is_end = False
            if i == len(tokens) - 1:
                is_end = True
            elif tokens[i+1][0] == "text":
                m = re.match(r'^([ \t]*)(?:\r?\n|$)', tokens[i+1][1])
                if m:
                    is_end = True
            
            is_block = (is_start and is_end) or val.startswith("$$")
            
            # Clean up preceding text space on the same line if it's block math
            if is_block and i > 0 and tokens[i-1][0] == "text" and len(result) > 0:
                result[-1] = re.sub(r'[ \t]*$', '', result[-1])
                
            # Request stripping of leading spaces/tabs on the next line
            if is_block and i < len(tokens) - 1 and tokens[i+1][0] == "text":
                strip_leading_spaces = True
                
            # Extract raw formula
            pure_formula = val
            if token_type == "sympy":
                if val.startswith("sympy:{"):
                    pure_formula = val[7:-1].strip()
                elif val.startswith("py:{"):
                    pure_formula = val[4:-1].strip()
            elif token_type == "typst":
                if val.startswith("typst:{"):
                    pure_formula = val[7:-1].strip()
                elif val.startswith("t:{"):
                    pure_formula = val[3:-1].strip()
            
            try:
                if token_type == "sympy":
                    latex_eq = convert_sympy_to_latex(pure_formula)
                elif token_type == "mathml":
                    latex_eq = mathml_to_latex(pure_formula)
                else:  # typst
                    latex_eq = typst_to_latex(pure_formula)
                    
                if is_block:
                    result.append(f"$$\n{latex_eq}\n$$")
                else:
                    result.append(f"${latex_eq}$")
            except Exception as e:
                snippet = pure_formula[:30] + "..." if len(pure_formula) > 30 else pure_formula
                print(f"Warning: Failed to convert {token_type.upper()} formula '{snippet}' at line {current_line}: {e}", file=sys.stderr)
                result.append(val)
        else:
            result.append(val)
            
        current_line += token_lines
            
    return "".join(result)


def run_pandoc_file(input_file: str, output_file: str, from_fmt: str, to_fmt: str, extra_args: list = [], cwd: str = None):
    """Run pandoc command to convert files."""
    cmd = [get_pandoc_command(), "-f", from_fmt, "-t", to_fmt] + extra_args + [input_file, "-o", output_file]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", cwd=cwd)
    if res.returncode != 0:
        raise RuntimeError(f"Pandoc file conversion failed: {res.stderr}")


def check_pandoc_installed():
    """Verify that Pandoc is installed and accessible."""
    cmd = get_pandoc_command()
    try:
        subprocess.run([cmd, "--version"], capture_output=True)
        return True
    except FileNotFoundError:
        return False


def print_syntax_examples():
    """Print screen-reader-friendly equation typesetting examples in Typst, LaTeX, and SymPy."""
    print("=" * 60)
    print(" EQUATION CONVERTER MATH SYNTAX CHEATSHEET")
    print("=" * 60)
    print("\nYou can write math equations inside your Markdown document using:")
    print("  1. Typst Math   : t:{...} or typst:{...} (Recommended for simplicity)")
    print("  2. LaTeX Math   : $...$ (inline) or $$...$$ (block)")
    print("  3. SymPy Math   : py:{...} or sympy:{...} (Python expressions)")
    print("\n" + "-" * 50)
    print(" ALGEBRA EXAMPLES")
    print("-" * 50)
    print("  * Quadratic Formula:")
    print("    - Typst: t:{x = (-b +- sqrt(b^2 - 4a c)) / (2a)}")
    print("    - LaTeX: $$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$")
    print("    - SymPy: py:{x = (-b + sqrt(b**2 - 4*a*c))/(2*a)}")
    print("\n  * Fractions & Powers:")
    print("    - Typst: t:{1 / (x^2 + 1)}")
    print("    - LaTeX: $1 / (x^2 + 1)$ or $\\frac{1}{x^2 + 1}$")
    print("    - SymPy: py:{1 / (x**2 + 1)}")
    print("\n" + "-" * 50)
    print(" CALCULUS EXAMPLES")
    print("-" * 50)
    print("  * Derivative (Limit Definition):")
    print("    - Typst: t:{f'(x) = lim_(h -> 0) (f(x+h) - f(x)) / h}")
    print("    - LaTeX: $$f'(x) = \\lim_{h \\to 0} \\frac{f(x+h) - f(x)}{h}$$")
    print("\n  * Definite Integral:")
    print("    - Typst: t:{integral_a^b f(x) dif x}")
    print("    - LaTeX: $$\\int_{a}^{b} f(x) \\, dx$$")
    print("\n" + "-" * 50)
    print(" LINEAR ALGEBRA & MATRICES")
    print("-" * 50)
    print("  * Column Vector:")
    print("    - Typst: t:{vec(x; y)}")
    print("    - LaTeX: $$\\begin{pmatrix} x \\\\ y \\end{pmatrix}$$")
    print("\n  * 2x2 Matrix:")
    print("    - Typst: t:{mat(a, b; c, d)}")
    print("    - LaTeX: $$\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$$")
    print("=" * 60)


def latex_to_mathml(latex_str: str) -> str:
    """Convert LaTeX math to MathML markup using Pandoc."""
    try:
        if not (latex_str.startswith("$") or latex_str.startswith("$$")):
            latex_str = f"${latex_str}$"
        html_out = run_pandoc_conversion(latex_str, "markdown", "html", ["--mathml"])
        match = re.search(r"<math.*?>.*?</math>", html_out, re.DOTALL)
        if match:
            return match.group(0).strip()
        return html_out.strip()
    except Exception as e:
        raise RuntimeError(f"Failed to convert LaTeX to MathML: {e}")


def run_repl():
    """Run an interactive REPL loop to translate formulas in real-time."""
    print("=" * 60)
    print(" EQUATION CONVERTER INTERACTIVE REPL")
    print("=" * 60)
    print("Translate formulas on-the-fly. Supported inputs:")
    print("  * Default           : Typst syntax (e.g. sum_(k=0)^n)")
    print("  * Prefix with 'py:' : SymPy syntax (e.g. py:x**2 + 2*x)")
    print("  * MathML markup     : Starts with '<math' (e.g. <math><mi>x</mi></math>)")
    print("\nType 'exit' or 'quit' to close the session.")
    print("=" * 60)
    
    while True:
        try:
            line = input("\nequation> ").strip()
            if not line:
                continue
            if line.lower() in ("exit", "quit"):
                print("Exiting REPL. Goodbye!")
                break
                
            if line.startswith("<math"):
                try:
                    latex = mathml_to_latex(line)
                    print(f"  LaTeX : ${latex}$")
                except Exception as e:
                    print(f"  Error converting MathML: {e}")
            elif line.startswith("py:") or line.startswith("sympy:"):
                if line.startswith("py:"):
                    formula = line[3:].strip()
                else:
                    formula = line[6:].strip()
                if formula.startswith("{") and formula.endswith("}"):
                    formula = formula[1:-1].strip()
                try:
                    latex = convert_sympy_to_latex(formula)
                    mathml = latex_to_mathml(latex)
                    print(f"  LaTeX  : ${latex}$")
                    print(f"  MathML : {mathml}")
                except Exception as e:
                    print(f"  Error converting SymPy: {e}")
            else:
                formula = line
                if formula.startswith("t:{") and formula.endswith("}"):
                    formula = formula[3:-1].strip()
                elif formula.startswith("typst:{") and formula.endswith("}"):
                    formula = formula[7:-1].strip()
                try:
                    latex = typst_to_latex(formula)
                    print(f"  LaTeX  : ${latex}$")
                    try:
                        mathml = latex_to_mathml(latex)
                        print(f"  MathML : {mathml}")
                    except Exception:
                        pass
                except Exception as e:
                    print(f"  Error converting Typst: {e}")
        except (KeyboardInterrupt, EOFError):
            print("\nExiting REPL. Goodbye!")
            break


def main():
    if len(sys.argv) < 3:
        # Check if the user is asking for examples, repl, or help
        if len(sys.argv) == 2:
            arg = sys.argv[1].lower().strip()
            if arg in ("examples", "--examples", "-e"):
                print_syntax_examples()
                sys.exit(0)
            elif arg in ("repl", "interactive"):
                run_repl()
                sys.exit(0)
            elif arg in ("help", "--help", "-h"):
                # Fall through to standard help message below
                pass
            else:
                print(f"Error: Unknown option '{sys.argv[1]}'.", file=sys.stderr)
            
        print("Usage: python convert_documents.py <input_file> <word|html|pdf|md> [output_file]")
        print("Or view syntax examples: python convert_documents.py examples")
        print("Or run interactive translator: python convert_documents.py repl")
        print("\nExamples:")
        print("  python convert_documents.py input.md word")
        print("  python convert_documents.py input.html word")
        print("  python convert_documents.py input.docx html")
        print("  python convert_documents.py input.md pdf")
        print("  python convert_documents.py input.docx md")
        sys.exit(1)
        
    input_file = os.path.abspath(sys.argv[1])
    target_format = sys.argv[2].lower().strip()
    
    if target_format not in ("word", "html", "pdf", "md", "markdown"):  # epub and daisy disabled in this version
        print(f"Error: Unknown output format '{target_format}'. Must be 'word', 'html', 'pdf', or 'md'.", file=sys.stderr)
        sys.exit(1)
        
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.", file=sys.stderr)
        sys.exit(1)
        
    if not check_pandoc_installed():
        print("Error: Pandoc is not installed locally or globally.", file=sys.stderr)
        print("Please place 'pandoc.exe' in this script's directory (or 'bin/' folder),", file=sys.stderr)
        print("or install Pandoc system-wide (https://pandoc.org/).", file=sys.stderr)
        sys.exit(1)
        
    if target_format == "pdf" and not check_typst_installed():
        print("Error: Typst is not installed locally or globally.", file=sys.stderr)
        print("Please place 'typst.exe' in this script's directory (or 'bin/' folder),", file=sys.stderr)
        print("or install Typst system-wide (https://typst.app/).", file=sys.stderr)
        sys.exit(1)
        
    # Determine output file path
    if len(sys.argv) >= 4:
        output_file = sys.argv[3]
    else:
        base_name, _ = os.path.splitext(input_file)
        # Append _out to the default output filename to prevent overwriting and distinguish output files
        if target_format == "word":
            output_file = base_name + "_out.docx"
        elif target_format == "pdf":
            output_file = base_name + "_out.pdf"
        # elif target_format == "epub":
        #     output_file = base_name + "_out.epub"
        # elif target_format == "daisy":
        #     output_file = base_name + "_daisy.epub"
        elif target_format in ("md", "markdown"):
            output_file = base_name + "_out.md"
        else:
            output_file = base_name + "_out.html"
            
    try:
        input_dir = os.path.dirname(os.path.abspath(input_file))
        # Determine the input format based on file extension
        _, input_ext = os.path.splitext(input_file)
        input_ext = input_ext.lower().strip()
        
        if input_ext in (".md", ".markdown", ".txt"):
            from_fmt = "markdown"
        elif input_ext == ".docx":
            from_fmt = "docx"
        elif input_ext in (".html", ".htm"):
            from_fmt = "html"
        else:
            from_fmt = input_ext.lstrip(".") if input_ext else "markdown"

        # 1. Read input file (convert to intermediate markdown if not already markdown)
        if from_fmt != "markdown":
            # Convert non-markdown input to intermediate Markdown via Pandoc
            with tempfile.NamedTemporaryFile("w", dir=input_dir, suffix=".md", delete=False, encoding="utf-8") as temp_in_f:
                temp_in_file_path = temp_in_f.name
            try:
                print(f"Converting input {from_fmt.upper()} to intermediate Markdown...")
                extra_args = []
                if from_fmt == "docx":
                    extra_args = ["--extract-media", "."]
                run_pandoc_file(input_file, temp_in_file_path, from_fmt=from_fmt, to_fmt="markdown-grid_tables-multiline_tables-simple_tables", extra_args=extra_args, cwd=input_dir)
                content = read_file_with_fallback(temp_in_file_path)
                content = unescape_escaped_math(content)
            finally:
                if os.path.exists(temp_in_file_path):
                    os.remove(temp_in_file_path)
        else:
            # Read input markdown file directly
            content = read_file_with_fallback(input_file)
            
        # Parse YAML frontmatter (metadata) and extract remaining content
        metadata, content = parse_yaml_frontmatter(content)
        
        # Check and set metadata properties
        lang = metadata.get("lang", "en-US")
        title = metadata.get("title", "")
        author = metadata.get("author", "")
        
        if not title:
            # Fallback title to filename base
            title, _ = os.path.splitext(os.path.basename(input_file))
            title = title.replace("_", " ").replace("-", " ").title()
            print(f"Accessibility Info: No title found in frontmatter. Using fallback title: '{title}'")
            
        if "lang" not in metadata:
            print(f"Accessibility Info: No document language specified. Defaulting to 'en-US'")
            
        # Clean brand names and other old references
        lines = content.splitlines()
        filtered_lines = []
        for line in lines:
            if any(name in line.lower() for name in ("quill math test document", "edsharp math test document", "portable math creator math test document")):
                continue
            # Replace old branding references to keep the document aligned with Portable Math Creator (PMC)
            cleaned = re.sub(r"\b(?:Quill|EdSharp)'s\b", "the", line, flags=re.IGNORECASE)
            cleaned = re.sub(r"\b(?:Quill|EdSharp)\b", "Portable Math Creator (PMC)", cleaned, flags=re.IGNORECASE)
            filtered_lines.append(cleaned)
        # Normalize tables to ensure equal column count per row (Stage 2/3 requirement)
        content = normalize_markdown_tables("\n".join(filtered_lines))
        
        # Validate markdown accessibility (WCAG 2.2 audit)
        validate_markdown_accessibility(content)
        
        # Remediation: Fix missing alt texts using custom string format
        content = fix_missing_alt_texts(content)
            
        # 2. Process SymPy and MathML equations, converting them to LaTeX
        print("Parsing equations...")
        converted_content = process_markdown_equations(content, metadata)
        
        # 3. Write intermediate converted markdown to a temp file
        with tempfile.NamedTemporaryFile("w", dir=input_dir, suffix=".md", delete=False, encoding="utf-8") as temp_f:
            temp_f.write(converted_content)
            temp_file_path = temp_f.name
            
        try:
            # 4. Invoke Pandoc to build the output file
            pandoc_bin = get_pandoc_command()
            print(f"Using Pandoc: {os.path.basename(pandoc_bin)} ({os.path.abspath(pandoc_bin)})")
            
            if target_format == "word":
                print(f"Converting to MS Word (OMML equations) -> {output_file}...")
                docx_args = ["-M", f"lang={lang}", "-M", f"title={title}"]
                if author:
                    docx_args += ["-M", f"author={author}"]
                run_pandoc_file(temp_file_path, output_file, from_fmt="markdown", to_fmt="docx", extra_args=docx_args)
                
                # Apply Stage 3 accessibility enhancements!
                print("Post-processing DOCX for table accessibility and repeating headers...")
                post_process_docx_accessibility(output_file, lang)
            elif target_format == "pdf":
                # Try to use MS Word COM automation first for highly accessible mathematical tagged PDFs
                print("Attempting to compile to PDF via MS Word COM for accessibility...")
                with tempfile.NamedTemporaryFile("w", dir=input_dir, suffix=".docx", delete=False) as temp_docx_f:
                    temp_docx_path = temp_docx_f.name
                
                word_success = False
                try:
                    docx_args = ["-M", f"lang={lang}", "-M", f"title={title}"]
                    if author:
                        docx_args += ["-M", f"author={author}"]
                    run_pandoc_file(temp_file_path, temp_docx_path, from_fmt="markdown", to_fmt="docx", extra_args=docx_args)
                    post_process_docx_accessibility(temp_docx_path, lang)
                    
                    word_success = convert_docx_to_pdf_via_word(temp_docx_path, output_file)
                except Exception as e:
                    print(f"Word PDF conversion failed or not supported: {e}", file=sys.stderr)
                    word_success = False
                finally:
                    if os.path.exists(temp_docx_path):
                        try:
                            os.remove(temp_docx_path)
                        except:
                            pass
                
                if word_success:
                    print("PDF conversion completed successfully via MS Word (OMML equations)!")
                else:
                    print("MS Word COM not available. Falling back to Typst compiling engine...")
                    typst_bin = get_typst_command()
                    print(f"Using Typst: {os.path.basename(typst_bin)} ({os.path.abspath(typst_bin)})")
                    
                    # Create a temporary typst markup file
                    with tempfile.NamedTemporaryFile("w", dir=input_dir, suffix=".typ", delete=False, encoding="utf-8") as temp_typ_f:
                        temp_typ_path = temp_typ_f.name
                    
                    try:
                        print("Generating intermediate Typst markup...")
                        # We pass --standalone and metadata to generate full document structure with pandoc helper definitions
                        typst_args = ["--standalone", "-M", f"title={title}", "-M", f"lang={lang}"]
                        if author:
                            typst_args += ["-M", f"author={author}"]
                        run_pandoc_file(temp_file_path, temp_typ_path, from_fmt="markdown", to_fmt="typst", extra_args=typst_args)
                        
                        # Prepend a professional layout and accessibility template
                        typst_content = read_file_with_fallback(temp_typ_path)
                        
                        # We look for "#show: doc => conf(" and inject our parameters
                        target_str = "#show: doc => conf("
                        replacement_str = (
                            "#show: doc => conf(\n"
                            '  paper: "us-letter",\n'
                            '  margin: (x: 2.5cm, y: 2.5cm),\n'
                            '  fontsize: 11pt,\n'
                            '  font: "Segoe UI",\n'
                            '  mathfont: "Cambria Math",\n'
                            '  linestretch: 1.35,\n'
                        )
                        
                        if target_str in typst_content:
                            typst_content = typst_content.replace(target_str, replacement_str)
                        else:
                            # Fallback: Prepend standard set rules at the top if standalone structure is missing
                            layout_template = (
                                '#set document(\n'
                                f'  title: "{title}",\n'
                                f'  author: "{author}",\n'
                                f'  language: "{lang[:2]}",\n'
                                ')\n'
                                '#set page(\n'
                                '  paper: "us-letter",\n'
                                '  margin: (x: 2.5cm, y: 2.5cm),\n'
                                '  numbering: "1",\n'
                                ')\n'
                                '#set text(size: 11pt, font: "Segoe UI")\n'
                                '#show math.equation: set text(font: "Cambria Math")\n'
                                '#set par(justify: true, leading: 0.65em)\n\n'
                            )
                            typst_content = layout_template + typst_content
                        
                        # Let's replace the default basic horizontal rule with a premium padded one
                        typst_content = typst_content.replace(
                            "#let horizontalrule = line(start: (25%,0%), end: (75%,0%))",
                            '#let horizontalrule = pad(y: 0.8em, line(start: (0%,0%), end: (100%,0%), stroke: 0.5pt + rgb("#cbd5e1")))'
                        )
                        
                        # Prepend custom styles to the document
                        custom_styles = (
                            "// Custom premium styling overrides\n"
                            "#let frac = math.frac\n"
                            "#show heading: set text(fill: rgb(\"#1e293b\"))\n"
                            "#show heading.where(level: 1): it => block(below: 1em, above: 1.8em, it)\n"
                            "#show heading.where(level: 2): it => block(below: 0.8em, above: 1.4em, it)\n"
                            "#show heading.where(level: 3): it => block(below: 0.6em, above: 1.2em, it)\n"
                            "#show par: set block(spacing: 1.2em)\n\n"
                        )
                        # Clean up unsupported newer Typst attributes to match local compiler capabilities
                        typst_content = typst_content.replace('scope: "parent",', '')
                        
                        typst_content = custom_styles + typst_content
                        
                        with open(temp_typ_path, "w", encoding="utf-8") as f:
                            f.write(typst_content)
                            
                        print(f"Compiling to PDF (Typst engine) -> {output_file}...")
                        cmd = [typst_bin, "compile", temp_typ_path, output_file]
                        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
                        if res.returncode != 0:
                            raise RuntimeError(f"Typst compilation failed: {res.stderr}")
                    finally:
                        if os.path.exists(temp_typ_path):
                            os.remove(temp_typ_path)
            elif target_format in ("md", "markdown"):
                print(f"Saving converted Markdown -> {output_file}...")
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(converted_content)
            # epub and daisy targets disabled in this version
            # elif target_format == "epub":
            #     print(f"Converting to EPUB (MathML equations) -> {output_file}...")
            #     epub_args = ["--mathml", "-s", "-M", f"lang={lang}", "-M", f"title={title}"]
            #     if author:
            #         epub_args += ["-M", f"author={author}"]
            #     run_pandoc_file(temp_file_path, output_file, from_fmt="markdown", to_fmt="epub3", extra_args=epub_args)
            # elif target_format == "daisy":
            #     print(f"Converting to DAISY-compliant EPUB 3 (MathML equations) -> {output_file}...")
            #     daisy_args = [
            #         "--mathml", 
            #         "-s", 
            #         "-M", f"lang={lang}", 
            #         "-M", f"title={title}",
            #         "-M", "accessibilityFeature=structuralNavigation",
            #         "-M", "accessibilityFeature=MathML",
            #         "-M", "accessibilityHazard=none"
            #     ]
            #     if author:
            #         daisy_args += ["-M", f"author={author}"]
            #     run_pandoc_file(temp_file_path, output_file, from_fmt="markdown", to_fmt="epub3", extra_args=daisy_args)
            else:
                print(f"Converting to HTML (MathML equations) -> {output_file}...")
                html_args = ["--mathml", "-s", "-M", f"lang={lang}", "-M", f"title={title}"]
                if author:
                    html_args += ["-M", f"author={author}"]
                run_pandoc_file(temp_file_path, output_file, from_fmt="markdown", to_fmt="html", extra_args=html_args)
                
                # Apply Stage 2 accessibility enhancements!
                print("Post-processing HTML for WCAG 2.2 accessibility compliance...")
                post_process_html_accessibility(output_file, title)
                
                # Run automated checks
                run_automated_accessibility_checks(output_file)
                
            print("Conversion completed successfully!")
        finally:
            # Clean up temp file
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                
    except Exception as e:
        print(f"Error during conversion: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()