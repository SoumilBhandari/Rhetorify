import re
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import webbrowser
import tempfile

def render_html(str_list):
    """Render a list of HTML fragments as a temporary HTML file in the browser."""
    html_lines = [
        '<!DOCTYPE html>',
        '<html>',
        '<head>',
        '  <meta charset="utf-8">',
        '  <style>',
        '    body { font-family: Calibri, sans-serif; font-size: 12pt; padding: 20px; }',
        '    .card { margin-bottom: 1em; }',
        '  </style>',
        '</head>',
        '<body>'
    ]
    for fragment in str_list:
        html_lines.append(f"<div class='card'>{fragment}</div>")
    html_lines.append('</body></html>')

    with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
        f.write("\n".join(html_lines))
        path = f.name
    webbrowser.open(f'file://{path}')

def render_string(text_str):
    """Render a single HTML string as a standalone card."""
    html = f"""
<!DOCTYPE html>
<html>
<head>
  <meta charset='utf-8'>
  <style>
    body {{ font-family: Arial, sans-serif; padding: 20px; }}
    .card {{ margin-bottom: 40px; border-bottom: 1px solid #ccc; padding-bottom: 20px; }}
  </style>
</head>
<body>
  <div class='card'>{text_str}</div>
</body>
</html>
"""
    with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
        f.write(html)
        path = f.name
    webbrowser.open(f'file://{path}')

def style(run):
    """Convert a run of text into HTML based on Word formatting."""
    text = run.text or ''
    if not text.strip():
        return ''
    if run.bold:
        text = f"<b>{text}</b>"
    if run.underline:
        text = f"<u>{text}</u>"
    if run.font.highlight_color and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
        text = f"<mark>{text}</mark>"
    if run.font.size:
        size_pt = run.font.size.pt
        text = f"<span style='font-size:{size_pt}pt'>{text}</span>"
    return text

def markdown(paragraph):
    """Convert a paragraph into HTML by concatenating styled runs."""
    return ''.join(style(r) for r in paragraph.runs)

def is_citation(paragraph):
    """Heuristic to detect citation paragraphs."""
    txt = paragraph.text.strip()
    if 'http://' in txt or 'https://' in txt or 'www.' in txt:
        return True
    score = 0
    if re.search(r"\b[A-Z][a-z]+,? \d{4}\b", txt):
        score += 1
    keywords = ['professor', 'journal', 'researcher', 'university', 'institute', 'PhD', 'published']
    if any(k in txt for k in keywords):
        score += 1
    return score >= 2

def is_tag(paragraph):
    """Heuristic to detect context tags preceding citations."""
    for run in paragraph.runs:
        if run.font.highlight_color is not None:
            return False
        if run.font.size and run.font.size.pt <= 10:
            return False
    return len(paragraph.text) < 500

def rhetorify(docx_path):
    """Main function: extract citation-context pairs from a .docx file."""
    doc = Document(docx_path)
    paras = [p for p in doc.paragraphs if p.text.strip()]
    output = []
    i = 0
    while i < len(paras):
        if is_citation(paras[i]):
            tag_html = ''
            if i > 0 and is_tag(paras[i-1]):
                tag_html = markdown(paras[i-1])
            cite_html = ''
            for run in paras[i].runs:
                if run.bold or (run.font.highlight_color and run.font.highlight_color != WD_COLOR_INDEX.AUTO):
                    cite_html += f"<b>{run.text}</b> "
            j = i + 1
            body_html = ''
            while j < len(paras) and not is_citation(paras[j]):
                for run in paras[j].runs:
                    if run.font.highlight_color and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
                        body_html += run.text + ' '
                j += 1
            combined = f"{tag_html} {cite_html}: {body_html}".strip()
            output.append(combined)
            i = j
        else:
            i += 1
    return output

if __name__ == '__main__':
    import sys
    if len(sys.argv) != 2:
        print("Usage: python rhetorify.py path/to/document.docx")
    else:
        frags = rhetorify(sys.argv[1])
        render_html(frags)
