# Rhetorify
Rhetorify is a lightweight Python utility to parse Microsoft Word (.docx) files, extract highlighted or bolded citation passages (author–year patterns, URLs, academic keywords) along with any preceding context tags, and render each snippet as a clean, browser-viewable HTML card.

Features
1. Heuristic detection of citations and context tags
2. Retains Word styling: bold, underline, highlights, font size

Two render modes:
1. render_html for batch export of HTML cards
2. render_string for single-card preview

Simple CLI interface: just point at your .docx file

# Installation
pip install python-docx

# Usage
python rhetorify.py path/to/document.docx
