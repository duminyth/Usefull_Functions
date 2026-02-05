# Create_BookOfAbstract.py

This python file generates a **Book of Abstracts** PDF from an Excel file that contains:
- URLs to PDF abstracts (one per row),
- a session/category (“area”) field used to order abstracts,
- title + authors information used to build a Table of Contents and an Author Index.

The output is a single **LaTeX `.tex`** file that:
- prepends an `Intro.pdf`,
- adds a grouped, clickable Table of Contents (TOC),
- embeds each abstract PDF (without extracting text),
- generates an **Author Index** with page hyperlinks to each abstract.

---

## Features

- Optional `Intro.pdf` inclusion at the beginning
- Works best with **LuaLaTeX** (recommended)

---

### Python
- Python 3.9+ recommended
- Packages:
  - `openpyxl`
  - `requests`

---
