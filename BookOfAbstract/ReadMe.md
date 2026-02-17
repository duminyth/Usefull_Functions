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

- Optional `Book-of-Abstracts_Front-part.pdf` inclusion at the beginning.
- Optional `Book-of-Book-of-Abstracts_final-page.pdf` inclusion at the end.
- Works best with **LuaLaTeX** (recommended)
- Use the font dinish, which has to be download from https://github.com/playbeing/dinish

---

### Python
- Python 3.9+ recommended
- Packages:
  - `openpyxl`
  - `requests`

---
