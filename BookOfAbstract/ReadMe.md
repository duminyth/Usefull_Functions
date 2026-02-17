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
### Excel manuel processing
- The .py assume that each pdf can be found following an URL or a local path.
- The "Area" of each Abstract must match a predefined list. Either enforce this matching to the researcher registering, or modify it manually:
    - Variable `AREA_ORDER` in the .py file. Pay attention that `AREA_ORDER` is a list, the ordering of this list will be the ordering of the abstract in the .tex and .pdf files.
- The withdrawn abstract must have a `Withdrawn`status in Excel.
- The plenary presentation must have an "Area" being `Plenary`.
- The contributor naming is based on the Excel file on a *A. B. Name* system. Because some contributors are present in different works, it is important that the same naming room is followed when registering, especially regarding the (potential) second and third name. If one person register a coauthor only with the first name, and another register the same coauthor with first and last name, this coauthor will happen twice in the **List of Author**

---
### Features Latex

- Optional `Book-of-Abstracts_Front-part.pdf` inclusion at the beginning.
- Optional `Book-of-Book-of-Abstracts_final-page.pdf` inclusion at the end.
- Works best with **LuaLaTeX** (recommended)
- Use the font dinish, which has to be download from https://github.com/playbeing/dinish
- 

---

#### Features Python
- Python 3.9+ recommended
- The 

---
