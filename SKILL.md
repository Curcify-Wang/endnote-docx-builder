---
name: endnote-docx-builder
description: 	 Trigger this skill whenever a manuscript .docx is being produced for a user who manages references in EndNote. Triggers include any mention of "EndNote", "Cite While You Write" / "CWYW", "Update Citations and Bibliography", "Format Bibliography", "更新引用", "更新參考文獻", "可更新引用", "EndNote 整合", "Vancouver 編號 + Word 輸出", or any time the user asks for a .docx with citations that they expect EndNote to manage. ALWAYS use this skill in combination with academic-paper-writing when the deliverable is a Word manuscript; do NOT fall back to plain numbered [N] citations or to EndNote temporary citations of the form {Author, Year #N} — both fail in EndNote X9+ for the reasons documented below. Skip this skill only when the user explicitly does not use EndNote (e.g., Zotero, Mendeley, Paperpile, BibTeX/LaTeX).
---

# EndNote-Compatible .docx Builder

This skill produces Microsoft Word manuscripts in which **every inline citation is a real EndNote field code** — the same Word-field structure that EndNote inserts when you click *Insert Citation* in the CWYW ribbon. The result is a .docx that EndNote treats as fully its own: *Update Citations and Bibliography*, *Format Bibliography*, style switching, renumbering, and bibliography insertion all work without prompting the user to "Select Matching Reference".

This skill exists because **all three lighter-weight approaches fail in EndNote X9 and later**, and they fail in ways that are easy to mistake for the format being correct. The history below is not background colour — it is the diagnostic tree future sessions must check against before assuming a Word output is EndNote-ready.

## What does not work, and why (do not repeat these mistakes)

| Approach | Inline form | Why it fails |
|---|---|---|
| 1. Plain numbered text | `[1]`, `[2, 3]` | EndNote does not scan for `[N]` patterns. The numbers are literal text. CWYW ignores them. |
| 2. EndNote temporary citation, no record number | `{Baker, 2013}` | Strictly invalid format in EndNote X9 and later. The `#RecordNumber` token is required. CWYW silently skips these. |
| 3. EndNote temporary citation with record number | `{Baker, 2013 #1}` | Syntactically valid, but only some EndNote configurations auto-scan these on Update; others require a manual "Convert Citations" step. Behaviour varies across versions, libraries, and CWYW settings. Not reliable enough to ship as a default. |
| **4. Real Word field code with `ADDIN EN.CITE`** | (Word field, displays as `[1]` in unformatted view) | **This is the format EndNote itself produces.** Recognised by every modern EndNote version unconditionally. The skill produces this. |

The user has reproduced failures 1–3 in EndNote 21. Approach 4 is the one that works.

## What a working citation looks like

Each citation is a Word complex field — five `<w:r>` runs in the underlying OOXML — wrapping an EndNote ADDIN instruction:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> ADDIN EN.CITE
  &lt;EndNote&gt;
    &lt;Cite&gt;
      &lt;Author&gt;Baker&lt;/Author&gt;
      &lt;Year&gt;2013&lt;/Year&gt;
      &lt;RecNum&gt;1&lt;/RecNum&gt;
      &lt;DisplayText&gt;[1]&lt;/DisplayText&gt;
      &lt;record&gt; ...full reference data (traveling library)... &lt;/record&gt;
    &lt;/Cite&gt;
  &lt;/EndNote&gt;
</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>[1]</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

Three things matter, in order of importance:

1. **The five-run Word field structure** — `begin`, `instrText`, `separate`, display-text `<w:t>`, `end`. Without all five, Word cannot store the field at all.
2. **The `ADDIN EN.CITE` literal** at the start of the field instruction — this is the magic string EndNote CWYW scans for.
3. **The `<record>` block** — the *traveling library*. Including the full reference data inline (authors, title, journal, year, volume, issue, pages, publisher, URL) makes the document self-formatting: EndNote can format citations and build the bibliography even before the user imports the references into their library. Strongly recommended; nearly doubles the docx size but eliminates an entire failure mode.

## The trap that bit us (must keep checked)

**EndNote CWYW scans *all visible text* in the document for temporary citation patterns**, not just text in citation positions. If any paragraph anywhere — a placeholder, an instructional note, a figure caption, a comment — contains a literal `{Something, Something}` substring, EndNote will treat it as an unresolved temporary citation and pop up *"Select Matching Reference"* asking the user to pick a match by hand.

Before saving the .docx, **the build script must scan its own visible text output for any `{...}` patterns and fail loudly if found**. The reference template (`build_docx_template.py`) ships with this check built in. Do not bypass it.

This includes:

- README-style instructional paragraphs in the manuscript
- "EndNote will insert the bibliography here" placeholder text
- Figure 1 / Table 1 captions
- Author note sections, declarations, abbreviation lists

If you genuinely need to write the *string* `{Author, Year}` in a manuscript (e.g., explaining the format in a methodology paper about citation tools), use a non-curly alternative — `[Author, Year]` or `(Author, Year)` — or render as an image.

## Standard workflow when this skill fires

1. **Confirm the user uses EndNote.** If they use Zotero, Mendeley, BibTeX, or anything else, stop and use a different output path. EndNote field codes will not be recognised by other reference managers.

2. **Read the template** at `references/build_docx_template.py` in this skill folder. Copy it into the user's working directory (typically `<project>/build_docx.py` next to the .docx output folder). Adapt — do not rewrite from scratch.

3. **Populate the `REFERENCES` list** with one entry per cited work. Each entry has:
   - `temp`: `"Author, Year"` — first-author surname + year, used as the inline match key for EndNote
   - `ris`: a complete RIS block for the reference (multi-line string)
   
   The `temp` string must match the first author's surname and year in the RIS. EndNote matches by these on first format.

4. **Write manuscript text** with `{cite(n)}` or `{cite(n1, n2, n3)}` interpolations where citations belong. The template's `cite()` function returns an internal placeholder like `__CITE_1_14_16__`; the `populate_paragraph()` function later replaces these placeholders with real Word fields. This indirection is what keeps user-facing text free of literal `{...}` patterns.

5. **Use `populate_paragraph()` for every text-bearing paragraph and every table cell** — including abstract sections, captions, and table cells. Direct `paragraph.add_run(text)` will leave any `{cite(...)}` placeholders as raw text.

6. **Verify before declaring done.** After saving, the template runs an automatic post-check that:
   - Counts `ADDIN EN.CITE` Word fields (should equal expected citation count)
   - Scans visible text for any leftover `{...}` patterns (must be zero)
   - Reports both numbers in the build output

7. **Ship two files alongside the .docx**:
   - `references_<date>.ris` — a one-shot RIS bundle the user can import to their EndNote library. Although the traveling library makes this technically optional, importing once gives EndNote real record numbers in the user's library and unlocks "right-click → Edit Library Reference" workflows.
   - `ENDNOTE_README.md` — short workflow note: open .docx → set EndNote style → Update Citations and Bibliography. Document any references where Author+Year are not unique in the bundle (EndNote will prompt to disambiguate on first format).

## Versioning convention

The user has standardised on `Text<YYYY-MM-DD>_<N>.docx`, with `N` auto-incrementing on every build. Older versions are preserved. The template ships with the `next_version_path()` function that implements this. **Do not overwrite previous versions** — the user uses them for diff comparison and rollback.

## When the user reports "EndNote can't update"

Run this diagnostic tree:

1. **Open the .docx with `unzip -p file.docx word/document.xml`** and count `ADDIN EN.CITE` occurrences. If zero, citations were written as plain text — go back to step 5 above.
2. **Grep visible text for `{...}` patterns** (the template's post-check does this; if it was skipped, do it manually). Any hit triggers "Select Matching Reference" prompts.
3. **Confirm a bibliography style is selected in EndNote** — if it is "Annotated" or none, Update Citations does nothing visible.
4. **Confirm EndNote CWYW add-in is active in Word** — EndNote tab visible in the ribbon.
5. **Check for two references with identical Author+Year** — EndNote will prompt to disambiguate on first format. This is expected, not a bug. Document such pairs in the README.
6. **Ask the user what specifically happens** — "no response", "error dialog", "prompts for every citation", "prompts for one citation" — each points to a different cause.

## Reference example in the wild

A complete worked example lives in `/Users/curcifywang/Project/Markerless_Mocap_Review/`. The build script there is the same template shipped here, populated with 38 references for a markerless motion capture narrative review. Use it as a reading example; do not edit it.

## Files in this skill

- `SKILL.md` — this file
- `build_docx_template.py` — the working generator template (copy, adapt, run)
- `references/endnote_field_format.md` — deeper reference on the EndNote OOXML field format, RIS to record-XML mapping, and per-version EndNote quirks

## Cross-reference

The parent skill **academic-paper-writing** delegates all .docx output to this skill. When invoking academic-paper-writing for an EndNote user, both skills are active: academic-paper-writing handles IMRaD structure / Vancouver vs APA / user-specific style; this skill handles the mechanical Word-field output.
