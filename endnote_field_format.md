# EndNote Field Code — Technical Reference

Deeper notes on the OOXML field structure, EndNote XML schema, and per-version quirks. The `SKILL.md` covers the workflow; this file is for when something does not work and you need to inspect the bytes.

## Anatomy of a working citation field

In OOXML (`word/document.xml`), a single citation occupies five consecutive `<w:r>` runs inside one `<w:p>` paragraph:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r>
  <w:instrText xml:space="preserve"> ADDIN EN.CITE
    &lt;EndNote&gt;
      &lt;Cite&gt;
        &lt;Author&gt;Baker&lt;/Author&gt;
        &lt;Year&gt;2013&lt;/Year&gt;
        &lt;RecNum&gt;1&lt;/RecNum&gt;
        &lt;DisplayText&gt;[1]&lt;/DisplayText&gt;
        &lt;record&gt; ... &lt;/record&gt;
      &lt;/Cite&gt;
    &lt;/EndNote&gt;
  </w:instrText>
</w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" .../></w:rPr><w:t>[1]</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

Why each piece is mandatory:

- **`begin` / `end` field chars** — Without them Word cannot store the construct as a field at all; the EndNote XML would survive as raw text.
- **`xml:space="preserve"` on `instrText`** — Without this, leading/trailing whitespace inside the instruction can be collapsed by some Word versions, which corrupts the ` ADDIN EN.CITE ` prefix.
- **`separate` field char** — Marks the boundary between the field instruction and the field result. The display text (`<w:t>[1]</w:t>`) lives between `separate` and `end`. Without `separate`, Word treats the entire field as instruction-only and shows nothing.
- **Display text runs** — What the reader sees before EndNote ever formats anything. Set the font here too, otherwise Word may inherit Calibri or whatever the doc default is.

The five `<w:r>` runs must all be inside the **same** `<w:p>` paragraph. A paragraph break between `begin` and `end` invalidates the field.

## Multiple citations in one field

Two references cited together (e.g., `[1, 2]`) share one Word field with two `<Cite>` children:

```xml
ADDIN EN.CITE
  <EndNote>
    <Cite>
      <Author>Baker</Author><Year>2013</Year><RecNum>1</RecNum>
      <DisplayText>[1, 2]</DisplayText>
      <record>...</record>
    </Cite>
    <Cite>
      <Author>Windolf</Author><Year>2008</Year><RecNum>2</RecNum>
      <record>...</record>
    </Cite>
  </EndNote>
```

`<DisplayText>` appears only on the first `<Cite>` — it applies to the whole group.

Do NOT emit two separate Word fields side by side for `[1, 2]`. EndNote will format them as `[1][2]` instead of `[1, 2]` because each field is an independent citation point in its model.

## The `<record>` block (traveling library)

The minimum that lets EndNote *match* a citation is `<Author>`, `<Year>`, `<RecNum>`. The minimum that lets EndNote *format* the bibliography without the user having imported the references is the full `<record>` block:

```xml
<record>
  <rec-number>1</rec-number>
  <foreign-keys>
    <key app="EN" db-id="local" timestamp="0">1</key>
  </foreign-keys>
  <ref-type name="Journal Article">17</ref-type>
  <contributors>
    <authors>
      <author>Surname, F. I.</author>
      <author>Surname2, G. J.</author>
    </authors>
  </contributors>
  <titles>
    <title>Sentence-case title here</title>
    <secondary-title>J Abbrev Name</secondary-title>
  </titles>
  <periodical><full-title>J Abbrev Name</full-title></periodical>
  <pages>123-130</pages>
  <volume>43</volume>
  <number>1</number>
  <dates><year>2021</year></dates>
  <electronic-resource-num>10.1234/example.2021</electronic-resource-num>
  <urls><related-urls><url>https://...</url></related-urls></urls>
</record>
```

Recommended always: include `<record>`. The size cost (~300 B per citation) is trivial; the failure mode it eliminates (user has not yet imported references) is common.

## Reference type codes (EndNote ref-type)

EndNote uses both a name attribute and a numeric ID. Both should be present.

| RIS TY | Name | EndNote ID |
|---|---|---|
| JOUR | Journal Article | 17 |
| BOOK | Book | 6 |
| CHAP | Book Section | 5 |
| CONF | Conference Proceedings | 10 |
| COMP | Computer Program | 9 |
| RPRT | Report | 27 |
| THES | Thesis | 32 |
| ELEC | Electronic Article | 43 |
| GEN  | Generic | 13 |

If the bundle contains other types (datasets, software, preprints, web pages, government documents), look them up in the EndNote *Reference Types* preference pane before guessing — wrong ref-type causes wrong style template selection and wrong bibliography format.

## RIS → record-XML field mapping

| RIS tag | Record XML element |
|---|---|
| AU | `<authors>/<author>` (one element per AU) |
| TI | `<titles>/<title>` |
| T2 | `<titles>/<secondary-title>` and `<periodical>/<full-title>` |
| PY | `<dates>/<year>` |
| VL | `<volume>` |
| IS | `<number>` |
| SP–EP | `<pages>` (`SP-EP` joined with hyphen) |
| SP only | `<pages>` (start page alone) |
| PB | `<publisher>` |
| CY | `<pub-location>` |
| DO | `<electronic-resource-num>` |
| UR | `<urls>/<related-urls>/<url>` |

The template's `make_record_xml()` implements this map. Add more mappings only if RIS bundles in the project actually carry them.

## XML escaping rules

Inside `<w:instrText>`, the EndNote XML is stored as **plain text inside an XML element** — angle brackets must be escaped:

- `<` → `&lt;`
- `>` → `&gt;`
- `&` → `&amp;`
- `"` and `'` are usually OK unescaped inside instrText but `xml.sax.saxutils.escape()` will escape them too.

Inside the EndNote XML itself, the same rules apply to user-supplied text (author names, titles, journal names). Use `xml_escape()` from `xml.sax.saxutils`.

When Word reads the docx back, both layers un-escape, and the EndNote XML reaches CWYW as raw angle-bracketed XML — which is what EndNote expects.

## Per-version EndNote behaviour

| Version | Notes |
|---|---|
| **EndNote X8 and earlier** | `{Author, Year #N}` temporary citations work via "Convert Citations" command. ADDIN EN.CITE fields also work. |
| **EndNote X9** | Temporary citations begin to require strict syntax (`#N` mandatory). Behaviour around Update Citations is inconsistent. ADDIN EN.CITE fields work. |
| **EndNote 20** | Temporary citations work in some libraries but not others; depends on CWYW settings. ADDIN EN.CITE fields work. |
| **EndNote 21** | Temporary citations effectively deprecated for new documents — Update Citations will not auto-scan them unless the user clicks Convert Citations first. ADDIN EN.CITE fields work. **This is the version we have validated against.** |
| **EndNote 2024+** | Same as 21 — ADDIN EN.CITE field codes remain the supported integration. |

The user has EndNote 21 on macOS, validated 2026-05-26.

## CWYW pitfall — "Select Matching Reference" prompts on any visible `{...}`

EndNote CWYW's first pass after Update Citations is a regex sweep over **all visible text** for the pattern `\{[^}]*\}` (default temporary citation delimiters). Anything matched is treated as an unresolved temporary citation, and a *Select Matching Reference* dialog is raised for it.

This includes text in:

- Body paragraphs (obvious)
- Figure captions and table cells
- Headers and footers
- Comments and revision marks
- Footnotes and endnotes
- Text boxes and shape labels

**Whatever the build script writes as user-visible text must contain zero `{...}` substrings.** The template's post-check enforces this. Examples that have caused real failures:

- README-style placeholder text mentioning the `{Author, Year}` format
- Methodology paragraphs in a citation-tools paper explaining citation syntax
- Style examples in supplementary material

Workaround if you genuinely need to typeset `{...}` text in the manuscript: use Unicode look-alikes (`｛` U+FF5B and `｝` U+FF5D, full-width braces), or escape as `[Author, Year]` with square brackets, or render as an image.

## Inspecting a docx after the fact

Quick diagnostic commands:

```bash
# Count real EndNote citation fields
unzip -p mydoc.docx word/document.xml | grep -c 'ADDIN EN.CITE'

# Dump the first field instruction (verify XML structure)
unzip -p mydoc.docx word/document.xml | grep -oE 'ADDIN EN.CITE[^<]+' | head -1

# Scan visible text for stray {...} patterns
python3 -c "
import zipfile, re
xml = zipfile.ZipFile('mydoc.docx').read('word/document.xml').decode()
text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml))
hits = re.findall(r'\{[^}]*\}', text)
print(f'{len(hits)} suspect curly-brace fragments:')
for h in hits[:20]: print(' ', repr(h))
"
```

If the ADDIN count matches the expected citation count and the curly-brace scan is empty, the document is structurally ready for EndNote.
