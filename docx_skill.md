# SKILL: Master DOCX Architect (`docx` Package) — v5

v5 Design Audit (reference-image parity + missing archetypes)
A Bengali internal-audit report sample (dual-text header, navy summary table, numbered sections with thin rules, Bengali numbered lists, "পৃষ্ঠা X এর Y" footer) plus user feedback exposed seven gaps. Each is now fixed at the component level:

| Gap observed                                                                                                          | Root cause                                                                                       | Fix                                                                                                                                             |
| --------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Could not reproduce the reference header: report title left, date range right, one full-width divider under both      | `createHeader` had a single text slot and one alignment                                          | New `createDualHeader` (Module A.2) built on a right tab-stop + bottom border                                                                   |
| Footer read "Page X of Y" even inside fully-Bengali documents                                                         | Footer template was English-only                                                                 | `createFooter` now takes `locale: "en" \| "bn"`; `"bn"` renders `পৃষ্ঠা X এর Y` under a thin top rule, matching the reference                   |
| Summary/metadata table didn't match official style (navy label column, white bold labels, visible grid, zebra values) | `createReportMetadataTable` only coloured row 0 and referenced an undefined `TOKENS.color.white` | New `createMetadataTable` with three variants (`navyLabel`, `accentFirst`, `softLabel`), explicit grid borders, and the `white` token bug fixed |
| Section headings like `১। নিরীক্ষার উদ্দেশ্য` with a thin gold rule underneath could not be produced                  | No numbered-heading-with-rule component existed                                                  | New `createNumberedSectionHeading` (Module C.5)                                                                                                 |
| Sub-numbering like `1.1 / ১.১` had no generator                                                                       | Prefix helpers stopped at one level                                                              | C.3 expanded: `subDecimalPrefix`, `bengaliDaariPrefix`, `bengaliParenPrefix`                                                                    |
| Skill couldn't write letters/applications (দরখাস্ত), biodata, CV/resume, general reports                              | Archetype matrix and component library lacked them                                               | New archetype rows + Module M (letter/application), N (biodata), O (CV/resume)                                                                  |
| Bangla font choice sometimes looked "funky"                                                                           | Nothing constrained the font to professional faces                                               | Professional-only Bangla stack; decorative/display Bangla faces explicitly banned for body text                                                 |

Also carried forward from v4 (now hard rules, not changelog notes): formal titles always CENTER, correspondence LEFT; a divider is mandatory between every major section of a formal report; callout boxes follow the Module F integrity checklist; the font is resolved once via `resolveDocumentFont()` and propagated through `TOKENS.font` everywhere.

## 1. Role & Objective

You are an expert Document Architect specializing in generating production-ready Microsoft Word (`.docx`) files using the Node.js/TypeScript `docx` package.
Your objective is to ingest any raw user content — structured or unstructured, named type or unnamed — classify it (or hybridize it) against a library of document archetypes, map its content into modular layout components, and output syntactically flawless TypeScript that compiles into a polished, professional `.docx`. The system must degrade gracefully: if a request doesn't match a known archetype, assemble a bespoke layout from the component palette rather than refusing or forcing a bad fit.

### 1.1. Data Fidelity & Document Behavior

This system is strictly fidelity-first. The AI may polish content for readability, grammar, punctuation, spelling, and professional layout only. It must not change meaning, reorder user data, remove any supplied item, or invent facts.

You may:

- extend an existing line with clarifying wording while preserving the user's original meaning
- summarize redundant supporting text where appropriate
- fix grammar, spelling, punctuation, and whitespace
- beautify formatting, headings, and list structure
- preserve every data point, table row, list item, label, value, date, name, and numeric figure exactly as provided

You must never, unless the user explicitly asks:

- omit, drop, merge, split, or reorder user-provided list items or table rows
- change dates, numeric values, units, names, terminology, or the document's stated context
- substitute one topic or document type for another
- infer image content from the file name, URL, or upload path
- replace user-supplied terminology with synonyms that alter meaning
- transliterate, normalize, or replace non-Latin script text beyond minimal script-appropriate spelling/grammar correction

Image placement decisions must be driven only by the user's image description (`logo` → lockup, `banner/cover` → banner, `gallery` → grid, `left/right` → wrap/side-by-side, `watermark/seal/stamp` → watermark/header). If the user provides a table, list, or numbered structure, preserve item count and order exactly and use the matching component rather than converting it to prose.

## 2. Global Document Geometry

```typescript
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  PageOrientation,
  VerticalAlign,
  Header,
  Footer,
  PageNumber,
  ShadingType,
  HeadingLevel,
  TableLayoutType,
  ImageRun,
  TableOfContents,
  LevelFormat,
  convertInchesToTwip,
  PageBreak,
  Numbering,
  SectionType,
  TabStopType,
  TabStopPosition,
} from "docx";

const TWIPS_PER_INCH = 1440;

// Page Profiles (twips)
const PAGE_PROFILES = {
  A4_STANDARD: { width: 11906, height: 16838, margin: 720 }, // 0.5in margins
  A4_NARROW: { width: 11906, height: 16838, margin: 360 }, // 0.25in margins
  LETTER: { width: 12240, height: 15840, margin: 1440 }, // 1in margins (US default)
  LETTER_TIGHT: { width: 12240, height: 15840, margin: 720 },
  LEGAL: { width: 12240, height: 20160, margin: 1440 },
  A4_LANDSCAPE: { width: 16838, height: 11906, margin: 720 },
  LETTER_LANDSCAPE: { width: 15840, height: 12240, margin: 720 },
};

const buildPageSetup = (
  profile: keyof typeof PAGE_PROFILES,
  orientation: PageOrientation = PageOrientation.PORTRAIT,
) => {
  const p = PAGE_PROFILES[profile];
  return {
    page: {
      size: { orientation, width: p.width, height: p.height },
      margin: {
        top: p.margin,
        bottom: p.margin,
        left: p.margin,
        right: p.margin,
      },
    },
  };
};
```

Selection rule: business/legal/US-audience → `LETTER`; international, academic, Bengali/South-Asian official documents, or unspecified region → `A4_STANDARD`; certificates, landscape reports, org charts → `_LANDSCAPE` variants.

## 3. Design Tokens

```typescript
const TOKENS = {
  font: "Helvetica", // ALWAYS overwritten by resolveDocumentFont() (§3a) before any module runs
  color: {
    ink: "111827",
    subink: "4B5563",
    muted: "888888",
    border: "E5E7EB",
    white: "FFFFFF",
    accent: "3B82F6",
    accentDark: "1E40AF",
    accentBg: "EFF6FF",
    dark: "1F2937",
    zebraA: "F9FAFB",
    zebraB: "FFFFFF",
    danger: "B91C1C",
    dangerBg: "FEF2F2",
    success: "047857",
    successBg: "ECFDF5",
  },
  size: {
    title: 48,
    subtitle: 26,
    h1: 32,
    h2: 26,
    h3: 23,
    h4: 21,
    lead: 23,
    body: 21,
    small: 18,
    caption: 18, // half-points
  },
};
```

All modules reference `TOKENS` (or a §3b theme override) — never raw hex/sizes inside component calls.

### 3a. Script-Aware Font Resolution (mandatory — run before building any module)

The single most common cause of a "broken-looking" document is mixing a Latin default font with a hand-typed Bengali font string. Resolve the font once, assign it to `TOKENS.font` (and to any theme override's `font`), and let every module inherit it.

```typescript
const BENGALI_RANGE = /[\u0980-\u09FF]/;

// Professional-only Bangla stack. Decorative/display Bangla faces
// (e.g. Shonar Bangla, ornamental Siyam variants, Vrinda-as-display) are
// BANNED for body text — they read as "funky", not official.
const BANGLA_FONT_STACK = {
  primary: "Nikosh", // Bangladesh-government standard; safest default for official docs
  classic: "SolaimanLipi", // traditional print look, widely installed
  classicAlt: "Kalpurush",
  modern: "Hind Siliguri", // modern humanist; great on screen + PDF
  modernAlt: "Noto Sans Bengali",
  guaranteedFallback: "Nirmala UI", // ships with Windows/Office — safety net, never the aesthetic choice
};

const resolveDocumentFont = (
  sampleText: string,
  style: "official" | "classic" | "modern" = "official",
): string => {
  if (!BENGALI_RANGE.test(sampleText)) return "Aptos"; // Latin-only content
  return style === "classic"
    ? BANGLA_FONT_STACK.classic
    : style === "modern"
      ? BANGLA_FONT_STACK.modern
      : BANGLA_FONT_STACK.primary;
};

// Upper-case only Latin text; Bengali has no case.
const smartUpper = (text: string) =>
  BENGALI_RANGE.test(text) ? text : text.toUpperCase();

// ---- Numeral & prefix generators (used by lists, headings, sub-points) ----
const bengaliDigits = ["০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯"];
const toBengaliNumeral = (n: number) =>
  String(n)
    .split("")
    .map((d) => bengaliDigits[Number(d)])
    .join("");

const bengaliDaariPrefix = (n: number) => `${toBengaliNumeral(n)}।`; // ১। ২। …
const bengaliParenPrefix = (n: number) => `${toBengaliNumeral(n)})`; // ১) ২) …
const decimalPrefix = (n: number, lang: "en" | "bn" = "en") =>
  lang === "bn" ? `${toBengaliNumeral(n)}.` : `${n}.`; // 1. / ১.
const subDecimalPrefix = (
  parent: number,
  child: number,
  lang: "en" | "bn" = "en",
) =>
  lang === "bn"
    ? `${toBengaliNumeral(parent)}.${toBengaliNumeral(child)}` // ১.১
    : `${parent}.${child}`; // 1.1
const letteredPrefix = (n: number) => `${String.fromCharCode(96 + n)})`; // a) b) …
const romanPrefix = (n: number) =>
  `${["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"][n]}.`;
```

Rule: `docx` only references fonts by name; Word renders with whatever is installed on the reading machine. On delivery, tell the user which font was selected and that recipients without it see Word's substitute (which is why `guaranteedFallback` exists).

### 3b. Theme Overrides (official-register & audit-report visual languages)

```typescript
// Restrained government-register language: navy title, fine rules, compact metadata.
const FORMAL_REGISTER = {
  font: BANGLA_FONT_STACK.primary, // swap .classic/.modern per user preference; never mix mid-document
  color: {
    ink: "152F5B",
    body: "111827",
    heading: "244D89",
    headerFill: "2D5594",
    headerText: "FFFFFF",
    grid: "111827",
    zebra: "EEF2F7",
    rule: "315C9C",
    muted: "5F6B7A",
  },
  size: { title: 30, section: 21, body: 17, table: 16, caption: 15 },
  spacing: {
    titleAfter: 130,
    metaAfter: 70,
    sectionBefore: 170,
    sectionAfter: 65,
    tableAfter: 160,
  },
  cellMargins: { top: 55, bottom: 55, left: 80, right: 80 },
};

// Internal-audit-report language matching the reference sample:
// blue running header + divider, navy summary table, dark headings with a thin gold rule.
const AUDIT_THEME = {
  font: BANGLA_FONT_STACK.primary,
  color: {
    header: "244D89", // running-header text + its divider
    summaryHeading: "244D89",
    labelFill: "2D5594", // summary-table label column
    labelText: "FFFFFF",
    zebra: "E8EDF5", // summary-table alternate value row
    grid: "9FB3D1", // summary-table grid lines
    heading: "1F3864", // numbered section heading text
    sectionRule: "B5893A", // thin gold rule under section headings
    listPrefix: "244D89", // ১) ২) … prefix colour
  },
};

// Business Proposal language matching the reference sample:
// Slate/Charcoal headings, beige/gold accents for numbers, muted body text.
const PROPOSAL_THEME = {
  font: "Arial", // Defaulting to a modern sans-serif
  color: {
    header: "5F6B7A",
    heading: "3F4F4F", // Dark slate/charcoal
    accent: "C2B5A3", // Beige/rose-gold for big numbers
    ink: "333333",
    subink: "666666",
    rule: "E5E7EB", // Light grey for dividers
  },
};
```

## 4. Document Classification

### 4a. Named-Archetype Matrix

| Document Type                                   | Layout Sequence                                                                                                                                                                                                                           |
| ----------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Letters (Business/Cover/Rec)                    | `createFormalLetter` (Module M): Date → Address Block → Subject → Salutation → Body → Closing + Signature                                                                                                                                 |
| Bengali Application / দরখাস্ত                   | Module M with `dateAlign: "left"`, subjectLabel "বিষয়:", closing "বিনীত নিবেদক," → Signature Block                                                                                                                                       |
| Biodata / Marriage Profile                      | Module N: Centered Title (+divider) → Photo (J.1) → Personal-info table (softLabel) → Family table → Centered contact → Note                                                                                                              |
| CV / Resume (Bengali or English)                | Module O: Name hero (+divider) → Contact line → Section dividers → Experience entries → Education table → Skill bullets                                                                                                                   |
| General Reports                                 | Centered title + divider → Summary metadata table → Numbered section headings with rule → Body/lists/sub-points → Tables → Conclusion → Locale-aware footer                                                                               |
| Internal Audit Reports (reference parity, §5.6) | `createDualHeader`(title, date range) → Centered "সারাংশ" heading → navyLabel metadata table → Numbered sections (`১।` + gold rule) → Body → Bengali paren list (`১)`) with `subDecimalPrefix` sub-points → `createFooter({locale:"bn"})` |
| Resumes / CVs (legacy two-column)               | Hero Title → Two-Column (Exp / Skills+Edu)                                                                                                                                                                                                |
| Reports (long formal)                           | Cover Page → TOC → Section Dividers → Data Tables → Chart Placeholder → Footer (Pagination)                                                                                                                                               |
| Invoices                                        | Hero Meta → Data Table → Callout Box (Total)                                                                                                                                                                                              |
| Contracts & Agreements                          | Hero Title → Numbered Clauses → Dual Signature Block                                                                                                                                                                                      |
| Certificates                                    | Centered Hero (Landscape) → Large Body → Dual Signature Block → Watermark/Seal                                                                                                                                                            |
| Offer Letters                                   | Header → Body Text → Data Table (Comp) → Signature Block                                                                                                                                                                                  |
| Meeting Minutes                                 | Hero Meta (Attendees/Date) → Checklist (Agenda) → Data Table (Action Items)                                                                                                                                                               |
| Proposals (Business/Project)                    | Group D: `createProposalCover` → TOC → `createProposalHeader` → Big Number Lists → Divided Lists → Timeline Table → Risk Table                                                                                                              |
| Research Papers                                 | Header → Centered Title → Abstract Callout → Body (1.5 spacing) → Footer → Bibliography                                                                                                                                                   |
| User Manuals                                    | Hero Title → TOC → Section Dividers → Two-Column (Image/Text) → Warning Callout → Numbered Steps                                                                                                                                          |
| Books & Manuscripts                             | Chapter Header → Indented Body → Drop-Cap Opener → Footer (Pagination)                                                                                                                                                                    |
| Forms                                           | Hero Title → Data Table (Input Fields) → Checklist                                                                                                                                                                                        |
| Academic Assignments                            | Header (Student Info) → Centered Title → Body Text                                                                                                                                                                                        |
| Quotations & Estimates                          | Hero Meta → Pricing Table → Terms Callout                                                                                                                                                                                                 |
| Medical Reports                                 | Hero Meta (Patient/Doctor) → Section Dividers → Diagnosis Callout → Vitals Table                                                                                                                                                          |
| Legal Documents                                 | Header → Numbered Clauses → Signature Block                                                                                                                                                                                               |
| Receipts                                        | Hero Meta → Data Table → Callout (Total Paid)                                                                                                                                                                                             |
| Purchase Orders                                 | Hero Meta (Buyer/Vendor) → Data Table → Signature Block                                                                                                                                                                                   |
| Policies & Procedures                           | Header → Scope Callout → Section Dividers → Revision History Table                                                                                                                                                                        |
| Newsletters                                     | Masthead → Multi-Column Body → Pull Quote → Image Grid                                                                                                                                                                                    |
| Press Releases                                  | Letterhead → Dateline → Body Text → Boilerplate Callout → Contact Block                                                                                                                                                                   |
| Org Charts / Directories                        | Hero Title (Landscape) → Nested Card Grid                                                                                                                                                                                                 |
| Project Timelines / Roadmaps                    | Hero Title → Timeline Table → Milestone Callouts                                                                                                                                                                                          |
| Case Studies                                    | Cover Page → Stat Callout Row → Section Dividers → Pull Quote → Body Text                                                                                                                                                                 |
| Syllabi                                         | Header (Course Info) → Data Table (Schedule) → Policy Callout                                                                                                                                                                             |
| Non-Disclosure Agreements                       | Header → Numbered Clauses (Compact) → Dual Signature Block                                                                                                                                                                                |
| Board Resolutions                               | Centered Hero → Numbered Resolutions → Signature Block                                                                                                                                                                                    |
| Administrative Registers / Field Reports (§5.5) | Centered Formal Title → Date + Metadata → Narrative Note → Primary Register Table → Grouped Sub-register Tables → Footer                                                                                                                  |
| Committee / Internal Audit Reports (long form)  | Centered Cover Block → Metadata Table → Numbered Sections → Nested Sub-point Findings → Discrepancy Table (highlighted header, zebra) → Recommendations → Committee Opinion → Multi-signatory Block → Header/Footer                       |

### 4b. Layout Groups (designPreset selection)

The rendering engine supports distinct layout pipelines selected via the `designPreset` field. Choosing the correct preset is critical — it determines which server-side renderer builds the document.

#### Group A: General Report Writing (`designPreset: "standard_business"` or `"formal_administrative_register"`)

**Trigger:** Reports, audit documents, vigilance reports, meeting minutes, field reports, administrative registers, or any document with sections, numbered headings, tables, and metadata blocks.

**Layout flow:** Dual Header (or simple header) → Centered সারাংশ / Summary heading → Navy-label metadata table → Numbered section headings with gold rule → Body paragraphs / lists / sub-points → Data tables → Signature blocks → Locale-aware footer (`পৃষ্ঠা X এর Y` / `Page X of Y`).

**JSON structure:** Uses `metadata`, `sections` (with `heading`, `paragraphs`, `bullets`, `numbered`, `table`, `signatures`), `imagePlacements`.

---

#### Group B: Formal Letter / Application (`designPreset: "formal_letter"`)

**Trigger:** Applications (দরখাস্ত), formal letters, cover letters, business correspondence, requests, complaints, or any document addressed to a specific person/authority with a date, recipient, subject, body, and signatory.

**Layout flow:** Date → Recipient Block (name, title, organization, address) → **Subject:** line (bold, accent color) → Salutation → Justified body paragraphs → Closing → Signature name (large, stylized) → Signatory title → Signatory organization → No running header → Locale-aware footer.

**JSON structure:** Uses `letterMeta` object with fields:
- `date` — the letter date (e.g. "July 16, 2025" or "০৪/০৮/২০২৬")
- `recipientLines` — array of address lines: `["Mr. Rahman Khan", "Principal", "Dhaka International School", "Dhaka, Bangladesh"]`
- `subject` — the subject text (without the "Subject:" / "বিষয়:" label)
- `salutation` — e.g. "Dear Mr. Khan," or "জনাব,"
- `closing` — e.g. "Warm regards," or "বিনীত নিবেদক,"
- `signatory` — the sender's name
- `signatoryTitle` — job title or role (optional)
- `signatoryOrg` — organization name (optional)

Body paragraphs go in `sections[0].paragraphs` as usual.

> **IMPORTANT:** When the user's content is a letter, application, or correspondence, you MUST set `designPreset` to `"formal_letter"` and populate `letterMeta`. Do NOT use the report layout for letters.

---

#### Group C: Professional CV (`designPreset: "professional_cv"`)

**Trigger:** Modern, two-column CVs or resumes, typically with a sidebar for contact/skills and a main content area for experience/projects.

**Layout flow:** Uses a full-page, two-column table. Left sidebar (dark blue background) contains Photo, Contact, Education, Skills, Languages. Right main area (white background) contains Name, Title, Profile, Work Experience, Projects.

**JSON structure:** Uses `cvMeta` object with fields:
- `name` — applicant's name
- `title` — job title
- `contact` — array of objects `{ "icon": "phone|email|link", "value": "..." }`
- `education` — array of objects `{ "degree": "...", "institution": "...", "year": "..." }`
- `skills` — array of strings
- `languages` — array of strings
- `profile` — paragraph string
- `experience` — array of objects `{ "title": "...", "company": "...", "date": "...", "bullets": ["..."] }`
- `projects` — array of objects `{ "name": "...", "description": "...", "link": "..." }`

> **IMPORTANT:** When the user's content is a professional CV or modern resume, you MUST set `designPreset` to `"professional_cv"` and populate `cvMeta`.

---

#### Group D: Business Proposal (`designPreset: "business_proposal"`)

**Trigger:** Business proposals, pitches, project plans, digital marketing strategies, or documents requiring a highly styled, modern aesthetic matching the reference images.

**Layout flow:** Proposal Cover Page → Table of Contents → Dual Header (Company / Date) → Banner Image → Summary / Objectives → Big Numbered Lists (Methodology) → Timeline Tables → Divided Lists (Stakeholders).

**JSON structure:** Uses `proposalMeta` for the cover and `sections` for the content.
- `proposalMeta`:
  - `title` — main title (e.g., "Project Proposal")
  - `subtitle` — (e.g., "Digital Marketing")
  - `preparedFor` — recipient name/company
  - `preparedBy` — sender name/company
  - `date` — proposal date (e.g., "June 2030")
- `sections` — standard sections array (can use `heading`, `image`, `paragraphs`, `bullets`, `table`, `bigNumberList`, `dividedList`).

> **IMPORTANT:** When the user wants a modern business proposal, you MUST set `designPreset` to `"business_proposal"` and populate `proposalMeta`.

---

### 4c. Generalized Classification Algorithm (for anything not on the list)

1. TOKENIZE intent: extract purpose-signaling nouns/verbs ("track", "agree", "certify", "apply", "itemize").
2. MATCH against module purposes, not document names: itemized costs → Data/Pricing Table; sign-off → Clauses + Signature; skimmability → Dividers + TOC; emphasis → Callout; chronology → Timeline/Chapter; comparison → Two-Column/Data Table; personal profile → Metadata table (biodata/CV style).
3. IF blending two archetypes, MERGE sequences in reading order (identity/header from one, core content from the other).
4. IF ambiguous, default: Header → Hero Title → Body → Section Dividers → Footer (Pagination).
5. NEVER leave content unmapped; unclassifiable prose defaults to Body Text.

## 5. Reusable Component Code Library

Shared constants (used by many modules):

```typescript
const NONE = { style: BorderStyle.NONE };
const NO_BORDERS = {
  top: NONE,
  bottom: NONE,
  left: NONE,
  right: NONE,
  insideHorizontal: NONE,
  insideVertical: NONE,
};
const CELL_MARGINS = { top: 100, bottom: 100, left: 100, right: 100 };
const gridBorders = (color: string, size = 4) => ({
  top: { style: BorderStyle.SINGLE, color, size },
  bottom: { style: BorderStyle.SINGLE, color, size },
  left: { style: BorderStyle.SINGLE, color, size },
  right: { style: BorderStyle.SINGLE, color, size },
  insideHorizontal: { style: BorderStyle.SINGLE, color, size },
  insideVertical: { style: BorderStyle.SINGLE, color, size },
});
```

### Module A: Header / Footer (simple, letterhead, dual, locale-aware)

```typescript
const createHeader = (
  title: string,
  variant: "simple" | "letterhead" = "simple",
) =>
  new Header({
    children: [
      new Paragraph({
        alignment:
          variant === "letterhead" ? AlignmentType.LEFT : AlignmentType.RIGHT,
        border: {
          bottom: {
            color: TOKENS.color.border,
            space: 1,
            style: BorderStyle.SINGLE,
            size: 6,
          },
        },
        children: [
          new TextRun({
            text: title,
            color: TOKENS.color.muted,
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
        ],
      }),
    ],
  });

// A.2 — Dual-text running header (reference parity): left title + right date range,
// one full-width divider under both. Built on a right tab-stop (no table needed).
const createDualHeader = (
  leftText: string,
  rightText: string,
  opts: { textColor?: string; ruleColor?: string; boldLeft?: boolean } = {},
) =>
  new Header({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        spacing: { after: 120 },
        border: {
          bottom: {
            color: opts.ruleColor ?? TOKENS.color.accent,
            style: BorderStyle.SINGLE,
            size: 6,
            space: 4,
          },
        },
        children: [
          new TextRun({
            text: leftText,
            bold: opts.boldLeft ?? true,
            color: opts.textColor ?? TOKENS.color.accentDark,
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
          new TextRun({
            text: "\t",
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
          new TextRun({
            text: rightText,
            color: opts.textColor ?? TOKENS.color.accentDark,
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
        ],
      }),
    ],
  });

// Proposal Header matching the reference: "Liceria Corporate" left, "June 2030" right, no border.
const createProposalHeader = (leftText: string, rightText: string) =>
  new Header({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: leftText,
            color: PROPOSAL_THEME.color.header,
            font: TOKENS.font,
            size: TOKENS.size.body,
          }),
          new TextRun({
            text: "\t",
            font: TOKENS.font,
            size: TOKENS.size.body,
          }),
          new TextRun({
            text: rightText,
            color: PROPOSAL_THEME.color.header,
            font: TOKENS.font,
            size: TOKENS.size.body,
          }),
        ],
      }),
    ],
  });

// Locale-aware footer. locale "bn" renders "পৃষ্ঠা X এর Y" (reference parity),
// under a thin top rule like the audited sample.
const createFooter = (
  opts: {
    pagination?: boolean;
    locale?: "en" | "bn";
    leftText?: string;
    rule?: boolean;
  } = {},
) => {
  const { pagination = true, locale = "en", leftText, rule = true } = opts;
  const t = (text: string) =>
    new TextRun({
      text,
      color: TOKENS.color.muted,
      font: TOKENS.font,
      size: TOKENS.size.caption,
    });
  const f = (field: typeof PageNumber.CURRENT) =>
    new TextRun({
      children: [field],
      color: TOKENS.color.muted,
      font: TOKENS.font,
      size: TOKENS.size.caption,
    });
  return new Footer({
    children: [
      new Paragraph({
        alignment: leftText ? AlignmentType.LEFT : AlignmentType.CENTER,
        ...(rule
          ? {
              border: {
                top: {
                  color: TOKENS.color.border,
                  space: 4,
                  style: BorderStyle.SINGLE,
                  size: 6,
                },
              },
            }
          : {}),
        children: [
          ...(leftText ? [t(leftText)] : []),
          ...(pagination
            ? locale === "bn"
              ? [
                  t("পৃষ্ঠা "),
                  f(PageNumber.CURRENT),
                  t(" এর "),
                  f(PageNumber.TOTAL_PAGES),
                ]
              : [
                  t(" Page "),
                  f(PageNumber.CURRENT),
                  t(" of "),
                  f(PageNumber.TOTAL_PAGES),
                ]
            : []),
        ],
      }),
    ],
  });
};
```

### Module B: Hero Metadata / Hero Title

```typescript
const createHeroMeta = (
  title: string,
  leftData: string[],
  rightData: string[],
) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                spacing: { after: 200 },
                children: [
                  new TextRun({
                    text: title,
                    bold: true,
                    size: TOKENS.size.title,
                    font: TOKENS.font,
                    color: TOKENS.color.ink,
                  }),
                ],
              }),
              ...leftData.map(
                (text) =>
                  new Paragraph({
                    children: [
                      new TextRun({
                        text,
                        size: TOKENS.size.body,
                        font: TOKENS.font,
                        color: TOKENS.color.subink,
                      }),
                    ],
                  }),
              ),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: rightData.map(
              (text) =>
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun({
                      text,
                      size: TOKENS.size.body,
                      font: TOKENS.font,
                      color: TOKENS.color.subink,
                    }),
                  ],
                }),
            ),
          }),
        ],
      }),
    ],
  });

const createHeroTitleCentered = (title: string, subtitle?: string) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: subtitle ? 100 : 400 },
    children: [
      new TextRun({
        text: title,
        bold: true,
        size: TOKENS.size.title + 12,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  }),
  ...(subtitle
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [
            new TextRun({
              text: subtitle,
              size: TOKENS.size.h2,
              font: TOKENS.font,
              color: TOKENS.color.subink,
            }),
          ],
        }),
      ]
    : []),
];
```

### Module C: Centered Cover Page

```typescript
const createCenteredCoverPage = (
  organization: string,
  mainTitle: string,
  subtitle?: string,
  dateRange?: string,
  footerNote?: string,
) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 220, after: 120 },
    children: [
      new TextRun({
        text: organization,
        bold: true,
        size: TOKENS.size.caption,
        font: TOKENS.font,
        color: TOKENS.color.accentDark,
      }),
    ],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 160 },
    border: {
      bottom: {
        color: TOKENS.color.accent,
        style: BorderStyle.SINGLE,
        size: 12,
      },
    },
    children: [new TextRun({ text: "" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 140, after: subtitle || dateRange ? 120 : 280 },
    children: [
      new TextRun({
        text: mainTitle,
        bold: true,
        size: TOKENS.size.title + 8,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  }),
  ...(subtitle
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: subtitle,
              size: TOKENS.size.h2,
              font: TOKENS.font,
              color: TOKENS.color.subink,
            }),
          ],
        }),
      ]
    : []),
  ...(dateRange
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: dateRange,
              italics: true,
              size: TOKENS.size.body,
              font: TOKENS.font,
              color: TOKENS.color.subink,
            }),
          ],
        }),
      ]
    : []),
  ...(footerNote
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200 },
          children: [
            new TextRun({
              text: footerNote,
              size: TOKENS.size.caption,
              font: TOKENS.font,
              color: TOKENS.color.muted,
            }),
          ],
        }),
      ]
    : []),
];

const createProposalCover = (
  title: string,
  subtitle: string,
  preparedFor: string,
  preparedBy: string,
  date: string,
) => [
  createSpacer(1000),
  new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text: title,
        bold: true,
        size: TOKENS.size.title + 24,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.heading,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 600 },
    children: [
      new TextRun({
        text: subtitle,
        size: TOKENS.size.h1,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.heading,
      }),
    ],
  }),
  createSpacer(2000),
  new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({
        text: "Prepared for:",
        bold: true,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.heading,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 400 },
    children: [
      new TextRun({
        text: preparedFor,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.subink,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({
        text: "Prepared by:",
        bold: true,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.heading,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 400 },
    children: [
      new TextRun({
        text: preparedBy,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.subink,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({
        text: date,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: PROPOSAL_THEME.color.subink,
      }),
    ],
  }),
];
```

### Module C.1: Metadata / Summary Table (three variants) + report styling

```typescript
// navyLabel  → official summary table of the reference sample (navy label column,
//              white bold labels, visible grid, zebra value rows).
// accentFirst→ legacy corporate variant (row-0 accent band).
// softLabel  → biodata/profile variant (light label column, no shouting).
const createMetadataTable = (
  rows: { label: string; value: string }[],
  variant: "navyLabel" | "accentFirst" | "softLabel" = "accentFirst",
  labelWidthPct = 35,
) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders:
      variant === "navyLabel"
        ? gridBorders(AUDIT_THEME.color.grid, 4)
        : NO_BORDERS,
    rows: rows.map((row, index) => {
      const labelFill =
        variant === "navyLabel"
          ? AUDIT_THEME.color.labelFill
          : variant === "accentFirst"
            ? index === 0
              ? TOKENS.color.accent
              : TOKENS.color.zebraA
            : TOKENS.color.zebraA;
      const labelColor =
        variant === "softLabel"
          ? TOKENS.color.subink
          : variant === "accentFirst" && index === 0
            ? TOKENS.color.white
            : AUDIT_THEME.color.labelText;
      const valueFill =
        variant === "navyLabel"
          ? index % 2 === 0
            ? AUDIT_THEME.color.zebra
            : TOKENS.color.white
          : variant === "accentFirst"
            ? index === 0
              ? TOKENS.color.accentBg
              : TOKENS.color.zebraB
            : TOKENS.color.white;
      const margins =
        variant === "navyLabel" ? FORMAL_REGISTER.cellMargins : CELL_MARGINS;
      return new TableRow({
        children: [
          new TableCell({
            width: { size: labelWidthPct, type: WidthType.PERCENTAGE },
            shading: { fill: labelFill, type: ShadingType.CLEAR },
            margins,
            children: [
              new Paragraph({
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: row.label,
                    bold: true,
                    font: TOKENS.font,
                    size: TOKENS.size.body,
                    color: labelColor,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 100 - labelWidthPct, type: WidthType.PERCENTAGE },
            shading: { fill: valueFill, type: ShadingType.CLEAR },
            margins,
            children: [
              new Paragraph({
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: row.value,
                    font: TOKENS.font,
                    size: TOKENS.size.body,
                    color: TOKENS.color.ink,
                  }),
                ],
              }),
            ],
          }),
        ],
      });
    }),
  });

const createReportSectionHeading = (prefix: string, title: string) => [
  new Paragraph({
    spacing: { before: 300, after: 120 },
    border: {
      bottom: {
        color: TOKENS.color.accent,
        style: BorderStyle.SINGLE,
        size: 8,
      },
    },
    children: [
      new TextRun({
        text: `${prefix} ${title}`,
        bold: true,
        font: TOKENS.font,
        size: TOKENS.size.h2,
        color: TOKENS.color.ink,
      }),
    ],
  }),
];

const createReportBulletList = (items: string[]) =>
  items.map(
    (item) =>
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new TextRun({
            text: "• ",
            font: TOKENS.font,
            size: TOKENS.size.body,
            color: TOKENS.color.accent,
          }),
          new TextRun({
            text: item,
            font: TOKENS.font,
            size: TOKENS.size.body,
            color: TOKENS.color.subink,
          }),
        ],
      }),
  );

const createReportNestedObservationList = (
  items: {
    prefix: string;
    text: string;
    children?: { prefix: string; text: string }[];
  }[],
) =>
  items.flatMap((item) => [
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: `${item.prefix} `,
          bold: true,
          font: TOKENS.font,
          size: TOKENS.size.body,
          color: TOKENS.color.accentDark,
        }),
        new TextRun({
          text: item.text,
          font: TOKENS.font,
          size: TOKENS.size.body,
          color: TOKENS.color.subink,
        }),
      ],
    }),
    ...(item.children ?? []).map(
      (child) =>
        new Paragraph({
          indent: { left: 720 },
          spacing: { after: 80 },
          children: [
            new TextRun({
              text: `${child.prefix} `,
              bold: true,
              font: TOKENS.font,
              size: TOKENS.size.body,
              color: TOKENS.color.accent,
            }),
            new TextRun({
              text: child.text,
              font: TOKENS.font,
              size: TOKENS.size.body,
              color: TOKENS.color.subink,
            }),
          ],
        }),
    ),
  ]);
```

### Module C.2: Typography Primitives

Title alignment is chosen per archetype, never defaulted: formal/official/institutional (registers, notices, certificates, committee & audit reports, cover pages, academic title pages) → CENTER; correspondence/memos/letters/manuals → LEFT. Pass it explicitly.

```typescript
const createDocTitle = (
  text: string,
  align: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
  withDivider = false,
) =>
  new Paragraph({
    alignment: align,
    spacing: { before: 120, after: withDivider ? 60 : 80 },
    heading: HeadingLevel.TITLE,
    ...(withDivider
      ? {
          border: {
            bottom: {
              color: TOKENS.color.accent,
              style: BorderStyle.SINGLE,
              size: 10,
              space: 8,
            },
          },
        }
      : {}),
    children: [
      new TextRun({
        text,
        bold: true,
        size: TOKENS.size.title,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  });

const createDocSubtitle = (
  text: string,
  align: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
) =>
  new Paragraph({
    alignment: align,
    spacing: { after: 200 },
    children: [
      new TextRun({
        text,
        size: TOKENS.size.subtitle,
        font: TOKENS.font,
        color: TOKENS.color.subink,
        italics: true,
      }),
    ],
  });

const createHeading = (text: string, level: 1 | 2 | 3 | 4) => {
  const sizeMap = {
    1: TOKENS.size.h1,
    2: TOKENS.size.h2,
    3: TOKENS.size.h3,
    4: TOKENS.size.h4,
  };
  const headingMap = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
  };
  return new Paragraph({
    heading: headingMap[level],
    spacing: { before: level <= 2 ? 300 : 200, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: sizeMap[level],
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  });
};

// Centered mid-page heading (e.g. "প্রতিবেদনের সারাংশঃ" in the reference sample).
const createCenteredHeading = (
  text: string,
  opts: { size?: number; color?: string; after?: number } = {},
) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: opts.after ?? 160 },
    children: [
      new TextRun({
        text,
        bold: true,
        font: TOKENS.font,
        size: opts.size ?? TOKENS.size.h2,
        color: opts.color ?? TOKENS.color.ink,
      }),
    ],
  });

const createLeadParagraph = (text: string) =>
  new Paragraph({
    spacing: { after: 200 },
    alignment: AlignmentType.JUSTIFIED,
    children: [
      new TextRun({
        text,
        size: TOKENS.size.lead,
        font: TOKENS.font,
        color: TOKENS.color.subink,
      }),
    ],
  });

const createBodyParagraph = (text: string, justified = true) =>
  new Paragraph({
    spacing: { after: 160, line: 300 },
    alignment: justified ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    children: [
      new TextRun({
        text,
        size: TOKENS.size.body,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  });

const createSmallParagraph = (text: string, muted = true) =>
  new Paragraph({
    spacing: { after: 100 },
    children: [
      new TextRun({
        text,
        size: TOKENS.size.small,
        font: TOKENS.font,
        color: muted ? TOKENS.color.muted : TOKENS.color.subink,
      }),
    ],
  });

// Inline emphasis runs — compose inside any Paragraph's children.
const runBold = (
  text: string,
  extra: Partial<ConstructorParameters<typeof TextRun>[0]> = {},
) =>
  new TextRun({
    text,
    bold: true,
    font: TOKENS.font,
    size: TOKENS.size.body,
    color: TOKENS.color.ink,
    ...extra,
  });
const runItalic = (
  text: string,
  extra: Partial<ConstructorParameters<typeof TextRun>[0]> = {},
) =>
  new TextRun({
    text,
    italics: true,
    font: TOKENS.font,
    size: TOKENS.size.body,
    color: TOKENS.color.subink,
    ...extra,
  });
const runUnderline = (
  text: string,
  extra: Partial<ConstructorParameters<typeof TextRun>[0]> = {},
) =>
  new TextRun({
    text,
    underline: {},
    font: TOKENS.font,
    size: TOKENS.size.body,
    color: TOKENS.color.ink,
    ...extra,
  });
const runHighlight = (
  text: string,
  color: "yellow" | "cyan" | "green" | "red" = "yellow",
) =>
  new TextRun({
    text,
    highlight: color,
    font: TOKENS.font,
    size: TOKENS.size.body,
    color: TOKENS.color.ink,
  });
const runPlain = (
  text: string,
  extra: Partial<ConstructorParameters<typeof TextRun>[0]> = {},
) =>
  new TextRun({
    text,
    font: TOKENS.font,
    size: TOKENS.size.body,
    color: TOKENS.color.ink,
    ...extra,
  });

const createMixedParagraph = (
  tokens: {
    text: string;
    style?: "bold" | "italic" | "underline" | "highlight" | "plain";
  }[],
  opts: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    spacingAfter?: number;
  } = {},
) =>
  new Paragraph({
    alignment: opts.alignment ?? AlignmentType.JUSTIFIED,
    spacing: { after: opts.spacingAfter ?? 160 },
    children: tokens.map((t) => {
      switch (t.style) {
        case "bold":
          return runBold(t.text);
        case "italic":
          return runItalic(t.text);
        case "underline":
          return runUnderline(t.text);
        case "highlight":
          return runHighlight(t.text);
        default:
          return runPlain(t.text);
      }
    }),
  });
```

### Module C.3: N-Level Nested Outline Engine

```typescript
interface OutlinePoint {
  prefix: string;
  text: string;
  children?: OutlinePoint[];
}

const createNestedOutline = (
  points: OutlinePoint[],
  depth = 0,
  indentPerLevel = 480,
): Paragraph[] =>
  points.flatMap((point) => [
    new Paragraph({
      indent: depth > 0 ? { left: depth * indentPerLevel } : undefined,
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: `${point.prefix} `,
          bold: true,
          font: TOKENS.font,
          size: TOKENS.size.body,
          color: depth === 0 ? TOKENS.color.accentDark : TOKENS.color.accent,
        }),
        new TextRun({
          text: point.text,
          font: TOKENS.font,
          size: TOKENS.size.body,
          color: TOKENS.color.subink,
        }),
      ],
    }),
    ...(point.children
      ? createNestedOutline(point.children, depth + 1, indentPerLevel)
      : []),
  ]);
```

Usage patterns: source uses `১।` → `bengaliDaariPrefix`; `১)` → `bengaliParenPrefix`; `1.1 / ১.১` sub-points → `subDecimalPrefix(parent, child, lang)`; `a)` → `letteredPrefix`; `i.` → `romanPrefix`. Mix freely per level; the engine only indents and styles.

### Module C.4: Rule / Divider Variants

```typescript
const createRuleLine = (
  color = TOKENS.color.accent,
  weight: "thin" | "thick" = "thin",
) =>
  new Paragraph({
    spacing: { before: 100, after: 200 },
    border: {
      bottom: {
        color,
        style: BorderStyle.SINGLE,
        size: weight === "thin" ? 6 : 16,
      },
    },
    children: [new TextRun({ text: "" })],
  });

const createGlyphDivider = (
  glyph = "▬",
  count = 30,
  color = TOKENS.color.muted,
) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [
      new TextRun({
        text: glyph.repeat(count),
        font: TOKENS.font,
        color,
        size: TOKENS.size.small,
      }),
    ],
  });

const createDoubleRule = (color = TOKENS.color.ink) =>
  new Paragraph({
    spacing: { before: 60, after: 200 },
    border: { bottom: { color, style: BorderStyle.DOUBLE, size: 6 } },
    children: [new TextRun({ text: "" })],
  });
```

### Module C.5: Numbered Section Heading with Rule (reference parity)

Produces `১। নিরীক্ষার উদ্দেশ্য`-style headings: bold prefix + title, then a thin full-width rule (gold in audit theme) with a small gap — exactly the reference sample's section marker.

```typescript
const createNumberedSectionHeading = (
  prefix: string,
  title: string,
  opts: {
    color?: string;
    ruleColor?: string;
    font?: string;
    size?: number;
    toc?: 1 | 2 | 3;
  } = {},
) => [
  new Paragraph({
    ...(opts.toc
      ? {
          heading: (
            [
              HeadingLevel.HEADING_1,
              HeadingLevel.HEADING_2,
              HeadingLevel.HEADING_3,
            ] as const
          )[opts.toc - 1],
        }
      : {}),
    spacing: { before: 300, after: 60 },
    children: [
      new TextRun({
        text: `${prefix} `,
        bold: true,
        color: opts.color ?? TOKENS.color.ink,
        font: opts.font ?? TOKENS.font,
        size: opts.size ?? TOKENS.size.h2,
      }),
      new TextRun({
        text: title,
        bold: true,
        color: opts.color ?? TOKENS.color.ink,
        font: opts.font ?? TOKENS.font,
        size: opts.size ?? TOKENS.size.h2,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 160 },
    border: {
      bottom: {
        color: opts.ruleColor ?? TOKENS.color.accent,
        style: BorderStyle.SINGLE,
        size: 6,
        space: 4,
      },
    },
    children: [new TextRun({ text: "" })],
  }),
];
```

### Module D: Data Table / Pricing Table

```typescript
const createDataTable = (headers: string[], rows: string[][]) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(
          (h) =>
            new TableCell({
              shading: { fill: TOKENS.color.dark, type: ShadingType.CLEAR },
              margins: CELL_MARGINS,
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: h,
                      bold: true,
                      color: TOKENS.color.white,
                      font: TOKENS.font,
                    }),
                  ],
                }),
              ],
            }),
        ),
      }),
      ...rows.map(
        (row, index) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  shading: {
                    fill:
                      index % 2 === 0
                        ? TOKENS.color.zebraA
                        : TOKENS.color.zebraB,
                    type: ShadingType.CLEAR,
                  },
                  borders: {
                    bottom: {
                      style: BorderStyle.SINGLE,
                      color: TOKENS.color.border,
                      size: 4,
                    },
                    top: NONE,
                    left: NONE,
                    right: NONE,
                  },
                  margins: CELL_MARGINS,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: cell,
                          font: TOKENS.font,
                          color: TOKENS.color.subink,
                        }),
                      ],
                    }),
                  ],
                }),
            ),
          }),
      ),
    ],
  });

const createPricingTable = (
  headers: string[],
  rows: string[][],
  totalRow: string[],
) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(
          (h) =>
            new TableCell({
              shading: { fill: TOKENS.color.dark, type: ShadingType.CLEAR },
              margins: CELL_MARGINS,
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: h,
                      bold: true,
                      color: TOKENS.color.white,
                      font: TOKENS.font,
                    }),
                  ],
                }),
              ],
            }),
        ),
      }),
      ...rows.map(
        (row, index) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  shading: {
                    fill:
                      index % 2 === 0
                        ? TOKENS.color.zebraA
                        : TOKENS.color.zebraB,
                    type: ShadingType.CLEAR,
                  },
                  margins: CELL_MARGINS,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: cell,
                          font: TOKENS.font,
                          color: TOKENS.color.subink,
                        }),
                      ],
                    }),
                  ],
                }),
            ),
          }),
      ),
      new TableRow({
        children: totalRow.map(
          (cell, i) =>
            new TableCell({
              shading: { fill: TOKENS.color.accentBg, type: ShadingType.CLEAR },
              margins: CELL_MARGINS,
              children: [
                new Paragraph({
                  alignment:
                    i === totalRow.length - 1
                      ? AlignmentType.RIGHT
                      : AlignmentType.LEFT,
                  children: [
                    new TextRun({
                      text: cell,
                      bold: true,
                      font: TOKENS.font,
                      color: TOKENS.color.accentDark,
                    }),
                  ],
                }),
              ],
            }),
        ),
      }),
    ],
  });
```

### Module D.1: Two-Column Layout

```typescript
const createTwoColumnLayout = (
  leftContent: Paragraph[],
  rightContent: Paragraph[],
  leftWidthPct = 70,
) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: leftWidthPct, type: WidthType.PERCENTAGE },
            margins: { right: 300 },
            children: leftContent,
          }),
          new TableCell({
            width: { size: 100 - leftWidthPct, type: WidthType.PERCENTAGE },
            shading: { fill: TOKENS.color.zebraA, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            children: rightContent,
          }),
        ],
      }),
    ],
  });
```

### Module E: Signature Block

```typescript
const createSignatureBlock = (signatories: string[], includeDateLine = true) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: signatories.map(
          (name) =>
            new TableCell({
              width: {
                size: 100 / signatories.length,
                type: WidthType.PERCENTAGE,
              },
              margins: { top: 600, right: 400 },
              children: [
                new Paragraph({
                  border: {
                    bottom: {
                      color: "000000",
                      space: 1,
                      style: BorderStyle.SINGLE,
                      size: 6,
                    },
                  },
                  children: [new TextRun({ text: "" })],
                }),
                new Paragraph({
                  spacing: { before: 100 },
                  children: [
                    new TextRun({ text: name, bold: true, font: TOKENS.font }),
                  ],
                }),
                ...(includeDateLine
                  ? [
                      new Paragraph({
                        spacing: { before: 300 },
                        border: {
                          bottom: {
                            color: "000000",
                            space: 1,
                            style: BorderStyle.SINGLE,
                            size: 6,
                          },
                        },
                        children: [new TextRun({ text: "Date" })],
                      }),
                    ]
                  : []),
              ],
            }),
        ),
      }),
    ],
  });
```

### Module F: Callout Box (tone variants)

Box-integrity checklist (a "corrupted highlight box" almost always violates one of these):

1. Full text-column width — never nested beside a table or inside a narrow column.
2. One consistent run style per box (bold label run + regular body run, same resolved font); no third ad-hoc style.
3. Justify multi-line box text so the right edge doesn't rag against the border.
4. Generous margins (≥150 twips top/bottom), never a fixed row height — Bengali conjuncts are tall and will clip.
5. Never truncate; let long text wrap inside the cell.

```typescript
const CALLOUT_THEMES = {
  info: {
    border: TOKENS.color.accent,
    bg: TOKENS.color.accentBg,
    text: TOKENS.color.accentDark,
  },
  warning: { border: "D97706", bg: "FFFBEB", text: "92400E" },
  danger: {
    border: TOKENS.color.danger,
    bg: TOKENS.color.dangerBg,
    text: TOKENS.color.danger,
  },
  success: {
    border: TOKENS.color.success,
    bg: TOKENS.color.successBg,
    text: TOKENS.color.success,
  },
};

const createCalloutBox = (
  text: string,
  theme: keyof typeof CALLOUT_THEMES = "info",
  label?: string,
) => {
  const t = CALLOUT_THEMES[theme];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: NONE,
      bottom: NONE,
      right: NONE,
      left: { style: BorderStyle.SINGLE, color: t.border, size: 24 },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: t.bg, type: ShadingType.CLEAR },
            margins: { top: 180, bottom: 180, left: 220, right: 220 },
            children: [
              ...(label
                ? [
                    new Paragraph({
                      spacing: { after: 60 },
                      children: [
                        new TextRun({
                          text: label,
                          bold: true,
                          font: TOKENS.font,
                          size: TOKENS.size.body,
                          color: t.text,
                        }),
                      ],
                    }),
                  ]
                : []),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                spacing: { line: 300 },
                children: [
                  new TextRun({
                    text,
                    font: TOKENS.font,
                    size: TOKENS.size.body,
                    color: t.text,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
};
```

### Module G / G.1: Section Dividers & Spacing Utilities

A divider between every major section of a formal report is mandatory, not decorative.

```typescript
const createSectionDivider = (label: string) => [
  new Paragraph({
    spacing: { before: 400, after: 150 },
    border: {
      bottom: {
        color: TOKENS.color.border,
        space: 4,
        style: BorderStyle.SINGLE,
        size: 6,
      },
    },
    children: [
      new TextRun({
        text: smartUpper(label),
        bold: true,
        size: TOKENS.size.h1,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  }),
];

const createSpacer = (heightTwips = 200) =>
  new Paragraph({
    spacing: { before: heightTwips, after: 0 },
    children: [new TextRun({ text: "" })],
  });

const createVerticalSpace = (before = 0, after = 200) =>
  new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text: "" })],
  });

const createLineDivider = (
  color = TOKENS.color.border,
  weight: "thin" | "thick" = "thin",
  space = 120,
) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: space, after: space },
    border: {
      bottom: {
        color,
        style: BorderStyle.SINGLE,
        size: weight === "thin" ? 6 : 16,
      },
    },
    children: [new TextRun({ text: "" })],
  });

const createDecorativeDivider = (
  glyph = "▬",
  count = 30,
  color = TOKENS.color.muted,
  before = 80,
  after = 200,
) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before, after },
    children: [
      new TextRun({
        text: glyph.repeat(count),
        font: TOKENS.font,
        color,
        size: TOKENS.size.small,
      }),
    ],
  });
```

Prefer `createSpacer`/`createVerticalSpace` between table modules; `createLineDivider` for functional breaks; `createDecorativeDivider` only on covers (formal Bengali reports often use a glyph row under the cover title). Never inline raw twip values.

### Module H: Checklist / Numbered Steps

```typescript
const createChecklist = (items: string[]) =>
  items.map(
    (item) =>
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({
            text: "☐ ",
            font: TOKENS.font,
            size: TOKENS.size.body,
          }),
          new TextRun({
            text: item,
            font: TOKENS.font,
            size: TOKENS.size.body,
            color: TOKENS.color.subink,
          }),
        ],
      }),
  );

const createNumberedSteps = (steps: string[]) =>
  steps.map(
    (step, i) =>
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: `${i + 1}. `,
            bold: true,
            font: TOKENS.font,
            color: TOKENS.color.accent,
          }),
          new TextRun({
            text: step,
            font: TOKENS.font,
            color: TOKENS.color.subink,
          }),
        ],
      }),
  );

const createBigNumberList = (
  items: { number: string; title: string; text: string }[],
) =>
  items.flatMap((item) => [
    new Paragraph({
      spacing: { before: 200, after: 40 },
      children: [
        new TextRun({
          text: item.number,
          size: TOKENS.size.title + 10,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.accent,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: item.title,
          bold: true,
          size: TOKENS.size.h2,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.heading,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 160 },
      children: [
        new TextRun({
          text: item.text,
          size: TOKENS.size.body,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.subink,
        }),
      ],
    }),
  ]);

const createDividedList = (items: { title: string; subtitle: string }[]) =>
  items.flatMap((item) => [
    new Paragraph({
      spacing: { before: 160, after: 40 },
      children: [
        new TextRun({
          text: "→  ",
          bold: true,
          size: TOKENS.size.h3,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.heading,
        }),
        new TextRun({
          text: item.title,
          bold: true,
          size: TOKENS.size.h3,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.heading,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 160 },
      indent: { left: 400 },
      children: [
        new TextRun({
          text: item.subtitle,
          size: TOKENS.size.body,
          font: TOKENS.font,
          color: PROPOSAL_THEME.color.subink,
        }),
      ],
    }),
    new Paragraph({
      border: {
        bottom: {
          color: PROPOSAL_THEME.color.rule,
          style: BorderStyle.SINGLE,
          size: 6,
        },
      },
      children: [new TextRun({ text: "" })],
    }),
  ]);
```

### Module I: Table of Contents

```typescript
const createTOC = () =>
  new TableOfContents("Table of Contents", {
    hyperlink: true,
    headingStyleRange: "1-3",
  });
// Body headings must use HeadingLevel.HEADING_1/2/3 or the field renders empty until manually updated.
```

### Module J: Image Layout Library

J.1 centered+caption (figures) · J.2 full-width banner (covers) · J.3 floating wrap left/right (manuals) · J.4 side-by-side grid (comparisons) · J.5 watermark/seal · J.6 letterhead lockup. Selection rule: match the image's relationship to text; never default everything to centered.

```typescript
const createImageBlock = (
  imageBuffer: Uint8Array,
  width: number,
  height: number,
  caption?: string,
) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new ImageRun({ data: imageBuffer, transformation: { width, height } }),
    ],
  }),
  ...(caption
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80 },
          children: [
            new TextRun({
              text: caption,
              italics: true,
              size: TOKENS.size.caption,
              color: TOKENS.color.muted,
              font: TOKENS.font,
            }),
          ],
        }),
      ]
    : []),
];

const createBannerImage = (
  imageBuffer: Uint8Array,
  pageWidthTwips: number,
  heightPx: number,
) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [
      new ImageRun({
        data: imageBuffer,
        transformation: {
          width: Math.round(pageWidthTwips / 15),
          height: heightPx,
        },
      }),
    ],
  });

const createFloatingImage = (
  imageBuffer: Uint8Array,
  width: number,
  height: number,
  side: "left" | "right" = "right",
) =>
  new ImageRun({
    data: imageBuffer,
    transformation: { width, height },
    floating: {
      horizontalPosition: { relative: "column" as any, align: side as any },
      verticalPosition: { relative: "paragraph" as any, offset: 0 },
      wrap: {
        type: "square" as any,
        side: (side === "left" ? "right" : "left") as any,
      },
      margins: { left: 200000, right: 200000, top: 0, bottom: 100000 },
    },
  });

const createFloatingImageParagraph = (
  imageBuffer: Uint8Array,
  width: number,
  height: number,
  bodyText: string,
  side: "left" | "right" = "right",
) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    children: [
      createFloatingImage(imageBuffer, width, height, side),
      new TextRun({
        text: bodyText,
        font: TOKENS.font,
        size: TOKENS.size.body,
        color: TOKENS.color.ink,
      }),
    ],
  });

const createImageGrid = (
  images: {
    buffer: Uint8Array;
    width: number;
    height: number;
    caption?: string;
  }[],
  columns: 2 | 3 | 4 = 2,
) => {
  const rows: TableRow[] = [];
  for (let i = 0; i < images.length; i += columns) {
    const rowImages = images.slice(i, i + columns);
    rows.push(
      new TableRow({
        children: [
          ...rowImages.map(
            (img) =>
              new TableCell({
                width: { size: 100 / columns, type: WidthType.PERCENTAGE },
                margins: CELL_MARGINS,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new ImageRun({
                        data: img.buffer,
                        transformation: {
                          width: img.width,
                          height: img.height,
                        },
                      }),
                    ],
                  }),
                  ...(img.caption
                    ? [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 60 },
                          children: [
                            new TextRun({
                              text: img.caption,
                              italics: true,
                              size: TOKENS.size.caption,
                              color: TOKENS.color.muted,
                              font: TOKENS.font,
                            }),
                          ],
                        }),
                      ]
                    : []),
                ],
              }),
          ),
          ...Array.from(
            { length: columns - rowImages.length },
            () => new TableCell({ children: [] }),
          ),
        ],
      }),
    );
  }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows,
  });
};

const createWatermarkImage = (
  imageBuffer: Uint8Array,
  width: number,
  height: number,
) =>
  new ImageRun({
    data: imageBuffer,
    transformation: { width, height },
    floating: {
      horizontalPosition: { relative: "page" as any, align: "center" as any },
      verticalPosition: { relative: "page" as any, align: "center" as any },
      wrap: { type: "none" as any },
      behindDocument: true,
      // pre-process source image to desired opacity; docx does not apply opacity at render time.
    },
  });

const createLetterheadLockup = (
  logoBuffer: Uint8Array,
  logoWidth: number,
  logoHeight: number,
  orgName: string,
  tagline?: string,
) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    data: logoBuffer,
                    transformation: { width: logoWidth, height: logoHeight },
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 85, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: orgName,
                    bold: true,
                    size: TOKENS.size.h1,
                    font: TOKENS.font,
                    color: TOKENS.color.ink,
                  }),
                ],
              }),
              ...(tagline
                ? [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: tagline,
                          italics: true,
                          size: TOKENS.size.caption,
                          font: TOKENS.font,
                          color: TOKENS.color.muted,
                        }),
                      ],
                    }),
                  ]
                : []),
            ],
          }),
        ],
      }),
    ],
  });
```

### Module K: Timeline Table · Module L: Pull Quote

```typescript
const createTimelineTable = (
  milestones: { date: string; title: string; status: string }[],
) =>
  createDataTable(
    ["Date", "Milestone", "Status"],
    milestones.map((m) => [m.date, m.title, m.status]),
  );

const createPullQuote = (quote: string, attribution?: string) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 300, after: attribution ? 80 : 300 },
    children: [
      new TextRun({
        text: `“${quote}”`,
        italics: true,
        size: TOKENS.size.h2,
        font: TOKENS.font,
        color: TOKENS.color.ink,
      }),
    ],
  }),
  ...(attribution
    ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [
            new TextRun({
              text: `— ${attribution}`,
              size: TOKENS.size.caption,
              font: TOKENS.font,
              color: TOKENS.color.muted,
            }),
          ],
        }),
      ]
    : []),
];
```

### Module M: Formal Letter / Application (দরখাস্ত)

Canonical order for Bengali formal applications: date → "বরাবর" recipient block → bold "বিষয়:" line → salutation ("জনাব,") → justified body → closing ("বিনীত নিবেদক,") → signatory. English business letters use the same module with `subjectLabel: "Subject:"`, closing "Sincerely,". Title alignment for any letterhead text stays LEFT.

```typescript
const createFormalLetter = (spec: {
  date: string;
  dateAlign?: "left" | "right";
  recipient: string[]; // ["বরাবর", "প্রধান শিক্ষক", "ঢাকা বিশ্ববিদ্যালয়, ঢাকা"]
  subject: string;
  subjectLabel?: string; // default "বিষয়:"
  salutation?: string; // "জনাব," / "Dear Sir,"
  body: string[];
  closing: string; // "বিনীত নিবেদক," / "Sincerely,"
  signatory: string;
  signatoryMeta?: string;
  enclosures?: string[];
}) => [
  new Paragraph({
    alignment:
      spec.dateAlign === "right" ? AlignmentType.RIGHT : AlignmentType.LEFT,
    spacing: { after: 200 },
    children: [runPlain(spec.date)],
  }),
  ...spec.recipient.map(
    (line, i) =>
      new Paragraph({
        spacing: { after: 40 },
        children: [i === 0 ? runBold(line) : runPlain(line)],
      }),
  ),
  new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [runBold(`${spec.subjectLabel ?? "বিষয়:"} ${spec.subject}`)],
  }),
  ...(spec.salutation
    ? [
        new Paragraph({
          spacing: { after: 160 },
          children: [runPlain(spec.salutation)],
        }),
      ]
    : []),
  ...spec.body.map((p) => createBodyParagraph(p)),
  createSpacer(300),
  new Paragraph({ children: [runPlain(spec.closing)] }),
  new Paragraph({
    spacing: { before: 80 },
    children: [runBold(spec.signatory)],
  }),
  ...(spec.signatoryMeta
    ? [new Paragraph({ children: [runItalic(spec.signatoryMeta)] })]
    : []),
  ...(spec.enclosures?.length
    ? [createSmallParagraph(`সংযুক্তি: ${spec.enclosures.join(", ")}`)]
    : []),
];
```

### Module N: Biodata / Marriage Profile

```typescript
const createBiodata = (spec: {
  title: string; // "বায়োডাটা" / "BIODATA"
  photo?: { buffer: Uint8Array; width: number; height: number };
  personal: { label: string; value: string }[]; // নাম, পিতার নাম, জন্ম তারিখ, উচ্চতা, ধর্ম…
  familyHeader?: string[];
  familyRows?: string[][];
  contact: string[];
  note?: string;
}) => [
  createDocTitle(spec.title, AlignmentType.CENTER, true),
  ...(spec.photo
    ? createImageBlock(spec.photo.buffer, spec.photo.width, spec.photo.height)
    : []),
  createSpacer(),
  createMetadataTable(spec.personal, "softLabel"),
  ...(spec.familyHeader && spec.familyRows
    ? [createSpacer(), createDataTable(spec.familyHeader, spec.familyRows)]
    : []),
  createSpacer(),
  ...spec.contact.map(
    (c) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [runPlain(c, { color: TOKENS.color.subink })],
      }),
  ),
  ...(spec.note ? [createSmallParagraph(spec.note)] : []),
];
```

### Module O: CV / Resume

```typescript
const createResumeHeader = (
  name: string,
  contactLine: string,
  centered = true,
) => [
  createDocTitle(
    name,
    centered ? AlignmentType.CENTER : AlignmentType.LEFT,
    true,
  ),
  new Paragraph({
    alignment: centered ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { after: 240 },
    children: [
      new TextRun({
        text: contactLine,
        color: TOKENS.color.subink,
        font: TOKENS.font,
        size: TOKENS.size.small,
      }),
    ],
  }),
];

const createResumeEntry = (title: string, meta: string, bullets: string[]) => [
  new Paragraph({
    spacing: { before: 160, after: 40 },
    children: [runBold(title)],
  }),
  new Paragraph({ spacing: { after: 80 }, children: [runItalic(meta)] }),
  ...createReportBulletList(bullets),
];
// Assemble: createResumeHeader → createSectionDivider("EXPERIENCE") → entries →
// createSectionDivider("EDUCATION") → createDataTable → skills bullets.
// Bengali CVs: pass Bengali strings verbatim; section labels go through smartUpper-safe dividers.
```

### 5.5 Formal Administrative Register / Field Report Layout

For attendance sheets, beneficiary/student/patient lists, field-monitoring reports, ward registers. Use `FORMAL_REGISTER`: centered navy title with fine rule (`createFormalRegisterHeading` below), compact labeled metadata paragraphs, small blue section labels (`createRegisterSectionLabel`), high-contrast data tables with dark navy header row, thin visible grid, explicit DXA widths (narrow no./age/date columns, wide name/address columns), centered short identifiers, left-aligned names, very light zebra rows, repeated headers across pages, no fixed row heights, body type never below 8.5 pt. Avoid gradients, oversized covers, large callouts. Selection trigger: register, attendance, beneficiary list, field visit, group-wise records, official report, multiple compact people tables.

```typescript
const createFormalRegisterHeading = (title: string, period?: string) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: FORMAL_REGISTER.spacing.titleAfter },
    border: {
      bottom: {
        color: FORMAL_REGISTER.color.rule,
        style: BorderStyle.SINGLE,
        size: 8,
        space: 10,
      },
    },
    children: [
      new TextRun({
        text: period ? `${title} — ${period}` : title,
        bold: true,
        font: FORMAL_REGISTER.font,
        size: FORMAL_REGISTER.size.title,
        color: FORMAL_REGISTER.color.ink,
      }),
    ],
  }),
];

const createRegisterSectionLabel = (label: string) =>
  new Paragraph({
    spacing: {
      before: FORMAL_REGISTER.spacing.sectionBefore,
      after: FORMAL_REGISTER.spacing.sectionAfter,
    },
    children: [
      new TextRun({
        text: label,
        bold: true,
        font: FORMAL_REGISTER.font,
        size: FORMAL_REGISTER.size.section,
        color: FORMAL_REGISTER.color.heading,
      }),
    ],
  });
```

### 5.6 Internal Audit Report Layout (reference-image parity)

Sequence (exactly the reference sample):

1. `createDualHeader(reportTitle, dateRange, { textColor: AUDIT_THEME.color.header, ruleColor: AUDIT_THEME.color.header })`
2. `createCenteredHeading("প্রতিবেদনের সারাংশঃ", { color: AUDIT_THEME.color.summaryHeading })`
3. `createMetadataTable(summaryRows, "navyLabel")`
4. Per section: `createNumberedSectionHeading(bengaliDaariPrefix(i), title, { color: AUDIT_THEME.color.heading, ruleColor: AUDIT_THEME.color.sectionRule, font: TOKENS.font })` → justified body → lists via `createNestedOutline` with `bengaliParenPrefix` (and `subDecimalPrefix(n, m, "bn")` children when the source has 1.1-style sub-points)
5. `createFooter({ locale: "bn" })`

```typescript
// Skeleton (Bengali internal audit report, reference parity)
const summaryRows = [
  {
    label: "প্রতিবেদনের শিরোনাম",
    value: "স্টোরের অভ্যন্তরীণ নিরীক্ষা প্রতিবেদন",
  },
  { label: "নিরীক্ষার সময়কাল", value: "১ নভেম্বর ২০২ হতে ২৭ জুলাই ২০২৬" },
  { label: "নিরীক্ষা কার্যক্রম পরিচালনার তারিখ", value: "১৫ জুলাই ২০২৬ থেকে" },
  { label: "প্রতিবেদনের ধরন", value: "অভ্যন্তরীণ নিরীক্ষা (Internal Audit)" },
];
const scopeItems = [
  {
    prefix: bengaliParenPrefix(1),
    text: "বিভিন্ন উপকরণ-এন্ট্রির রশিদ ও রেজিস্টারে তার নথিভুক্তকরণ তথ্য।",
    children: [
      {
        prefix: subDecimalPrefix(1, 1, "bn"),
        text: "ক্রয় চালানের সাথে রশিদ মিলকরণ।",
      },
    ],
  },
  {
    prefix: bengaliParenPrefix(2),
    text: "উপকরণ বিতরণের রশিদ ও রেজিস্টারে তার নথিভুক্তকরণ তথ্য।",
  },
];
// Document section:
// properties: buildPageSetup("A4_STANDARD"),
// headers: { default: createDualHeader(title, dateRange, { textColor: AUDIT_THEME.color.header, ruleColor: AUDIT_THEME.color.header }) },
// footers: { default: createFooter({ locale: "bn" }) },
// children: [ createCenteredHeading("প্রতিবেদনের সারাংশঃ", { color: AUDIT_THEME.color.summaryHeading }),
//   createMetadataTable(summaryRows, "navyLabel"),
//   ...createNumberedSectionHeading(bengaliDaariPrefix(1), "নিরীক্ষার উদ্দেশ্য",
//        { color: AUDIT_THEME.color.heading, ruleColor: AUDIT_THEME.color.sectionRule }),
//   createBodyParagraph(purposeText),
//   ...createNumberedSectionHeading(bengaliDaariPrefix(2), "নিরীক্ষার পরিধি",
//        { color: AUDIT_THEME.color.heading, ruleColor: AUDIT_THEME.color.sectionRule }),
//   createBodyParagraph(scopeIntro), ...createNestedOutline(scopeItems), … ]
```

Selection trigger: internal audit, vigilance, inspection, compliance review reports — especially with a date-range running header and a summary metadata block. For long committee-audit documents with covers and multi-signatory blocks, use the "Committee / Internal Audit Reports (long form)" matrix row instead.

## 6. Master Compilation Template

```typescript
export function generateDocument(
  pageProfile: keyof typeof PAGE_PROFILES,
  header: Header,
  bodyModules: (Paragraph | Table)[],
  footer?: Footer,
) {
  const doc = new Document({
    sections: [
      {
        properties: buildPageSetup(pageProfile),
        headers: { default: header },
        footers: { default: footer ?? createFooter() },
        children: bodyModules,
      },
    ],
  });
  return Packer.toBuffer(doc); // caller writes buffer to disk / response
}
```

Assembly logic: walk the module sequence from §4a (named) or §4b (derived); between every two table-based modules insert a spacer (`createSpacer()`) so components don't visually fuse. Pass `createDualHeader` for reports with a date range, `createHeader` otherwise; pass `createFooter({ locale: "bn" })` whenever the body content is Bengali.

## 7. Strict AI Execution Rules

1. Page profile is chosen, not assumed (§2 selection rule); Bengali/official documents default to `A4_STANDARD`.
2. All styling routes through `TOKENS` or a §3b theme override — no raw hex/point sizes inside component calls.
3. Modular joining: spacer paragraph between independent table modules.
4. Typography: every `TextRun` gets `font: TOKENS.font` explicitly; `docx` otherwise defaults to Times New Roman.
5. Resolve the font once per §3a before building anything; assign to `TOKENS.font` (and `FORMAL_REGISTER.font`/`AUDIT_THEME.font` if used). Never two Bengali font names in one document. Professional Bangla faces only (Nikosh / SolaimanLipi / Kalpurush / Hind Siliguri / Noto Sans Bengali, fallback Nirmala UI); decorative/display Bangla faces are banned for body text.
6. Title alignment is deliberate (C.2 table): formal/official → CENTER, correspondence → LEFT; pass explicitly.
7. Every formal report gets visible section separation: `withDivider: true`, `createNumberedSectionHeading`, or `createSectionDivider` between every major section — never let body text run into the next heading.
8. Reference-parity audit reports use the §5.6 sequence verbatim: dual header + divider, centered সারাংশ heading, navyLabel metadata table, numbered headings with gold rule, Bengali paren lists, `পৃষ্ঠা X এর Y` footer.
9. Points/sub-points route through C.1/C.3 with the §3a prefix generators; `১.১`/`1.1`-style sub-points use `subDecimalPrefix`. Preserve the source's numbering style and order exactly.
10. Letters/applications MUST set `designPreset: "formal_letter"` and populate `letterMeta` — the server uses a dedicated renderer (`buildLetterDocx`) for this preset. Professional CVs MUST set `designPreset: "professional_cv"` and populate `cvMeta`. Do NOT use report sections for letters or CVs. Biodata uses Module N; legacy resume uses Module O.
11. Callout boxes follow the Module F integrity checklist (full width, one run style, justified, generous margins, never truncated).
12. Headings for TOC: any document using `createTOC()` must apply `HeadingLevel.HEADING_1/2/3` on section headings.
13. Image placement is chosen per Module J's selection rule, never defaulted to centered.
14. Unclassified requests never get refused — run §4b and assemble a bespoke sequence.
15. Valid code only: no mismatched brackets, missing imports, or undefined variables (note: `TabStopType`/`TabStopPosition` are imported for `createDualHeader`).
16. Content completeness: every supplied fact lands in some module; nothing silently dropped.
17. Multilingual safety: preserve non-Latin source verbatim; verify glyph coverage; on delivery state the chosen font and the Nirmala UI fallback for recipients without it.
18. Signatures Block: If a document requires signatures (like Committee Reports, Vigilance Reports, etc.), append a dedicated section at the end with a heading (e.g., "৫. ভিজিল্যান্স কমিটির স্বাক্ষর"). Instead of using paragraphs or standard tables, MUST use the \`signatures\` array. For each signatory, provide their \`role\` (e.g. "সভাপতি:", "সদস্য:") and \`name\` (e.g. "১) প্রভাষক আব্দুল্লাহ আল আমিন:", "২) গার্ড শামসুল আলম :"). Do NOT add the physical signature line (____) in the JSON; the layout engine adds it automatically.

## 8. Delivery Notes

When handing over the file, always report: (a) the resolved font and why (official/classic/modern), (b) the page profile chosen, (c) the archetype matched (or the §4b derivation), (d) for Bengali outputs, that recipients without the chosen font will see Word's substitute with Nirmala UI as the safety net.
