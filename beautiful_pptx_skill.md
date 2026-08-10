---
name: beautiful-pptx-design-master
description: "Ultimate skill for generating visually stunning PowerPoint presentations using pptxgenjs. Trigger on any request to create a .pptx, presentation, slide deck, lecture slides, classroom deck, or teaching slides. This skill contains a complete design system with tested reference code for every layout pattern — title slides, content cards, disease symptoms-and-prevention cards, KPI grids, timelines, comparisons, process flows, Venn diagrams, quiz slides, vocabulary cards, funnels, hierarchies, before/after splits, SWOT grids, stat callouts, and more. Each pattern has multiple variants (light mode, dark mode, gradient backgrounds). Includes curated color palettes, reference fonts for both Bangla and English with auto-detection, emoji/icon maps, background decoration patterns, data overflow slide replication, and an AI decision engine that auto-selects the best layout for any content type. The AI should NEVER create pptxgenjs code from scratch — always use the tested reference functions below."
---

# Beautiful PPTX Design Master — Complete Reference Skill

## Speaker notes compatibility

Do not call `.addNotes()` directly on a value returned by a layout function. Auto-replicating layouts may return an array of slides instead of one `Slide`. In the generator runtime, use `addSpeakerNotes(layoutResult, "your notes")` when notes are required.

This is a **self-contained, all-in-one skill**. Any AI reading this file gets:
1. Tested `pptxgenjs` reference code for every design pattern
2. Multiple visual variants (light, dark, gradient) per pattern
3. Curated color palettes, font stacks with Bangla auto-detection, emoji maps, and decoration patterns
4. Adaptive background system with 6 decoration patterns
5. Data overflow engine — auto-replicate slides when content exceeds capacity
6. An AI decision engine to auto-select layouts based on content analysis
7. A complete boilerplate to generate a final `.pptx` file

**Library**: `pptxgenjs` (Node.js). Never use `python-pptx` for fresh builds — it can't do gradients, has no reliable chart styling, and can't duplicate slides.

---

## §1. Boilerplate — Always Start Here

**CRITICAL COLOR RULE:** You MUST select exactly ONE palette from §2 (e.g., `PALETTES.forestMoss`) and use it for the ENTIRE presentation. Do NOT mix palettes, and do NOT mix Light and Dark modes. Pass the single chosen palette (as `C`) to every layout function.

```javascript
const pptxgen = require('pptxgenjs');
let pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5" — best for projection
pres.author = 'AI Presentation Generator';

// ─── Pick ONE palette (see §2) and set it here ───
// The ENTIRE presentation must use this single palette to maintain consistency!
const C = PALETTES.forestMoss; // or .midnightGold, .coralEnergy, etc.

// ─── Build slides using the pattern functions from §4–§40 ───
// Example: addTitleSlide(pres, "Title", "Subtitle", "Author", C);

// ─── Save ───
pres.writeFile({ fileName: 'output.pptx' })
  .then(() => console.log('✅ Presentation saved!'))
  .catch(err => console.error('❌ Error:', err));
```

---

## §2. Color Palettes

Every palette provides these keys. **All pattern functions below expect this shape. NEVER hardcode colors like 'FFFFFF' or '37474F' inside the functions — always use `C.bg`, `C.textDark`, `C.cardBg`, etc.**

```javascript
const PALETTES = {

  // ━━━ DARK THEMES ━━━

  midnightGold: {
    bg: '1B3A5C',         // Deep navy background
    bgAlt: '162F4A',      // Slightly darker variant
    primary: 'C9992A',    // Gold accent
    accent: 'E8B84B',     // Lighter gold
    gradient1: '1B3A5C',  // Gradient left (header bar)
    gradient2: 'C9992A',  // Gradient right (header bar)
    cardBg: '1E4470',     // Card surface on dark
    cardBgAlt: '234F7F',  // Hover/alt card
    textLight: 'FFFFFF',  // Primary text on dark bg
    textMuted: 'CADCEA',  // Subtitle/muted text on dark bg
    textDark: '1B3A5C',   // Text on light surfaces (rarely used in dark theme)
    textBody: 'E8EDF5',   // General body text for readability
    headerBar: '142D45',  // Top bar color
    border: '2A5A8C',     // Subtle borders
    shadow: '000000',     // Shadow color
    cardText: 'FFFFFF',   // Body text inside cards
  },

  darkCharcoal: {
    bg: '1E1E1E',
    bgAlt: '111111',
    primary: '00C6BB',
    accent: '6FEBA0',
    gradient1: '1E1E1E',
    gradient2: '00C6BB',
    cardBg: '2D2D2D',
    cardBgAlt: '383838',
    textLight: 'FFFFFF',
    textMuted: 'B0B0B0',
    textDark: '1E1E1E',
    textBody: 'E0E0E0',
    headerBar: '1A1A1A',
    border: '444444',
    shadow: '000000',
    cardText: 'F0F0F0',
  },

  cyberNeon: {
    bg: '0B0F19',
    bgAlt: '111827',
    primary: '00E5FF',
    accent: 'FF00FF',
    gradient1: '0B0F19',
    gradient2: '00E5FF',
    cardBg: '161E2E',
    cardBgAlt: '1F2937',
    textLight: 'FFFFFF',
    textMuted: '9CA3AF',
    textDark: '0B0F19',
    textBody: 'F3F4F6',
    headerBar: '111827',
    border: '374151',
    shadow: '000000',
    cardText: 'FFFFFF',
  },

  emeraldAbyss: {
    bg: '062024',
    bgAlt: '082B30',
    primary: '00E676',
    accent: '1DE9B6',
    gradient1: '062024',
    gradient2: '00E676',
    cardBg: '0B343B',
    cardBgAlt: '104751',
    textLight: 'FFFFFF',
    textMuted: '94A3B8',
    textDark: '062024',
    textBody: 'E2E8F0',
    headerBar: '082B30',
    border: '1E5761',
    shadow: '000000',
    cardText: 'FFFFFF',
  },

  crimsonForge: {
    bg: '121212',
    bgAlt: '1A1A1A',
    primary: 'D32F2F',
    accent: 'FF5252',
    gradient1: '121212',
    gradient2: 'D32F2F',
    cardBg: '242424',
    cardBgAlt: '333333',
    textLight: 'FFFFFF',
    textMuted: 'A1A1AA',
    textDark: '121212',
    textBody: 'F4F4F5',
    headerBar: '1A1A1A',
    border: '3F3F46',
    shadow: '000000',
    cardText: 'F4F4F5',
  },

  // ━━━ LIGHT THEMES ━━━

  forestMoss: {
    bg: 'FFFFFF',
    bgAlt: 'F1F8E9',
    primary: '1B5E20',
    accent: '4CAF50',
    gradient1: '1B5E20',
    gradient2: '43A047',
    cardBg: 'FFFFFF',
    cardBgAlt: 'E8F5E9',
    textLight: 'FFFFFF',
    textMuted: '757575',
    textDark: '1B5E20',
    textBody: '37474F',
    headerBar: '1B5E20',
    border: 'E0E0E0',
    shadow: '000000',
    cardText: '37474F',
  },

  oceanBreeze: {
    bg: 'FFFFFF',
    bgAlt: 'E3F2FD',
    primary: '0D47A1',
    accent: '2196F3',
    gradient1: '0D47A1',
    gradient2: '1976D2',
    cardBg: 'FFFFFF',
    cardBgAlt: 'BBDEFB',
    textLight: 'FFFFFF',
    textMuted: '78909C',
    textDark: '0D47A1',
    textBody: '263238',
    headerBar: '0D47A1',
    border: 'E0E0E0',
    shadow: '000000',
    cardText: '263238',
  },

  coralEnergy: {
    bg: 'FFFFFF',
    bgAlt: 'FBE9E7',
    primary: 'BF360C',
    accent: 'FF5722',
    gradient1: 'BF360C',
    gradient2: 'E64A19',
    cardBg: 'FFFFFF',
    cardBgAlt: 'FFCCBC',
    textLight: 'FFFFFF',
    textMuted: '8D6E63',
    textDark: '3E2723',
    textBody: '3E2723',
    headerBar: 'BF360C',
    border: 'E0E0E0',
    shadow: '000000',
    cardText: '3E2723',
  },

  purpleMajesty: {
    bg: 'FFFFFF',
    bgAlt: 'F3E5F5',
    primary: '4A148C',
    accent: '9C27B0',
    gradient1: '4A148C',
    gradient2: '7B1FA2',
    cardBg: 'FFFFFF',
    cardBgAlt: 'E1BEE7',
    textLight: 'FFFFFF',
    textMuted: '7B1FA2',
    textDark: '4A148C',
    textBody: '2E2E2E',
    headerBar: '4A148C',
    border: 'E0E0E0',
    shadow: '000000',
    cardText: '2E2E2E',
  },

  cherryBlossom: {
    bg: 'FFFFFF',
    bgAlt: 'FCE4EC',
    primary: 'C2185B',
    accent: 'E91E63',
    gradient1: 'C2185B',
    gradient2: 'D81B60',
    cardBg: 'FFFFFF',
    cardBgAlt: 'F8BBD0',
    textLight: 'FFFFFF',
    textMuted: 'C2185B',
    textDark: '880E4F',
    textBody: '3E2723',
    headerBar: 'C2185B',
    border: 'E0E0E0',
    shadow: '000000',
    cardText: '3E2723',
  },

  sunsetGlow: {
    bg: 'FFFAFA',
    bgAlt: 'FFF0F0',
    primary: 'FF7043',
    accent: 'FFB74D',
    gradient1: 'FF7043',
    gradient2: 'FF8A65',
    cardBg: 'FFFFFF',
    cardBgAlt: 'FFE0B2',
    textLight: 'FFFFFF',
    textMuted: '8D6E63',
    textDark: 'D84315',
    textBody: '4E342E',
    headerBar: 'FF7043',
    border: 'FFCCBC',
    shadow: '000000',
    cardText: '4E342E',
  },

  arcticBreeze: {
    bg: 'F8FAFC',
    bgAlt: 'F1F5F9',
    primary: '0EA5E9',
    accent: '06B6D4',
    gradient1: '0EA5E9',
    gradient2: '38BDF8',
    cardBg: 'FFFFFF',
    cardBgAlt: 'E0F2FE',
    textLight: 'FFFFFF',
    textMuted: '64748B',
    textDark: '0369A1',
    textBody: '334155',
    headerBar: '0EA5E9',
    border: 'E2E8F0',
    shadow: '000000',
    cardText: '334155',
  },

  lavenderDream: {
    bg: 'FDFBFF',
    bgAlt: 'F4F0F9',
    primary: '7E57C2',
    accent: 'AB47BC',
    gradient1: '7E57C2',
    gradient2: '9575CD',
    cardBg: 'FFFFFF',
    cardBgAlt: 'EDE7F6',
    textLight: 'FFFFFF',
    textMuted: '9575CD',
    textDark: '4527A0',
    textBody: '311B92',
    headerBar: '7E57C2',
    border: 'D1C4E9',
    shadow: '000000',
    cardText: '311B92',
  },

  earthyTerracotta: {
    bg: 'FAF8F5',
    bgAlt: 'F2ECE4',
    primary: '8D6E63',
    accent: '689F38',
    gradient1: '8D6E63',
    gradient2: 'A1887F',
    cardBg: 'FFFFFF',
    cardBgAlt: 'EFEBE9',
    textLight: 'FFFFFF',
    textMuted: '8D6E63',
    textDark: '4E342E',
    textBody: '3E2723',
    headerBar: '8D6E63',
    border: 'D7CCC8',
    shadow: '000000',
    cardText: '3E2723',
  },

  goldenHour: {
    bg: 'FFFCF5',
    bgAlt: 'FFF8E1',
    primary: 'FBC02D',
    accent: 'FFA000',
    gradient1: 'FBC02D',
    gradient2: 'FFCA28',
    cardBg: 'FFFFFF',
    cardBgAlt: 'FFECB3',
    textLight: 'FFFFFF',
    textMuted: 'A1887F',
    textDark: 'F57F17',
    textBody: '4E342E',
    headerBar: 'FBC02D',
    border: 'FFE082',
    shadow: '000000',
    cardText: '4E342E',
  },
};
```

---

## §3. Font Stack & Typography

```javascript
const FONTS = {
  // ─── Bangla ───
  banglaTitle: 'Kalpurush',
  banglaBody: 'Kalpurush',
  banglaFallback: 'Noto Sans Bengali',

  // ─── English ───
  englishTitle: 'Century Gothic',
  englishBody: 'Calibri',
};

/**
 * Auto-detect Bangla text and return the correct font.
 * Use style='title' for headings, 'body' (default) for body text.
 */
function detectFont(text, style) {
  if (!text) return style === 'title' ? FONTS.englishTitle : FONTS.englishBody;
  const hasBangla = /[\u0980-\u09FF]/.test(String(text));
  if (style === 'title') return hasBangla ? FONTS.banglaTitle : FONTS.englishTitle;
  return hasBangla ? FONTS.banglaBody : FONTS.englishBody;
}

/**
 * Helper: detect if background is light (for choosing text color).
 */
function isLightBg(C) {
  const lightBgs = ['FFFFFF', 'FFFBF0', 'FAF3E8', 'F5F5F5', 'FFFAFA', 'F8FAFC', 'FDFBFF', 'FAF8F5', 'FFFCF5'];
  return lightBgs.includes(C.bg);
}
```

---

## §3A. AI Layout Decision Engine — Content-to-Layout Mapping

**CRITICAL INSTRUCTION FOR AI:** Before generating ANY slide, you MUST consult the table below to select the correct layout function. NEVER guess or improvise a layout. If the content doesn't fit any category, use the closest match based on priority order. When multiple layouts can represent the same data, use **Priority 1** unless the user explicitly requests otherwise.

### Routing Protocol (Apply in Order)

1. **Honor explicit intent first.** Use a user-requested chart, timeline, comparison, quiz, SWOT, table, image layout, or dashboard unless the supplied content makes it impossible.
2. **Classify the relationship before counting items.** Route time-based content to a timeline, causal stages to a process, a repeating loop to a cycle, narrowing quantities to a funnel, parent-child content to a hierarchy, and set overlap to a Venn diagram.
3. **Classify the evidence.** Use a chart only for numeric data; use a table when exact values must be read; use a photo/visual layout only when a real, relevant visual is available; otherwise use a text-native layout.
4. **Use the content shape, not a decorative preference.** Cards require independent items. Columns require grouped categories. A two-column layout requires exactly two parallel narratives. Do not choose cards simply because the topic is unfamiliar.
5. **Apply capacity before rendering.** Use the replication limits in §3B. For non-replicating diagrams, simplify the content or use a table/list when it exceeds the safe limits below.
6. **Resolve ties with the first listed priority.** If still tied, choose the lower-density layout that preserves every supplied fact.

### Master Decision Table

| Content Type | Best Layout (Priority 1) | Alternative (Priority 2) | Alternative (Priority 3) | When to Use |
|---|---|---|---|---|
| **Opening / Cover** | `addTitleSlide` §6 | — | — | ALWAYS the first slide. Contains presentation title, subtitle, author. |
| **Chapter / Section Break** | `addSectionDivider` §11 | — | — | ALWAYS insert between major topic changes. Use incrementing numbers. |
| **Closing / End** | `addThankYouSlide` §12 | — | — | ALWAYS the last slide. |
| **One focused idea with 3–5 traits** | `addContentWithIconRow` §27 | `addStatCallout` §8 | `addCardGrid3` §7 | Use for a definition, principle, product benefit, or concept that needs one explanation plus supporting traits. |
| **Exactly two parallel ideas** | `addTwoColumnText` §40 | `addBeforeAfterSlide` §37 | `addComparisonMatrix` §10 | Use narrative columns for two balanced explanations; use Before/After only when change over time is the point. |
| **3 key concepts / features** | `addCardGrid3` §7 | `addIconListWithLines` §23 | `addCategoryColumns` §14A | Use cards when each item has a title AND body text. Use icon list for title-only items. |
| **4–6 concepts / features** | `addCardGrid6` §7 | `addCategoryColumns` §14A | `addIconListWithLines` §23 | Use 6-card grid for items with title+body. Use columns for grouped bullet lists. |
| **7–9 concepts / features** | `addCardGrid9` §7A | `addCardGrid6` §7 | `addCategoryColumns` §14A | Use 9-card grid for 7-9 items with title+body. |
| **Single big number / statistic** | `addStatCallout` §8 | — | — | Use when ONE metric needs dramatic visual impact (e.g., "95%", "$2.4M"). |
| **Executive summary: 3 KPIs + intro + objectives + quote** | `addExecutiveSummarySlide` §41 | `addKpiGrid` §9 + `addQuoteSlide` §24 | `addSummaryOverview` §28 | Use this composite layout whenever all four elements are requested. Do **not** substitute a chart or dashboard. |
| **Multiple KPIs / metrics** | `addKpiGrid` §9 | `addCardGrid3` §7 | `addDataTable` §29 | Use KPI grid for 2–4 numeric values. Use data table if there are labels+values in rows. |
| **Feature comparison (2–4 items)** | `addComparisonMatrix` §10 | `addDataTable` §29 | `addCategoryListColumns` §14B | Use a matrix when the same named criteria are evaluated across each item. Use a table for raw values and a category list for independent descriptions. |
| **Tabular / structured data** | `addDataTable` §29 | `addComparisonMatrix` §10 | — | Use when data has clear headers and rows (specs, pricing, schedules). |
| **Sequential process (3+ steps)** | `addProcessFlow` §13 | `addTimeline` §16 | `addCycleDiagram` §17 | Use one item per named stage. Five stages fit per slide; the layout automatically continues with the remaining stages. Use timeline for date-based milestones. |
| **Date-based milestones** | `addTimeline` §16 | `addProcessFlow` §13 | `addDataTable` §29 | Use for 3–6 dated or quarter-based milestones; use a process when dates are secondary, and a table when there are more milestones or detailed dates. |
| **Repeating / circular process** | `addCycleDiagram` §17 | `addProcessFlow` §13 | `addSummaryOverview` §28 | Use for a closed loop of 3–6 recurring stages (CI/CD, PDCA, feedback loops). Use a flow if it has a true start and end. |
| **Narrowing stages (sales, conversion)** | `addFunnel` §18 | `addProcessFlow` §13 | `addBarChartSlide` §15 | Use for 3–5 ordered stages whose quantity or scope narrows. Use a flow when no narrowing relationship exists. |
| **Overlap / intersection of 2 concepts** | `addVennDiagram` §19 | `addTwoColumnText` §40 | — | Use only for two sets plus their shared middle; do not force unrelated comparisons into a Venn diagram. |
| **Hierarchy / org chart** | `addHierarchyTree` §20 | `addCategoryColumns` §14A | — | Use for parent-child structures (org charts, classification trees). |
| **Grouped categories (bulleted)** | `addCategoryColumns` §14A | `addCardGrid3` §7 | `addIconListWithLines` §23 | Use when 2–3 categories each have multiple bullet items. |
| **Grouped categories (card rows)** | `addCategoryListColumns` §14B | `addDataTable` §29 | — | Use when categories have items with title+description pairs. |
| **Numerical chart / bar graph** | `addBarChartSlide` §15 | `addKpiGrid` §9 | `addDataTable` §29 | Use for 3–10 discrete category values. Use KPI cards for 2–4 headline totals and a table when exact values matter more than the visual pattern. |
| **Tabular data with comparative chart** | `addTableAndChartSlide` §34 | `addDataTable` §29 | `addBarChartSlide` §15 | Use when comparing multi-metric data with both raw numbers and visual bars. |
| **Tabular data with textual summary** | `addTableWithSummarySlide` §36 | `addDataTable` §29 | `addTableAndChartSlide` §34 | Use when a table requires a detailed explanatory text box or insights underneath it. |
| **Percentage breakdown / pie chart** | `addPieChartSlide` §32 | `addBarChartSlide` §15 | `addKpiGrid` §9 | Use only for 2–6 mutually exclusive categories that sum to 100%. Use bars for more categories, values that do not total 100%, or close comparisons. |
| **Trends over time / continuous data** | `addLineChartSlide` §33 | `addBarChartSlide` §15 | `addDataTable` §29 | Use for 4+ ordered time points where direction/rate of change matters. Use bars for discrete periods and a table when values must be read exactly. |
| **Definitions / vocabulary** | `addVocabGrid` §21 | `addCategoryListColumns` §14B | `addDataTable` §29 | Use for term+definition pairs (glossary, key terms). |
| **Quiz / multiple choice question** | `addQuizSlide` §22 | — | — | Use for assessment, review, or interactive engagement slides. |
| **Ordered list of key points** | `addIconListWithLines` §23 | `addSummaryOverview` §28 | `addCardGrid3` §7 | Use icon list for principles/values. Use summary for chapter recaps. |
| **Inspirational quote / key message** | `addQuoteSlide` §24 | `addStatCallout` §8 | — | Use for impactful quotes, concluding thoughts, or mission statements. |
| **Overview with visual + list** | `addSplitInfoImage` §30 | `addLeftPanelIconList` §25 | `addIconListWithLines` §23 | Use the split layout when a real relevant image is available. Use the left panel when an illustrative icon/emoji is sufficient; never use an empty image placeholder as decoration. |
| **Warnings / consequences** | `addDarkIconCardRow` §26 | `addCardGrid3` §7 | `addIconListWithLines` §23 | Use dark card row for negative outcomes, risks, or dangers. |
| **Definition + characteristics** | `addContentWithIconRow` §27 | `addCardGrid3` §7 | `addVocabGrid` §21 | Use when explaining "What is X?" with key traits displayed as icons. |
| **Chapter summary / recap** | `addSummaryOverview` §28 | `addIconListWithLines` §23 | `addCardGrid3` §7 | Use numbered bars for end-of-chapter summaries and key takeaways. |
| **Disease / health topic: image + signs & symptoms + prevention** | `addDiseaseSymptomsPreventionCards` §43 | `addSplitInfoImage` §30 | `addCardGrid3` §7 | ALWAYS use this as the first choice when each condition needs a photo, a named title, symptom/sign bullets, and prevention bullets. Pass every condition in one `diseases` array; never combine signs and prevention or replace the image with an emoji. |
| **Product/topic with image** | `addSplitInfoImage` §30 | `addLeftPanelIconList` §25 | — | Use when content has bullet points AND a relevant image or diagram. |
| **Map or diagram with right-side details** | `addVisualPlaceholderWithCards` §35 | `addSplitInfoImage` §30 | `addLeftPanelIconList` §25 | Use when a complex visual (like a map or diagram) needs explanatory cards on the right. |
| **Photo gallery / portfolio** | `addImageGallery` §31 | — | — | Use when showcasing multiple images, screenshots, or visual artifacts. |
| **Before/After comparison** | `addBeforeAfterSlide` §37 | `addComparisonMatrix` §10 | `addTwoColumnText` §40 | Use when showing a transformation, improvement, or change between two states. |
| **Tutorial / how-to steps** | `addNumberedStepCards` §38 | `addProcessFlow` §13 | `addIconListWithLines` §23 | Use for tutorial steps where each step needs a title + description card. |
| **SWOT / 2×2 analysis** | `addSwotGrid` §39 | `addCardGrid6` §7 | `addCategoryColumns` §14A | Use for SWOT, 2×2 matrix, or any four-quadrant analysis. |
| **Two-column narrative text** | `addTwoColumnText` §40 | `addSplitInfoImage` §30 | `addCategoryColumns` §14A | Use when comparing or presenting two parallel blocks of written text. |
| **Complex financial / multi-metric dashboard** | `addDashboardSlide` §42 | `addDataTable` §29 | `addKpiGrid` §9 | Use when presenting multiple diverse metrics, gauges, or KPIs together. |

### Selection Guardrails

- **Images:** Pass a real `imagePath` to `addSplitInfoImage` and real image paths to `addImageGallery`. If no relevant image exists, select a text, card, chart, or table layout instead.
- **Charts:** Never invent numeric values to justify a chart. Do not use a pie chart unless categories are mutually exclusive and total 100%; do not use a line chart for unordered categories.
- **Comparisons:** Use `addBeforeAfterSlide` for a state change, `addComparisonMatrix` for repeated criteria, and `addTwoColumnText` for prose-only parallel perspectives. These layouts are not interchangeable.
- **Diagrams:** Keep `addTimeline`, `addCycleDiagram`, and `addFunnel` within their safe item counts. For denser inputs, preserve all data with `addDataTable`, `addSummaryOverview`, or continuation process slides rather than shrinking labels.
- **Cards:** Use `addCardGrid3`, `addCardGrid6`, or `addCardGrid9` only when each item has comparable weight and an independently meaningful title. Use category columns when items belong under shared headings.
- **Composite requests:** If a prompt contains multiple content types, allocate one layout per coherent message rather than forcing all data into a dashboard. The exception is `addExecutiveSummarySlide`, which is mandatory when its KPI + introduction + objectives + quote input shape is present.
- **Fallback:** If no row matches, identify the slide's dominant communication job—explain, compare, sequence, quantify, group, recap, or prompt action—and select the first matching row. Preserve all facts and prefer the least dense valid layout.

### Presentation Structure Rules

**COMPOSITE-CONTENT PRIORITY:** When a prompt requests KPIs together with an introduction, objectives, and a quote, it is an executive summary—not a dashboard. Use `addExecutiveSummarySlide` before considering `addDashboardSlide`. Only use `addDashboardSlide` when the user explicitly requests a dashboard or provides chart/graph data.

**MANDATORY SLIDE ORDER:**
1. **First slide** → ALWAYS `addTitleSlide` §6
2. **Before each new topic** → ALWAYS `addSectionDivider` §11
3. **Content slides** → Use the Decision Table above
4. **Last slide** → ALWAYS `addThankYouSlide` §12

**SLIDE COUNT GUIDELINES:**
- Short presentation (5–10 slides): Title + 1 divider + 3–6 content + Thank You
- Medium presentation (10–20 slides): Title + 2–3 dividers + 6–15 content + Thank You
- Long presentation (20–30 slides): Title + 4–6 dividers + 14–24 content + Thank You

**PALETTE SELECTION GUIDE:**
- Agriculture / Nature / Environment → `forestMoss` (green)
- Technology / Corporate / Finance → `oceanBreeze` (blue)
- Energy / Food / Health / Sports → `coralEnergy` (red/orange)
- Education / Culture / Arts → `purpleMajesty` (violet)
- Social / Charity / Wellness → `cherryBlossom` (pink)
- Optimism / Warmth / Friendly → `sunsetGlow` (orange/pink)
- Clean / Medical / Modern → `arcticBreeze` (ice/cyan)
- Soft / Creative / Dreamy → `lavenderDream` (lilac/purple)
- Natural / Grounded / Trust → `earthyTerracotta` (sand/terracotta)
- Success / Wealth / Energy → `goldenHour` (ivory/amber)
- Premium / Luxury / Awards → `midnightGold` (dark gold)
- Coding / Tech / Dark Mode → `darkCharcoal` (dark teal)
- Future / Cyber / High-Energy → `cyberNeon` (black/neon)
- Deep Calm / Elegant / Security → `emeraldAbyss` (dark teal/green)
- Power / Urgency / Intense → `crimsonForge` (black/red)

**STRICT DATA FIDELITY & OVERFLOW RULE (CRITICAL FOR AI):**
You MUST NOT invent, drop, skip, merge, or ignore the user's data points. If the user provides 15 items, extract and pass exactly 15 items to the layout function in the original order. You may summarize long descriptions, fix grammar, and polish the text, but preserve every item and factual point.
For a named sequence, each stage is atomic: one named stage equals one array object. Never merge consecutive stages to make a layout fit. For example, a seven-stage supply chain must be an array of seven process-flow objects.
You MUST pass ALL available data as a single large array to the layout function, even if it exceeds the slide's capacity (for example, pass all seven supply-chain stages to `addProcessFlow`). DO NOT truncate or manually split the data across multiple calls. The layout's internal `chunkArray` logic creates continuation slides automatically.

**BANGLA FONT RULE:**
Before placing ANY text on a slide, call `detectFont(text)` to auto-select the correct font. NEVER hardcode `'Calibri'` for user-provided text — it may break Bangla rendering. Use `detectFont(text, 'title')` for titles and `detectFont(text)` for body text. System labels (like footer, option letters) may use hardcoded English fonts.

**ADAPTIVE BACKGROUND RULE:**
- Title slides → ALWAYS use `addDecoCircles` + `addSubtleWave`
- Section dividers → Use `addDiagonalStripe` for dark palettes, `addCornerAccents` for light
- Quote slides → Use `addDiagonalStripe`
- Content slides with cards → Optionally use `addCornerAccents` (subtle)
- Summary / Thank You slides → Use `addSubtleWave`
- Stat Callout slides → Light bg: `addGradientBar`, Dark bg: `addDecoCircles`

---

## §3B. Data Overflow & Slide Replication Engine

When data exceeds what fits on one slide, the layout must automatically create additional continuation slides. This helper and the per-layout max-items table enable that behavior.

```javascript
/**
 * Split an array into chunks of maxPerSlide.
 * Used internally by layout functions to auto-replicate slides.
 */
function chunkArray(arr, maxPerSlide) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += maxPerSlide) {
    chunks.push(arr.slice(i, i + maxPerSlide));
  }
  return chunks;
}
```

### Max Items Per Layout (Replication Reference)

| Layout | Max Items/Slide | Overflow Behavior |
|---|---|---|
| `addCardGrid3` | 3 | Replicate for cards 4-6, 7-9, etc. |
| `addCardGrid6` | 6 | Replicate for items 7-12, etc. |
| `addCardGrid9` | 9 | Replicate for items 10-18, etc. |
| `addIconListWithLines` | 5 | Replicate with "(cont.)" |
| `addVocabGrid` | 6 | Replicate for terms 7-12, etc. |
| `addSummaryOverview` | 8 | Replicate with continued numbering |
| `addDataTable` | 8 rows | Same headers, next batch of rows |
| `addDarkIconCardRow` | 4 cards | Replicate for cards 5-8, etc. |
| `addImageGallery` | 6 images | Replicate for images 7-12, etc. |
| `addProcessFlow` | 5 steps | Preserve one card per stage; create a continuation flow slide for step 6 and later |
| `addTableWithSummarySlide` | 6 rows | Replicate table; summary only on last |
| `addTableAndChartSlide` | 8 data items | Replicate table+chart together |
| `addDiseaseSymptomsPreventionCards` | 3 disease cards | Replicate in three-card poster rows; retain each condition's image, signs/symptoms, and prevention list together. |
| `addTimeline` | 6 milestones | Does not auto-replicate; use a process for undated steps or a table for more/detailed milestones. |
| `addCycleDiagram` | 6 stages | Does not auto-replicate; use a process or summary list if the loop has more stages. |
| `addFunnel` | 5 stages | Does not auto-replicate; use a process or bar chart when there are more stages. |

---

## §4. Background Decoration Patterns

These decorative shapes add visual depth. Apply them to title slides, section dividers, and impact slides.

```javascript
// ─── Existing Patterns (Enhanced) ───

function addDecoCircles(slide, C) {
  slide.addShape('ellipse', { x: 8.5, y: -1.5, w: 6, h: 6, fill: { color: C.primary, transparency: 92 }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: -1.5, y: 5, w: 4, h: 4, fill: { color: C.primary, transparency: 92 }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: 5.5, y: 0.3, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 88 }, line: { type: 'none' } });
  // Extra subtle dot cluster for depth
  slide.addShape('ellipse', { x: 11.0, y: 5.8, w: 0.6, h: 0.6, fill: { color: C.accent, transparency: 85 }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: 12.0, y: 5.2, w: 0.4, h: 0.4, fill: { color: C.primary, transparency: 88 }, line: { type: 'none' } });
}

function addGradientBar(slide, C) {
  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 0.12, fill: { color: C.accent }, line: { type: 'none' } });
  slide.addShape('rect', { x: 0, y: 7.38, w: 13.33, h: 0.12, fill: { color: C.primary }, line: { type: 'none' } });
}

// ─── New Adaptive Patterns ───

function addCornerAccents(slide, C) {
  // Top-left triangle accent
  slide.addShape('rtTriangle', { x: -0.1, y: -0.1, w: 1.8, h: 1.8, fill: { color: C.primary, transparency: 94 }, line: { type: 'none' }, rotate: 0 });
  // Bottom-right triangle accent
  slide.addShape('rtTriangle', { x: 11.63, y: 5.8, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 94 }, line: { type: 'none' }, rotate: 180 });
  // Small accent dot
  slide.addShape('ellipse', { x: 12.2, y: 0.4, w: 0.5, h: 0.5, fill: { color: C.accent, transparency: 88 }, line: { type: 'none' } });
}

function addDiagonalStripe(slide, C) {
  // Diagonal accent stripe across slide for visual energy
  slide.addShape('rect', { x: -2, y: 6.2, w: 17, h: 0.15, fill: { color: C.accent, transparency: 75 }, line: { type: 'none' }, rotate: -5 });
  slide.addShape('rect', { x: -2, y: 6.5, w: 17, h: 0.06, fill: { color: C.primary, transparency: 85 }, line: { type: 'none' }, rotate: -5 });
}

function addDotMatrix(slide, C) {
  // Scattered dots for subtle texture (top-right and bottom-left)
  const dotColor = C.primary;
  const positions = [
    { x: 10.5, y: 0.8 }, { x: 11.2, y: 1.1 }, { x: 11.8, y: 0.6 }, { x: 12.3, y: 1.3 },
    { x: 10.8, y: 1.6 }, { x: 11.5, y: 1.9 }, { x: 12.0, y: 0.3 },
    { x: 0.8, y: 6.0 }, { x: 1.4, y: 6.4 }, { x: 0.5, y: 6.8 }, { x: 1.8, y: 6.6 }
  ];
  positions.forEach(p => {
    slide.addShape('ellipse', { x: p.x, y: p.y, w: 0.12, h: 0.12, fill: { color: dotColor, transparency: 88 }, line: { type: 'none' } });
  });
}

function addSubtleWave(slide, C) {
  // Smooth wave shape at bottom for elegance
  slide.addShape('ellipse', { x: -3, y: 6.0, w: 20, h: 3.0, fill: { color: C.primary, transparency: 95 }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: -2, y: 6.4, w: 18, h: 2.5, fill: { color: C.accent, transparency: 96 }, line: { type: 'none' } });
}
```

---

## §5. Full-Width Header Bar

```javascript
function addHeaderBar(slide, titleText, C, opts = {}) {
  const barH = opts.barH || 0.95;
  const fontSize = opts.fontSize || 30;
  const emoji = opts.emoji || '';
  const displayText = emoji ? `${emoji}  ${titleText}` : titleText;
  const fontFace = detectFont(displayText, 'title');

  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: barH, fill: { color: C.headerBar }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.15, blur: 8, offset: 3, angle: 270 } });
  slide.addText(displayText, { x: 0.5, y: 0, w: 12.33, h: barH, fontSize: fontSize, bold: true, color: C.textLight, fontFace: fontFace, valign: 'middle' });
  // Thin accent underline for premium feel
  slide.addShape('rect', { x: 0, y: barH, w: 13.33, h: 0.04, fill: { color: C.accent }, line: { type: 'none' } });
}

function addFooter(slide, C, text) {
  if (!text) return; // No branding by default
  slide.addText(text, {
    x: 0.5, y: 7.0, w: 12.33, h: 0.3,
    align: 'right', valign: 'middle',
    fontSize: 10, color: C.textMuted || '888888', fontFace: FONTS.englishBody
  });
}
```

---

## §6. Title Slide Pattern (Adapts to palette)

```javascript
function addTitleSlide(pres, title, subtitle, author, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  slide.transition = { type: 'fade', speed: 'med' };

  addDecoCircles(slide, C);
  addSubtleWave(slide, C);
  
  const mainColor = isLightBg(C) ? C.primary : C.textLight;
  const authorColor = isLightBg(C) ? C.textDark : C.textLight;

  slide.addText(title, { x: 0.8, y: 2.0, w: 11.7, h: 1.5, align: 'center', fontSize: 48, bold: true, color: mainColor, fontFace: detectFont(title, 'title'), charSpacing: 1.5 });
  if (subtitle) slide.addText(subtitle, { x: 1.2, y: 3.5, w: 10.9, h: 0.8, align: 'center', fontSize: 20, italic: true, color: isLightBg(C) ? C.textDark : C.textLight, fontFace: detectFont(subtitle) });
  slide.addShape('rect', { x: 5.0, y: 4.5, w: 3.33, h: 0.06, fill: { color: C.accent }, line: { type: 'none' } });
  if (author) slide.addText(author, { x: 1, y: 5.5, w: 11.33, h: 1.2, align: 'center', fontSize: 24, color: authorColor, fontFace: detectFont(author) });
  
  return slide;
}
```

---

## §7. Card Grid Layout (Adapts to palette, auto-replicates)

```javascript
function addCardGrid3(pres, title, cards, C) {
  const MAX = 3;
  const chunks = chunkArray(cards, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);
    if (ci === 0) addCornerAccents(slide, C);

    const n = Math.min(cards.length, MAX);
    const gap = 0.35, cardW = (12.33 - gap * (n - 1)) / n, cardH = 2.8, yStart = 1.4, headerH = 0.6;
    const headerColors = [C.primary, C.accent, C.primary];

    chunk.forEach((card, i) => {
      const x = 0.5 + i * (cardW + gap);
      
      slide.addShape('roundRect', {
        x, y: yStart, w: cardW, h: cardH, rectRadius: 0.1,
        fill: { color: C.cardBg },
        line: { color: C.border, width: 1.5 },
        shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 8, offset: 2, angle: 225 }
      });

      slide.addShape('rect', {
        x, y: yStart, w: cardW, h: headerH,
        fill: { color: headerColors[i % headerColors.length] },
        line: { type: 'none' }, rectRadius: 0.1
      });

      const headerText = card.emoji ? `${card.emoji}  ${card.title}` : card.title;
      slide.addText(headerText, {
        x, y: yStart, w: cardW, h: headerH, align: 'center', valign: 'middle',
        fontSize: 16, bold: true, color: C.textLight, fontFace: detectFont(headerText, 'title')
      });

      slide.addText(card.body, {
        x: x + 0.15, y: yStart + headerH + 0.1, w: cardW - 0.3, h: cardH - headerH - 0.2,
        fontSize: 14, color: C.cardText, fontFace: detectFont(card.body), valign: 'top', paraSpaceAfter: 6
      });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

```javascript
function addCardGrid6(pres, title, cards, C) {
  const MAX = 6;
  const chunks = chunkArray(cards, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const cols = 3, gap = 0.3, cardW = (12.33 - gap * (cols - 1)) / cols, cardH = 2.5, headerH = 0.55;
    const headerColors = [C.primary, C.accent, C.primary, C.accent, C.primary, C.accent];

    chunk.forEach((card, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const x = 0.5 + col * (cardW + gap), y = 1.3 + row * (cardH + gap);

      slide.addShape('roundRect', {
        x, y, w: cardW, h: cardH, rectRadius: 0.14,
        fill: { color: C.cardBg }, line: { color: C.border, width: 1 },
        shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 6, offset: 2, angle: 225 }
      });

      // Keep the colored header inside the rounded card instead of masking its corners.
      slide.addShape('roundRect', {
        x: x + 0.08, y: y + 0.08, w: cardW - 0.16, h: headerH - 0.08,
        rectRadius: 0.08, fill: { color: headerColors[i] }, line: { type: 'none' }
      });

      const headerText = card.emoji ? `${card.emoji}  ${card.title}` : card.title;
      slide.addText(headerText, {
        x: x + 0.12, y: y + 0.04, w: cardW - 0.24, h: headerH - 0.08, align: 'center', valign: 'middle',
        fontSize: 14, bold: true, color: C.textLight, fontFace: detectFont(headerText, 'title')
      });

      slide.addText(card.body, {
        x: x + 0.12, y: y + headerH + 0.08, w: cardW - 0.24, h: cardH - headerH - 0.15,
        fontSize: 13, color: C.cardText, fontFace: detectFont(card.body), valign: 'top'
      });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §7A. 9-Card Compact Grid (Adapts to palette, auto-replicates)

```javascript
function addCardGrid9(pres, title, cards, C) {
  const MAX = 9;
  const chunks = chunkArray(cards, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const n = Math.min(cards.length, MAX);
    const cols = 3, gap = 0.2, cardW = (12.33 - gap * (cols - 1)) / cols, cardH = 1.8, headerH = 0.45;
    const headerColors = [C.primary, C.accent, C.primary, C.accent, C.primary, C.accent, C.primary, C.accent, C.primary];

    chunk.forEach((card, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const x = 0.5 + col * (cardW + gap), y = 1.3 + row * (cardH + gap);

      slide.addShape('roundRect', {
        x, y, w: cardW, h: cardH, rectRadius: 0.1,
        fill: { color: C.cardBg }, line: { color: C.border, width: 1 },
        shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 4, offset: 2, angle: 225 }
      });

      // Keep every visible box rounded: an inset header preserves the card outline.
      slide.addShape('roundRect', {
        x: x + 0.08, y: y + 0.08, w: cardW - 0.16, h: headerH - 0.08,
        rectRadius: 0.06, fill: { color: headerColors[i] }, line: { type: 'none' }
      });

      const numText = String(ci * MAX + i + 1);
      
      slide.addShape('roundRect', {
        x: x + 0.15, y: y + 0.1, w: 0.25, h: 0.25,
        rectRadius: 0.04, fill: { color: C.bg }, line: { type: 'none' }
      });
      slide.addText(numText, {
        x: x + 0.15, y: y + 0.1, w: 0.25, h: 0.25, align: 'center', valign: 'middle',
        fontSize: 10, bold: true, color: headerColors[i], fontFace: FONTS.englishTitle
      });

      const headerText = card.emoji ? `${card.emoji}  ${card.title}` : card.title;
      slide.addText(headerText, {
        x: x + 0.5, y: y + 0.04, w: cardW - 0.6, h: headerH - 0.08, align: 'left', valign: 'middle',
        fontSize: 12, bold: true, color: C.textLight, fontFace: detectFont(headerText, 'title')
      });

      slide.addText(card.body, {
        x: x + 0.12, y: y + headerH + 0.08, w: cardW - 0.24, h: cardH - headerH - 0.15,
        fontSize: 11, color: C.cardText, fontFace: detectFont(card.body), valign: 'top'
      });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §8. Stat Callout (Adapts to palette)

```javascript
function addStatCallout(pres, value, label, context, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  
  if (isLightBg(C)) {
    addGradientBar(slide, C);
    // Halo ring behind the number
    slide.addShape('ellipse', { x: 4.16, y: 1.5, w: 5.0, h: 3.5, fill: { color: C.accent, transparency: 93 }, line: { type: 'none' } });
    slide.addText(value, { x: 0, y: 2.0, w: 13.33, h: 2.2, align: 'center', fontSize: 96, bold: true, color: C.primary, fontFace: detectFont(value, 'title') });
    slide.addText(label, { x: 1, y: 4.2, w: 11.33, h: 0.7, align: 'center', fontSize: 24, color: C.textDark, fontFace: detectFont(label, 'title') });
  } else {
    addDecoCircles(slide, C);
    slide.addShape('ellipse', { x: 4.16, y: 1.5, w: 5.0, h: 3.5, fill: { color: C.accent, transparency: 90 }, line: { type: 'none' } });
    slide.addText(value, { x: 0, y: 2.0, w: 13.33, h: 2.2, align: 'center', fontSize: 96, bold: true, color: C.accent, fontFace: detectFont(value, 'title') });
    slide.addText(label, { x: 1, y: 4.2, w: 11.33, h: 0.7, align: 'center', fontSize: 24, color: C.textLight, fontFace: detectFont(label, 'title') });
  }
  
  if (context) {
    slide.addText(context, { x: 1.5, y: 5.0, w: 10.33, h: 0.6, align: 'center', fontSize: 14, italic: true, color: isLightBg(C) ? C.textDark : C.textLight, fontFace: detectFont(context) });
  }
  return slide;
}
```

---

## §9. KPI Dashboard Grid

```javascript
function addKpiGrid(pres, title, kpis, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const n = kpis.length, gap = 0.35, cardW = (12.33 - gap * (n - 1)) / n, y = 1.8, cardH = 4.2;

  kpis.forEach((kpi, i) => {
    const x = 0.5 + i * (cardW + gap);
    slide.addShape('roundRect', {
      x, y, w: cardW, h: cardH, rectRadius: 0.08,
      fill: { color: C.cardBg }, line: { type: 'none' },
      shadow: { type: 'outer', color: C.shadow, opacity: 0.12, blur: 8, offset: 2, angle: 90 }
    });

    // Colored top accent strip
    slide.addShape('rect', { x, y, w: cardW, h: 0.08, fill: { color: i % 2 === 0 ? C.primary : C.accent }, line: { type: 'none' } });

    // Trend arrow support: if kpi.trend is 'up' show ▲, 'down' show ▼
    const trendIcon = kpi.trend === 'up' ? ' ▲' : kpi.trend === 'down' ? ' ▼' : '';
    const displayValue = kpi.value + trendIcon;

    slide.addText(displayValue, {
      x, y: y + 0.4, w: cardW, h: 1.4, align: 'center',
      fontSize: 44, bold: true, color: C.accent, fontFace: detectFont(kpi.value, 'title')
    });

    slide.addText(kpi.label, {
      x: x + 0.15, y: y + 2.0, w: cardW - 0.3, h: 1.8, align: 'center',
      fontSize: 15, color: C.cardText, fontFace: detectFont(kpi.label), valign: 'top'
    });
  });

  return slide;
}
```

---

## §10. Comparison Matrix

```javascript
function addComparisonMatrix(pres, title, items, criteria, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const labelColW = 3.0, itemColW = (12.33 - labelColW) / items.length, rowH = Math.min(0.85, 5.5 / (criteria.length + 1));
  let y = 1.5;

  slide.addShape('rect', { x: 0.5, y, w: labelColW, h: rowH, fill: { color: C.bgAlt }, line: { type: 'none' } });
  items.forEach((item, i) => {
    slide.addShape('rect', { x: 0.5 + labelColW + i * itemColW, y, w: itemColW, h: rowH, fill: { color: C.primary }, line: { type: 'none' } });
    slide.addText(item, { x: 0.5 + labelColW + i * itemColW, y, w: itemColW, h: rowH, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.textLight, fontFace: detectFont(item, 'title') });
  });
  y += rowH;

  criteria.forEach((row, r) => {
    const shade = r % 2 === 0 ? C.cardBg : C.bgAlt;
    slide.addShape('rect', { x: 0.5, y, w: labelColW, h: rowH, fill: { color: shade }, line: { color: C.border, width: 0.5 } });
    slide.addText(row.label, { x: 0.65, y, w: labelColW - 0.3, h: rowH, valign: 'middle', fontSize: 13, bold: true, color: isLightBg(C) ? C.textDark : C.cardText, fontFace: detectFont(row.label) });
    
    row.values.forEach((val, i) => {
      slide.addShape('rect', { x: 0.5 + labelColW + i * itemColW, y, w: itemColW, h: rowH, fill: { color: shade }, line: { color: C.border, width: 0.5 } });
      slide.addText(val, { x: 0.5 + labelColW + i * itemColW + 0.1, y, w: itemColW - 0.2, h: rowH, valign: 'middle', fontSize: 12, color: C.cardText, fontFace: detectFont(val) });
    });
    y += rowH;
  });

  return slide;
}
```

---

## §11. Section Divider

```javascript
function addSectionDivider(pres, sectionTitle, sectionNumber, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  slide.transition = { type: 'fade', speed: 'med' };
  
  // Adaptive background
  if (isLightBg(C)) { addCornerAccents(slide, C); } else { addDiagonalStripe(slide, C); }

  // Two-tone professional split design with shadow edge
  slide.addShape('rect', { x: 0, y: 0, w: 4.0, h: 7.5, fill: { color: C.primary }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.2, blur: 12, offset: 4, angle: 0 } });
  slide.addShape('rect', { x: 3.92, y: 0, w: 0.08, h: 7.5, fill: { color: C.accent }, line: { type: 'none' } });
  
  if (sectionNumber) {
    slide.addText(String(sectionNumber).padStart(2, '0'), { x: 0, y: 2.5, w: 4.0, h: 2.0, align: 'center', fontSize: 96, bold: true, color: C.accent, fontFace: FONTS.englishTitle });
  }

  const textColor = isLightBg(C) ? C.textDark : C.textLight;
  slide.addText(sectionTitle, { x: 4.8, y: 2.8, w: 8.0, h: 1.5, align: 'left', fontSize: 40, bold: true, color: textColor, fontFace: detectFont(sectionTitle, 'title') });
  slide.addShape('rect', { x: 4.8, y: 4.5, w: 3.0, h: 0.08, fill: { color: C.accent }, line: { type: 'none' } });
  
  return slide;
}
```

---

## §12. Thank You / Impact Slide

```javascript
function addThankYouSlide(pres, mainText, subText, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  slide.transition = { type: 'fade', speed: 'med' };
  addDecoCircles(slide, C);
  addSubtleWave(slide, C);

  const mainColor = isLightBg(C) ? C.primary : C.textLight;
  const defaultText = 'ধন্যবাদ | Thank You';
  const displayText = mainText || defaultText;

  // Reserve enough vertical space for wrapped (including three-line) closing titles.
  // The divider must always stay outside the title text box.
  slide.addText(displayText, { x: 1, y: 1.9, w: 11.33, h: 2.85, align: 'center', valign: 'middle', fontSize: 48, bold: true, color: mainColor, fontFace: detectFont(displayText, 'title'), fit: 'shrink' });
  slide.addShape('rect', { x: 5.0, y: 5.0, w: 3.33, h: 0.06, fill: { color: C.accent }, line: { type: 'none' } });
  if (subText) slide.addText(subText, { x: 1, y: 5.3, w: 11.33, h: 1.2, align: 'center', valign: 'top', fontSize: 18, color: C.textMuted, fontFace: detectFont(subText), fit: 'shrink' });

  return slide;
}
```

---

## §13. Process Flow

```javascript
function addProcessFlow(pres, title, steps, C) {
  const MAX = 5;
  const chunks = chunkArray(steps, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    // Use simple cards and text chevrons only. They are robust native PPTX
    // elements and avoid SmartArt/connector compatibility problems.
    // Lock the grid to the first slide's capacity. Continuation slides may
    // contain fewer steps, but their cards must retain the same dimensions
    // and spacing rather than expanding to fill the available width.
    const slotsPerSlide = Math.min(MAX, steps.length);
    const n = chunk.length, connectorW = 0.28, gap = 0.16;
    const boxW = (12.33 - (connectorW + gap) * (slotsPerSlide - 1)) / slotsPerSlide;
    const y = 2.25, boxH = 2.85;

    chunk.forEach((step, i) => {
      const x = 0.5 + i * (boxW + connectorW + gap);
      const globalIdx = ci * MAX + i;
      const label = typeof step === 'string' ? step : step.label || step.title || step.text || '';
      const description = typeof step === 'string' ? '' : step.desc || step.description || step.body || '';
      const emoji = typeof step === 'string' ? '' : step.emoji || step.icon || '';

      slide.addShape('roundRect', { x, y, w: boxW, h: boxH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: i % 2 === 0 ? C.primary : C.accent, width: 1.25 } });
      slide.addShape('ellipse', { x: x + 0.16, y: y + 0.16, w: 0.42, h: 0.42, fill: { color: i % 2 === 0 ? C.primary : C.accent }, line: { type: 'none' } });
      slide.addText(String(globalIdx + 1), { x: x + 0.16, y: y + 0.16, w: 0.42, h: 0.42, align: 'center', valign: 'middle', fontSize: 8, bold: true, color: C.textLight, fontFace: FONTS.englishBody, margin: 0 });
      if (emoji) slide.addText(emoji, { x: x + 0.18, y: y + 0.72, w: boxW - 0.36, h: 0.5, align: 'center', valign: 'middle', fontSize: 24, fontFace: FONTS.englishBody, margin: 0 });
      slide.addText(label, { x: x + 0.14, y: emoji ? y + 1.35 : y + 0.9, w: boxW - 0.28, h: emoji ? 0.72 : 1.05, align: 'center', valign: 'middle', fontSize: 13, bold: true, color: C.cardText, fontFace: detectFont(label, 'title'), fit: 'shrink', margin: 0.02 });
      if (description) slide.addText(description, { x: x + 0.16, y: y + 2.18, w: boxW - 0.32, h: 0.42, align: 'center', valign: 'top', fontSize: 9, color: C.textMuted, fontFace: detectFont(description), fit: 'shrink', margin: 0 });

      if (i < n - 1) {
        slide.addText('›', { x: x + boxW, y: y + 1.12, w: connectorW + gap, h: 0.62, align: 'center', valign: 'middle', fontSize: 30, bold: true, color: C.accent, fontFace: FONTS.englishTitle, margin: 0 });
      }
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §14. Infographic Category Columns

### 14A. Category Column Grid (Bulleted Text)
```javascript
function addCategoryColumns(pres, title, columns, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const n = columns.length;
  const gap = 0.2;
  const colW = (13.33 - (n + 1) * gap) / n;
  const startY = 1.2;
  const colH = 6.0;
  const headerH = 0.8;

  const themes = [
    { primary: C.primary, light: C.bgAlt },
    { primary: C.accent, light: C.cardBgAlt },
    { primary: C.headerBar, light: C.bgAlt }
  ];

  columns.forEach((col, i) => {
    const x = gap + i * (colW + gap);
    const theme = themes[i % themes.length];

    slide.addShape('rect', { x, y: startY, w: colW, h: colH, fill: { color: theme.light } });
    slide.addShape('rect', { x, y: startY, w: colW, h: headerH, fill: { color: theme.primary } });
    const headerText = col.emoji ? `${col.emoji}  ${col.title}` : col.title;
    slide.addText(headerText, { x, y: startY, w: colW, h: headerH, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: C.textLight, fontFace: detectFont(headerText, 'title') });

    const bullets = col.items.map(item => ({
      text: item, options: { bullet: { code: '2022' }, fontSize: 16, color: C.cardText, fontFace: detectFont(item), breakLine: true }
    }));
    slide.addText(bullets, { x: x + 0.2, y: startY + headerH + 0.2, w: colW - 0.4, h: colH - headerH - 0.4, valign: 'top', paraSpaceAfter: 12 });
  });

  return slide;
}
```

### 14B. Category Column List (Rows / Cards)
```javascript
function addCategoryListColumns(pres, title, columns, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const n = columns.length;
  const gap = 0.2;
  const colW = (13.33 - (n + 1) * gap) / n;
  const startY = 1.2;
  const colH = 6.0;
  const headerH = 0.8;

  const themes = [
    { primary: C.primary, light: C.bgAlt },
    { primary: C.accent, light: C.cardBgAlt },
    { primary: C.headerBar, light: C.bgAlt }
  ];

  columns.forEach((col, i) => {
    const x = gap + i * (colW + gap);
    const theme = themes[i % themes.length];

    slide.addShape('rect', { x, y: startY, w: colW, h: colH, fill: { color: theme.light } });
    slide.addShape('rect', { x, y: startY, w: colW, h: headerH, fill: { color: theme.primary } });
    
    const headerText = col.emoji ? `${col.emoji}  ${col.title}` : col.title;
    slide.addText(headerText, { x, y: startY, w: colW, h: headerH, align: 'center', valign: 'middle', fontSize: 22, bold: true, color: C.textLight, fontFace: detectFont(headerText, 'title') });

    const rowGap = 0.15;
    const rowH = (colH - headerH - rowGap) / col.items.length - rowGap;
    let currentY = startY + headerH + rowGap;

    col.items.forEach(item => {
      slide.addShape('rect', { x: x + 0.15, y: currentY, w: colW - 0.3, h: rowH, fill: { color: C.cardBg }, line: { color: theme.primary, width: 1 } });
      
      if (item.desc) {
         slide.addText([
           { text: item.title, options: { bold: true, color: theme.primary, breakLine: true } },
           { text: item.desc, options: { color: C.cardText } }
         ], { x: x + 0.25, y: currentY, w: colW - 0.5, h: rowH, valign: 'middle', fontSize: 16, fontFace: detectFont(item.title) });
      } else {
         slide.addText(item.title, { x: x + 0.25, y: currentY, w: colW - 0.5, h: rowH, valign: 'middle', fontSize: 16, color: C.cardText, fontFace: detectFont(item.title) });
      }
      currentY += rowH + rowGap;
    });
  });

  return slide;
}
```

---

## §15. Native Charts (Bar Chart)

### 15A. Standard Bar Chart Slide
```javascript
function addBarChartSlide(pres, title, chartData, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  let textColor = isLightBg(C) ? C.textDark : C.textLight;

  let opts = {
    x: 0.8, y: 1.5, w: 11.73, h: 5.3,
    barDir: 'col',
    barGrouping: 'clustered',
    showLegend: true, legendPos: 'b', legendColor: textColor, legendFontFace: FONTS.englishBody, legendFontSize: 14,
    chartColors: [C.primary, C.accent, C.border],
    valGridLine: { color: C.border, style: 'solid' },
    valAxisLineShow: false,
    catAxisLineShow: true, catAxisLineColor: C.border,
    valAxisLabelColor: textColor, catAxisLabelColor: textColor,
    catAxisLabelFontFace: FONTS.englishBody, valAxisLabelFontFace: FONTS.englishBody,
    catAxisLabelFontSize: 14, valAxisLabelFontSize: 14,
    showValue: true, valAxisLabelFormatCode: '#,##0',
    dataLabelColor: textColor, dataLabelPosition: 'outEnd', dataLabelFontFace: FONTS.englishBody, dataLabelFontSize: 12
  };

  slide.addChart(pres.ChartType.bar, chartData, opts);
  return slide;
}
```
---

## §16. Project Roadmap (Timeline)

```javascript
function addTimeline(pres, title, milestones, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const n = milestones.length, startX = 1.6, endX = 11.73;
  const gap = n > 1 ? (endX - startX) / (n - 1) : 0;
  const yLine = 4.0;

  slide.addShape('line', { x: startX, y: yLine, w: endX - startX, h: 0, line: { color: C.primary, width: 3 } });

  milestones.forEach((ms, i) => {
    const x = startX + i * gap;
    // Milestone dot with glow
    slide.addShape('ellipse', { x: x - 0.2, y: yLine - 0.2, w: 0.4, h: 0.4, fill: { color: C.accent }, line: { color: C.bg, width: 2 }, shadow: { type: 'outer', color: C.accent, opacity: 0.25, blur: 4, offset: 0 } });
    
    const isTop = i % 2 === 0;
    const textY = isTop ? yLine - 1.2 : yLine + 0.4;
    // Vertical connector line
    const connY = isTop ? yLine - 0.2 : yLine + 0.2;
    const connH = isTop ? -(yLine - textY - 0.5) : textY - yLine;
    slide.addShape('line', { x: x, y: Math.min(connY, textY + 0.4), w: 0, h: Math.abs(connH), line: { color: C.border, width: 1, dashType: 'dash' } });

    slide.addText(ms.date, { x: x - 1.5, y: textY, w: 3.0, h: 0.4, align: 'center', fontSize: 16, bold: true, color: C.primary, fontFace: detectFont(ms.date, 'title') });
    slide.addText(ms.title, { x: x - 1.5, y: textY + 0.4, w: 3.0, h: 0.4, align: 'center', fontSize: 14, color: isLightBg(C) ? C.textDark : C.textLight, fontFace: detectFont(ms.title) });
  });

  return slide;
}
```

---

## §17. Continuous Integration Cycle

```javascript
function addCycleDiagram(pres, title, steps, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const cx = 6.66, cy = 4.2, r = 1.8;
  const n = steps.length;
  const angleStep = (2 * Math.PI) / n;

  slide.addShape('ellipse', { x: cx - r, y: cy - r, w: r*2, h: r*2, fill: { type: 'none' }, line: { color: C.primary, width: 4, dashType: 'sysDash' } });

  steps.forEach((step, i) => {
    const angle = i * angleStep - Math.PI / 2;
    const x = cx + r * Math.cos(angle);
    const y = cy + r * Math.sin(angle);

    const stepLabel = typeof step === 'string' ? step : step.label || step.title || step.text;
    const stepDesc = typeof step === 'string' ? '' : step.desc || step.body || step.description;
    const stepEmoji = typeof step === 'string' ? '' : step.emoji;

    slide.addShape('ellipse', { x: x - 0.5, y: y - 0.5, w: 1.0, h: 1.0, fill: { color: C.cardBg }, line: { color: C.accent, width: 2 }, shadow: { type: 'outer', blur: 4, offset: 2, opacity: 0.15 } });
    
    if (stepEmoji) {
      slide.addText(stepEmoji, { x: x - 0.5, y: y - 0.5, w: 1.0, h: 1.0, align: 'center', valign: 'middle', fontSize: 24, fontFace: FONTS.englishBody });
    } else {
      slide.addText(String(i + 1), { x: x - 0.5, y: y - 0.5, w: 1.0, h: 1.0, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: C.primary, fontFace: FONTS.englishTitle });
    }

    const textRadius = r + 1.2;
    const tx = cx + textRadius * Math.cos(angle) - 1.5;
    // The top node is close to the header bar. Place its copy below the
    // icon instead of above it so labels can never intrude into the title.
    // Side and bottom nodes retain their outward-facing label placement.
    const isTopNode = Math.sin(angle) < -0.85;
    const ty = isTopNode ? y + 0.62 : cy + textRadius * Math.sin(angle) - 0.6;

    if (stepDesc) {
      slide.addText(stepLabel, { x: tx, y: ty, w: 3.0, h: 0.4, align: 'center', valign: 'bottom', fontSize: 14, bold: true, color: C.textDark || C.primary, fontFace: detectFont(stepLabel, 'title') });
      slide.addText(stepDesc, { x: tx, y: ty + 0.4, w: 3.0, h: 0.8, align: 'center', valign: 'top', fontSize: 11, color: C.cardText || C.textBody, fontFace: detectFont(stepDesc) });
    } else {
      slide.addText(stepLabel, { x: tx, y: ty, w: 3.0, h: 1.2, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: C.textDark || C.primary, fontFace: detectFont(stepLabel, 'title') });
    }
  });

  return slide;
}
```

---

## §18. Funnel (Sales / Conversion)

```javascript
function addFunnel(pres, title, stages, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const n = stages.length;
  let currentW = 8.0, currentY = 1.5, h = 4.5 / n;
  const colors = [C.primary, C.accent, C.headerBar, C.primary, C.accent];

  stages.forEach((stage, i) => {
    const nextW = currentW - (6.0 / n);
    slide.addShape('rect', { x: (13.33 - currentW) / 2, y: currentY, w: currentW, h: h - 0.1, fill: { color: colors[i % colors.length] }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.08, blur: 4, offset: 1 } });
    slide.addText(stage.label + ' : ' + stage.value, { x: (13.33 - currentW) / 2, y: currentY, w: currentW, h: h - 0.1, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.textLight, fontFace: detectFont(stage.label, 'title') });
    
    currentW = nextW;
    currentY += h;
  });

  return slide;
}
```

---

## §19. Venn Diagram (Product Strategy)

```javascript
function addVennDiagram(pres, title, labelLeft, labelRight, labelCenter, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const yCenter = 3.5, radius = 4.0;
  const textColor = isLightBg(C) ? C.textDark : C.textLight;
  
  slide.addShape('ellipse', { x: 2.5, y: 1.5, w: radius, h: radius, fill: { color: C.primary, transparency: 50 }, line: { color: C.primary, width: 2 } });
  slide.addShape('ellipse', { x: 5.5, y: 1.5, w: radius, h: radius, fill: { color: C.accent, transparency: 50 }, line: { color: C.accent, width: 2 } });

  slide.addText(labelLeft, { x: 2.5, y: 2.5, w: 2.5, h: 2.0, align: 'center', valign: 'middle', fontSize: 24, bold: true, color: textColor, fontFace: detectFont(labelLeft) });
  slide.addText(labelRight, { x: 7.0, y: 2.5, w: 2.5, h: 2.0, align: 'center', valign: 'middle', fontSize: 24, bold: true, color: textColor, fontFace: detectFont(labelRight) });
  slide.addText(labelCenter, { x: 4.5, y: 2.5, w: 3.0, h: 2.0, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: textColor, fontFace: detectFont(labelCenter) });

  return slide;
}
```

---

## §20. Organization Classification Tree

```javascript
function addHierarchyTree(pres, title, rootNode, childNodes, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const rootLabel = typeof rootNode === 'string' ? rootNode : rootNode.label;
  const rootDesc = typeof rootNode === 'string' ? '' : rootNode.desc;

  slide.addShape('rect', { x: 4.66, y: 1.2, w: 4.0, h: 0.9, fill: { color: C.primary }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.12, blur: 6, offset: 2 } });
  
  if (rootDesc) {
    slide.addText(rootLabel, { x: 4.66, y: 1.25, w: 4.0, h: 0.45, align: 'center', valign: 'bottom', fontSize: 18, bold: true, color: C.textLight, fontFace: detectFont(rootLabel, 'title') });
    slide.addText(rootDesc, { x: 4.76, y: 1.7, w: 3.8, h: 0.35, align: 'center', valign: 'top', fontSize: 12, color: C.bgAlt, fontFace: detectFont(rootDesc) });
  } else {
    slide.addText(rootLabel, { x: 4.66, y: 1.2, w: 4.0, h: 0.9, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: C.textLight, fontFace: detectFont(rootLabel, 'title') });
  }

  slide.addShape('line', { x: 6.66, y: 2.1, w: 0, h: 0.4, line: { color: C.primary, width: 2 } });

  const n = childNodes.length;
  if (n > 0) {
    const gap = Math.min(3.2, 12.6 / n);
    const startX = (13.33 - (gap * (n - 1))) / 2;
    const nodeW = gap - 0.2;
    
    slide.addShape('line', { x: startX, y: 2.5, w: gap * (n - 1), h: 0, line: { color: C.primary, width: 2 } });

    childNodes.forEach((child, i) => {
      const cx = startX + i * gap;
      const childLabel = typeof child === 'string' ? child : child.label;
      const childDesc = typeof child === 'string' ? '' : child.desc;
      const children2 = typeof child === 'string' ? [] : (child.children || []);

      slide.addShape('line', { x: cx, y: 2.5, w: 0, h: 0.3, line: { color: C.primary, width: 2 } });

      slide.addShape('rect', { x: cx - nodeW/2, y: 2.8, w: nodeW, h: 0.9, fill: { color: C.accent }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 4, offset: 1 } });
      
      if (childDesc) {
        slide.addText(childLabel, { x: cx - nodeW/2, y: 2.85, w: nodeW, h: 0.45, align: 'center', valign: 'bottom', fontSize: 15, bold: true, color: C.textLight, fontFace: detectFont(childLabel, 'title') });
        slide.addText(childDesc, { x: cx - nodeW/2 + 0.1, y: 3.3, w: nodeW - 0.2, h: 0.35, align: 'center', valign: 'top', fontSize: 11, color: C.bg, fontFace: detectFont(childDesc) });
      } else {
        slide.addText(childLabel, { x: cx - nodeW/2, y: 2.8, w: nodeW, h: 0.9, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.textLight, fontFace: detectFont(childLabel, 'title') });
      }

      if (children2.length > 0) {
        const trunkX = cx - nodeW/2 + 0.2;
        const startY2 = 3.7;
        const itemH = 0.55;
        const itemGap = 0.1;
        
        slide.addShape('line', { x: trunkX, y: 3.7, w: 0, h: children2.length * (itemH + itemGap) - (itemH/2), line: { color: C.primary, width: 1.5 } });

        children2.forEach((gc, j) => {
          const gcLabel = typeof gc === 'string' ? gc : gc.label;
          const gcDesc = typeof gc === 'string' ? '' : gc.desc;
          const gcY = startY2 + 0.2 + j * (itemH + itemGap);
          
          slide.addShape('line', { x: trunkX, y: gcY + itemH/2, w: 0.2, h: 0, line: { color: C.primary, width: 1.5 } });
          
          slide.addShape('roundRect', { x: trunkX + 0.2, y: gcY, w: nodeW - 0.3, h: itemH, rectRadius: 0.05, fill: { color: C.cardBg }, line: { color: C.border, width: 1 } });
          
          if (gcDesc) {
            slide.addText(gcLabel, { x: trunkX + 0.25, y: gcY + 0.05, w: nodeW - 0.4, h: 0.25, align: 'left', valign: 'middle', fontSize: 11, bold: true, color: C.cardText, fontFace: detectFont(gcLabel, 'title') });
            slide.addText(gcDesc, { x: trunkX + 0.25, y: gcY + 0.3, w: nodeW - 0.4, h: 0.2, align: 'left', valign: 'top', fontSize: 9, color: C.textMuted, fontFace: detectFont(gcDesc) });
          } else {
            slide.addText(gcLabel, { x: trunkX + 0.25, y: gcY, w: nodeW - 0.4, h: itemH, align: 'left', valign: 'middle', fontSize: 12, bold: true, color: C.cardText, fontFace: detectFont(gcLabel) });
          }
        });
      }
    });
  }

  return slide;
}
```

---

## §21. Glossary of Terms (Vocab Grid, auto-replicates)

```javascript
function addVocabGrid(pres, title, terms, C) {
  const MAX = 6;
  const chunks = chunkArray(terms, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const cols = 2, colW = 5.8, rowH = 1.6, startX = 0.6, startY = 1.3, gapX = 0.5, gapY = 0.3;

    chunk.forEach((term, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = startX + col * (colW + gapX);
      const y = startY + row * (rowH + gapY);

      slide.addShape('roundRect', { x, y, w: colW, h: rowH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', color: C.shadow, opacity: 0.08, blur: 4, offset: 1 } });
      slide.addText(term.word, { x: x + 0.2, y: y + 0.1, w: colW - 0.4, h: 0.4, fontSize: 18, bold: true, color: C.primary, fontFace: detectFont(term.word, 'title') });
      slide.addText(term.definition, { x: x + 0.2, y: y + 0.5, w: colW - 0.4, h: 1.0, fontSize: 14, color: C.cardText, fontFace: detectFont(term.definition), valign: 'top' });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §22. Quiz Slide / Answer Reveal

```javascript
function addQuizSlide(pres, question, options, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, 'Quick Check / Quiz', C);

  slide.addText(question, { x: 1.0, y: 1.5, w: 11.33, h: 1.5, align: 'center', valign: 'middle', fontSize: 32, bold: true, color: isLightBg(C) ? C.textDark : C.textLight, fontFace: detectFont(question, 'title') });

  const n = options.length;
  const yStart = 3.5, gap = 0.2, rowH = 0.8;

  options.forEach((opt, i) => {
    const y = yStart + i * (rowH + gap);
    const prefix = String.fromCharCode(65 + i);
    
    // Option letter badge circle
    slide.addShape('ellipse', { x: 2.0, y: y + 0.1, w: 0.55, h: 0.55, fill: { color: C.accent }, line: { type: 'none' } });
    slide.addText(prefix, { x: 2.0, y: y + 0.1, w: 0.55, h: 0.55, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.textLight, fontFace: FONTS.englishTitle });

    // Option box
    slide.addShape('roundRect', { x: 2.7, y, w: 8.13, h: rowH, rectRadius: 0.1, fill: { color: C.cardBg }, line: { color: C.border, width: 1 } });
    slide.addText(opt, { x: 2.9, y, w: 7.8, h: rowH, valign: 'middle', fontSize: 18, color: C.cardText, fontFace: detectFont(opt) });
  });

  return slide;
}
```

---

## §23. Icon List with Line Separators (auto-replicates)

```javascript
function addIconListWithLines(pres, title, items, C) {
  const MAX = 5;
  const chunks = chunkArray(items, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const startY = 1.4, rowH = 1.05, lineX = 1.6, lineW = 11.0;
    const textColor = isLightBg(C) ? C.textDark : C.textLight;

    chunk.forEach((item, i) => {
      const y = startY + i * rowH;

      slide.addShape('ellipse', { x: 0.5, y: y + 0.1, w: 0.7, h: 0.7, fill: { color: C.primary }, line: { type: 'none' } });
      if (item.emoji) {
        slide.addText(item.emoji, { x: 0.5, y: y + 0.1, w: 0.7, h: 0.7, align: 'center', valign: 'middle', fontSize: 20, fontFace: FONTS.englishBody });
      }

      slide.addText(item.text, { x: lineX, y: y + 0.05, w: lineW - 0.5, h: 0.6, valign: 'middle', fontSize: 22, bold: true, color: textColor, fontFace: detectFont(item.text) });

      slide.addShape('line', { x: lineX, y: y + rowH - 0.15, w: lineW - 0.5, h: 0, line: { color: C.border, width: 1 } });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §24. Blockquote / Inspirational Quote Slide

```javascript
function addQuoteSlide(pres, quoteText, authorName, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.primary };
  slide.transition = { type: 'fade', speed: 'med' };
  addDecoCircles(slide, { ...C, primary: C.accent });
  addDiagonalStripe(slide, { ...C, accent: C.bg, primary: C.bg });

  slide.addText('❝', { x: 4.5, y: 0.3, w: 4.33, h: 1.5, align: 'center', valign: 'middle', fontSize: 96, color: C.accent, fontFace: FONTS.englishTitle });

  slide.addText(quoteText, { x: 1.5, y: 2.0, w: 10.33, h: 3.0, align: 'center', valign: 'middle', fontSize: 24, italic: true, color: C.textLight, fontFace: detectFont(quoteText), lineSpacingMultiple: 1.4 });

  // Closing quotation mark
  slide.addText('❞', { x: 4.5, y: 5.0, w: 4.33, h: 0.8, align: 'center', valign: 'middle', fontSize: 48, color: C.accent, fontFace: FONTS.englishTitle });

  slide.addShape('rect', { x: 5.8, y: 5.8, w: 1.73, h: 0.06, fill: { color: C.accent }, line: { type: 'none' } });

  // Author attribution
  if (authorName) {
    slide.addText('— ' + authorName, { x: 3, y: 6.0, w: 7.33, h: 0.5, align: 'center', fontSize: 16, color: C.textMuted, fontFace: detectFont(authorName) });
  }

  return slide;
}
```

---

## §25. Left Panel + Icon List (Split Layout)

```javascript
function addLeftPanelIconList(pres, title, panelEmoji, panelCaption, items, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const panelW = 4.2;
  const textColor = isLightBg(C) ? C.textDark : C.textLight;

  slide.addShape('roundRect', { x: 0.3, y: 1.3, w: panelW, h: 5.8, rectRadius: 0.1, fill: { color: C.primary }, line: { type: 'none' } });

  slide.addShape('ellipse', { x: 1.1, y: 2.5, w: 2.6, h: 2.6, fill: { color: C.accent }, line: { type: 'none' } });
  if (panelEmoji) {
    slide.addText(panelEmoji, { x: 1.1, y: 2.5, w: 2.6, h: 2.6, align: 'center', valign: 'middle', fontSize: 56, fontFace: FONTS.englishBody });
  }

  if (panelCaption) {
    slide.addText(panelCaption, { x: 0.5, y: 5.4, w: panelW - 0.4, h: 1.2, align: 'center', valign: 'middle', fontSize: 16, italic: true, color: C.textLight, fontFace: detectFont(panelCaption) });
  }

  const rightX = panelW + 0.8, rightW = 13.33 - rightX - 0.3;
  const startY = 1.5, rowH = 1.1;

  items.forEach((item, i) => {
    const y = startY + i * rowH;

    slide.addShape('ellipse', { x: rightX, y: y + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent }, line: { type: 'none' } });
    if (item.emoji) {
      slide.addText(item.emoji, { x: rightX, y: y + 0.15, w: 0.6, h: 0.6, align: 'center', valign: 'middle', fontSize: 16, fontFace: FONTS.englishBody });
    }

    slide.addText(item.text, { x: rightX + 0.8, y: y + 0.05, w: rightW - 1.0, h: 0.6, valign: 'middle', fontSize: 20, bold: true, color: textColor, fontFace: detectFont(item.text) });

    slide.addShape('line', { x: rightX + 0.8, y: y + rowH - 0.2, w: rightW - 1.0, h: 0, line: { color: C.border, width: 1 } });
  });

  return slide;
}
```

---

## §26. Dark Icon Card Row (auto-replicates)

```javascript
function addDarkIconCardRow(pres, title, subtitle, cards, C) {
  const MAX = 4;
  const chunks = chunkArray(cards, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.primary };

    slide.addText(slideTitle, { x: 0.5, y: 0.4, w: 12.33, h: 0.8, fontSize: 36, bold: true, color: C.textLight, fontFace: detectFont(slideTitle, 'title') });

    if (subtitle && ci === 0) {
      slide.addText(subtitle, { x: 0.5, y: 1.2, w: 12.33, h: 0.5, fontSize: 16, italic: true, color: C.textLight, fontFace: detectFont(subtitle) });
    }

    const n = Math.min(cards.length, MAX);
    const gap = 0.3, cardW = (12.33 - gap * (n - 1)) / n;
    const cardH = 4.8, startY = 2.2;

    chunk.forEach((card, i) => {
      const x = 0.5 + i * (cardW + gap);

      slide.addShape('roundRect', { x, y: startY, w: cardW, h: cardH, rectRadius: 0.12, fill: { color: C.cardBg }, line: { type: 'none' } });

      slide.addShape('ellipse', { x: x + cardW / 2 - 0.45, y: startY + 0.5, w: 0.9, h: 0.9, fill: { color: C.accent }, line: { type: 'none' } });
      if (card.emoji) {
        slide.addText(card.emoji, { x: x + cardW / 2 - 0.45, y: startY + 0.5, w: 0.9, h: 0.9, align: 'center', valign: 'middle', fontSize: 24, fontFace: FONTS.englishBody });
      }

      slide.addText(card.text, { x: x + 0.2, y: startY + 1.8, w: cardW - 0.4, h: cardH - 2.2, align: 'center', valign: 'middle', fontSize: 22, bold: true, color: C.cardText, fontFace: detectFont(card.text) });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §27. Content Block with Icon Row

```javascript
function addContentWithIconRow(pres, title, paragraph, subLabel, icons, C) {
  const chunks = chunkArray(icons || [], 4);
  const pages = chunks.length ? chunks : [[]];
  const slides = [];

  pages.forEach((chunk, pageIndex) => {
    const slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, pageIndex === 0 ? title : `${title} (cont.)`, C);

    if (pageIndex === 0) {
      // A concise editorial summary replaces the oversized, mostly empty panel.
      slide.addShape('roundRect', { x: 0.5, y: 1.15, w: 12.33, h: 1.25, rectRadius: 0.07, fill: { color: C.bgAlt }, line: { color: C.border, width: 1 } });
      slide.addShape('rect', { x: 0.5, y: 1.15, w: 0.1, h: 1.25, fill: { color: C.accent }, line: { type: 'none' } });
      slide.addText(paragraph, { x: 0.82, y: 1.43, w: 11.65, h: 0.62, fontSize: 15, color: C.textBody, fontFace: detectFont(paragraph), valign: 'middle', fit: 'shrink', margin: 0, lineSpacingMultiple: 1.18 });
    }

    const labelY = pageIndex === 0 ? 2.75 : 1.35;
    // Reserve clear breathing room after the accent underline before cards begin.
    const cardY = labelY + 0.74;
    const cardH = pageIndex === 0 ? 2.8 : 4.25;
    slide.addText(pageIndex === 0 ? subLabel : `${subLabel} (cont.)`, { x: 0.5, y: labelY, w: 8.8, h: 0.35, fontSize: 18, bold: true, color: C.primary, fontFace: detectFont(subLabel, 'title'), fit: 'shrink', margin: 0 });
    slide.addShape('rect', { x: 0.5, y: labelY + 0.48, w: 1.25, h: 0.06, fill: { color: C.accent }, line: { type: 'none' } });

    const n = Math.max(chunk.length, 1), gap = 0.24, cardW = (12.33 - gap * (n - 1)) / n;
    chunk.forEach((icon, i) => {
      const x = 0.5 + i * (cardW + gap);
      slide.addShape('roundRect', { x, y: cardY, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 4, offset: 1 } });
      slide.addShape('rect', { x, y: cardY, w: cardW, h: 0.08, fill: { color: i % 2 === 0 ? C.primary : C.accent }, line: { type: 'none' } });
      slide.addShape('ellipse', { x: x + 0.28, y: cardY + 0.33, w: 0.54, h: 0.54, fill: { color: C.bgAlt }, line: { type: 'none' } });
      slide.addText(String(i + 1 + pageIndex * 4).padStart(2, '0'), { x: x + 0.28, y: cardY + 0.33, w: 0.54, h: 0.54, align: 'center', valign: 'middle', fontSize: 9, bold: true, color: C.primary, fontFace: FONTS.englishBody, margin: 0 });
      if (icon.emoji) slide.addText(icon.emoji, { x: x + 0.25, y: cardY + 1.03, w: cardW - 0.5, h: 0.55, align: 'center', valign: 'middle', fontSize: 28, fontFace: FONTS.englishBody, margin: 0 });
      slide.addText(icon.label || '', { x: x + 0.22, y: cardY + 1.78, w: cardW - 0.44, h: 0.5, align: 'center', valign: 'middle', fontSize: 15, bold: true, color: C.cardText, fontFace: detectFont(icon.label, 'title'), fit: 'shrink', margin: 0 });
      if (icon.desc) slide.addText(icon.desc, { x: x + 0.25, y: cardY + 2.33, w: cardW - 0.5, h: cardH - 2.58, align: 'center', valign: 'top', fontSize: 11, color: C.textMuted, fontFace: detectFont(icon.desc), fit: 'shrink', margin: 0 });
    });
    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §28. Summary Overview (Numbered Bar List, auto-replicates)

```javascript
function addSummaryOverview(pres, title, subtitle, items, C) {
  const MAX = 8;
  const chunks = chunkArray(items, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);
    if (ci === 0) addSubtleWave(slide, C);

    if (subtitle && ci === 0) {
      slide.addText(subtitle, { x: 0.5, y: 1.1, w: 12.33, h: 0.5, fontSize: 16, italic: true, color: isLightBg(C) ? C.textDark : C.textLight, fontFace: detectFont(subtitle) });
    }

    const startY = 1.8;
    const totalH = 5.2;
    const maxRowH = 0.95;
    const numItems = Math.max(Math.min(items.length, MAX), 1);
    const gap = numItems > 5 ? 0.08 : 0.15;
    let rowH = (totalH - (gap * (numItems - 1))) / numItems;
    if (rowH > maxRowH) rowH = maxRowH;
    const fontSize = rowH < 0.6 ? 14 : 20;
    const accentW = 0.08;
    const accentColors = [C.accent, C.primary];

    chunk.forEach((item, i) => {
      const globalIdx = ci * MAX + i;
      const y = startY + i * (rowH + gap);

      slide.addShape('rect', { x: 0.5, y, w: accentW, h: rowH, fill: { color: accentColors[i % 2] }, line: { type: 'none' } });
      slide.addShape('rect', { x: 0.5 + accentW, y, w: 12.25 - accentW, h: rowH, fill: { color: C.primary }, line: { type: 'none' } });

      const numText = String(globalIdx + 1) + '.';
      slide.addText(numText, { x: 0.7, y, w: 0.6, h: rowH, valign: 'middle', fontSize: fontSize, bold: true, color: C.textLight, fontFace: FONTS.englishTitle });
      const cleanItem = item.replace(/^[\d০-৯]+[\.\)]\s*/, '');
      slide.addText(cleanItem, { x: 1.4, y, w: 11.2, h: rowH, valign: 'middle', fontSize: fontSize, color: C.textLight, fontFace: detectFont(cleanItem) });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §29. Data Table (auto-replicates)

```javascript
function addDataTable(pres, title, headers, rows, C) {
  const MAX = 8;
  const chunks = chunkArray(rows, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const startX = 0.5, startY = 1.4;
    const tableW = 12.33;
    const colW = tableW / headers.length;
    const headerH = 0.7;
    const rowH = 0.65;

    headers.forEach((hdr, i) => {
      const x = startX + i * colW;
      slide.addShape('rect', { x, y: startY, w: colW, h: headerH, fill: { color: C.primary }, line: { color: C.border, width: 1 } });
      slide.addText(hdr, { x, y: startY, w: colW, h: headerH, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.textLight, fontFace: detectFont(hdr, 'title') });
    });

    chunk.forEach((row, r) => {
      const y = startY + headerH + r * rowH;
      const rowBg = r % 2 === 0 ? C.cardBg : C.bgAlt;

      row.forEach((cell, c) => {
        const x = startX + c * colW;
        slide.addShape('rect', { x, y, w: colW, h: rowH, fill: { color: rowBg }, line: { color: C.border, width: 0.5 } });
        slide.addText(cell, { x: x + 0.15, y, w: colW - 0.3, h: rowH, valign: 'middle', fontSize: 16, color: C.cardText, fontFace: detectFont(cell), bold: c === 0 });
      });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §30. Split Layout — Left Info List + Right Image Placeholder

```javascript
function addSplitInfoImage(pres, title, bullets, imagePath, imageCaption, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const leftW = 6.0, rightX = 6.8, rightW = 6.0;

  const bulletItems = bullets.map(b => ({
    text: b,
    options: {
      bullet: { code: '25CF' }, fontSize: 18, color: C.cardText,
      fontFace: detectFont(b), breakLine: true, paraSpaceAfter: 10
    }
  }));

  slide.addText(bulletItems, { x: 0.5, y: 1.4, w: leftW, h: 5.5, valign: 'top', lineSpacingMultiple: 1.3 });

  if (imagePath) {
    slide.addImage({ path: imagePath, x: rightX, y: 1.4, w: rightW, h: 5.5, sizing: { type: 'cover', w: rightW, h: 5.5 } });
  } else {
    slide.addShape('roundRect', { x: rightX, y: 1.4, w: rightW, h: 5.5, rectRadius: 0.1, fill: { color: C.bgAlt }, line: { color: C.border, width: 2, dashType: 'dash' } });
    slide.addText(imageCaption || '[ Image ]', { x: rightX, y: 1.4, w: rightW, h: 5.5, align: 'center', valign: 'middle', fontSize: 24, italic: true, color: C.textMuted, fontFace: FONTS.englishBody });
  }

  return slide;
}
```

---

## §31. Image Gallery (auto-replicates)

```javascript
function addImageGallery(pres, title, images, C) {
  const MAX = 6;
  const chunks = chunkArray(images, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const colW = 3.8, gap = 0.3, startX = 0.5, startY = 1.3;
    const heights = [[3.2, 2.6], [2.6, 3.2], [3.2, 2.6]];

    chunk.forEach((img, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = startX + col * (colW + gap);
      let y = startY;
      if (row > 0) { y = startY + heights[col][0] + gap; }
      const h = heights[col][row] || 2.6;

      if (img.path) {
        slide.addImage({ path: img.path, x, y, w: colW, h, sizing: { type: 'cover', w: colW, h: h } });
      } else {
        slide.addShape('roundRect', { x, y, w: colW, h, rectRadius: 0.1, fill: { color: C.cardBgAlt }, line: { type: 'none' }, shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 6, offset: 2 } });
        slide.addText(img.label || 'Image ' + (ci * MAX + i + 1), { x, y, w: colW, h, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.primary, fontFace: FONTS.englishBody });
      }

      // Caption below each image
      if (img.caption) {
        slide.addText(img.caption, { x, y: y + h + 0.05, w: colW, h: 0.3, align: 'center', fontSize: 10, color: C.textMuted, fontFace: detectFont(img.caption) });
      }
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §32. Pie Chart with Info Panel

```javascript
function addPieChartSlide(pres, title, chartData, infoBoxTitle, infoBoxBullets, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addDecoCircles(slide, C);
  addHeaderBar(slide, title, C);

  let labels = chartData.map(d => d.label);
  let values = chartData.map(d => d.value);
  let pptxData = [{ name: 'Data', labels: labels, values: values }];
  let chartColors = [C.primary, C.accent, 'E0E0E0', C.cardBgAlt, C.textDark, C.headerBar];
  let textColor = isLightBg(C) ? C.textDark : C.textLight;

  slide.addChart(pres.ChartType.pie, pptxData, {
    x: 0.8, y: 1.5, w: 5.5, h: 4.5,
    showLegend: true, legendPos: 'b', legendColor: textColor,
    showValue: true, dataLabelColor: 'FFFFFF', dataLabelFontSize: 12,
    chartColors: chartColors, border: { type: 'none' }
  });

  if (infoBoxTitle || (infoBoxBullets && infoBoxBullets.length > 0)) {
    slide.addShape('roundRect', { x: 6.8, y: 1.6, w: 5.7, h: 4.3, rectRadius: 0.1, fill: { color: C.cardBg }, line: { color: C.primary, width: 1.5 }, shadow: { type: 'outer', color: C.shadow, opacity: 0.08, blur: 6, offset: 2 } });

    if (infoBoxTitle) {
      slide.addText(infoBoxTitle, { x: 7.2, y: 1.8, w: 5.0, h: 0.5, fontSize: 20, bold: true, color: C.primary, fontFace: detectFont(infoBoxTitle, 'title') });
    }

    if (infoBoxBullets && infoBoxBullets.length > 0) {
      slide.addText(infoBoxBullets.map(b => '• ' + b).join('\n\n'), { x: 7.2, y: 2.4, w: 5.0, h: 3.2, fontSize: 16, color: C.cardText, valign: 'top', fontFace: detectFont(infoBoxBullets[0]), lineSpacingMultiple: 1.3 });
    }
  }

  return slide;
}
```

---

## §33. Line Chart / Trend Graph

```javascript
function addLineChartSlide(pres, title, chartData, seriesName, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  let labels = chartData.map(d => String(d.label));
  let values = chartData.map(d => Number(d.value));
  let pptxData = [{ name: seriesName || 'Data Series', labels: labels, values: values }];
  let textColor = isLightBg(C) ? C.textDark : C.textLight;

  slide.addChart(pres.ChartType.line, pptxData, {
    x: 0.5, y: 1.2, w: 12.33, h: 5.5,
    showLegend: true, legendPos: 'b', legendColor: textColor,
    showValue: true, dataLabelColor: textColor, dataLabelFontSize: 11, dataLabelPosition: 't',
    chartColors: [C.primary, C.accent, C.textMuted],
    lineSize: 4, lineDataSymbol: 'circle', lineDataSymbolSize: 8,
    valGridLine: { color: C.border || 'EAEAEA', style: 'dash' },
    catGridLine: { type: 'none' },
    catAxisLabelColor: textColor, valAxisLabelColor: textColor,
    border: { type: 'none' }
  });

  return slide;
}
```

---

## §34. Table and Comparative Chart (Split)

```javascript
function addTableAndChartSlide(pres, title, dataItems, metrics, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  let tableW = 5.2, chartW = 7.5, rowH = 0.55, startX = 0.3, startY = 1.6;
  let colW = tableW / (metrics[1] ? 3 : 2);

  slide.addShape('rect', { x: startX, y: startY, w: tableW, h: rowH, fill: { color: C.primary } });
  slide.addText('Category', { x: startX, y: startY, w: colW, h: rowH, color: 'FFFFFF', bold: true, fontSize: 13, align: 'center', fontFace: FONTS.englishTitle });
  if (metrics[0]) slide.addText(metrics[0], { x: startX + colW, y: startY, w: colW, h: rowH, color: 'FFFFFF', bold: true, fontSize: 13, align: 'center', fontFace: FONTS.englishTitle });
  if (metrics[1]) slide.addText(metrics[1], { x: startX + colW*2, y: startY, w: colW, h: rowH, color: 'FFFFFF', bold: true, fontSize: 13, align: 'center', fontFace: FONTS.englishTitle });

  dataItems.forEach((item, i) => {
    let y = startY + rowH + (i * rowH);
    let bgColor = (i % 2 === 0) ? (C.cardBg || 'FFFFFF') : (C.bgAlt || 'F0F0F0');
    slide.addShape('rect', { x: startX, y: y, w: tableW, h: rowH, fill: { color: bgColor }, line: { color: C.border || 'E0E0E0', width: 0.5 } });
    slide.addText(item.label, { x: startX, y: y, w: colW, h: rowH, color: C.cardText, fontSize: 13, align: 'center', fontFace: detectFont(item.label) });
    
    let val1 = item.values[0] !== undefined ? String(item.values[0]) : '';
    slide.addText(val1, { x: startX + colW, y: y, w: colW, h: rowH, color: C.cardText, fontSize: 13, align: 'center', fontFace: FONTS.englishBody });
    
    if (metrics[1]) {
      let val2 = item.values[1] !== undefined ? String(item.values[1]) : '';
      slide.addText(val2, { x: startX + colW*2, y: y, w: colW, h: rowH, color: C.cardText, fontSize: 13, align: 'center', fontFace: FONTS.englishBody });
    }
  });

  let chartLabels = dataItems.map(d => String(d.label));
  let pptxData = [];
  if (metrics[0]) pptxData.push({ name: metrics[0], labels: chartLabels, values: dataItems.map(d => Number(d.values[0]) || 0) });
  if (metrics[1]) pptxData.push({ name: metrics[1], labels: chartLabels, values: dataItems.map(d => Number(d.values[1]) || 0) });
  let textColor = isLightBg(C) ? C.textDark : C.textLight;

  slide.addChart(pres.ChartType.bar, pptxData, {
    x: 5.7, y: 1.2, w: chartW, h: 5.5,
    barDir: 'col', showLegend: true, legendPos: 'b', legendColor: textColor,
    showValue: true, dataLabelColor: textColor, dataLabelFontSize: 11, dataLabelPosition: 'outEnd',
    chartColors: [C.primary, C.accent],
    valGridLine: { color: C.border || 'EAEAEA', style: 'solid' },
    catGridLine: { type: 'none' }, catAxisLabelColor: textColor, valAxisLabelColor: textColor,
    border: { type: 'none' }
  });

  return slide;
}
```

---

## §35. Visual Placeholder with Right Cards

```javascript
function addVisualPlaceholderWithCards(pres, title, placeholderTitle, placeholderSubtitle, cardItems, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  slide.addShape('roundRect', { x: 0.3, y: 1.4, w: 7.8, h: 5.0, rectRadius: 0.08, fill: { color: C.cardBgAlt || 'E8F5E9' }, line: { color: 'E0E0E0', width: 1.5, dashType: 'dash' } });
  
  if (placeholderTitle) {
    slide.addText('🗺️ ' + placeholderTitle, { x: 0.3, y: 3.4, w: 7.8, h: 0.4, align: 'center', color: isLightBg(C) ? C.textDark : C.textLight, bold: true, fontSize: 18, fontFace: detectFont(placeholderTitle, 'title') });
  }
  if (placeholderSubtitle) {
    slide.addText('[ ' + placeholderSubtitle + ' ]', { x: 0.3, y: 3.9, w: 7.8, h: 0.4, align: 'center', color: C.cardText, bold: true, fontSize: 16, fontFace: detectFont(placeholderSubtitle) });
  }

  let cardH = (5.0 / Math.max(cardItems.length, 1)) - 0.15;
  let startY = 1.4;
  
  cardItems.slice(0, 6).forEach((item, i) => {
    let y = startY + i * (cardH + 0.15);
    slide.addShape('roundRect', { x: 8.3, y: y, w: 4.8, h: cardH, rectRadius: 0.1, fill: { color: C.cardBg || 'FFFFFF' }, line: { color: C.primary, width: 1 } });
    slide.addText(item.title, { x: 8.5, y: y + 0.05, w: 4.4, h: cardH * 0.4, color: C.primary, bold: true, fontSize: 14, fontFace: detectFont(item.title, 'title') });
    slide.addText(item.desc, { x: 8.5, y: y + (cardH * 0.4), w: 4.4, h: cardH * 0.5, color: C.cardText, fontSize: 12, valign: 'top', fontFace: detectFont(item.desc) });
  });

  return slide;
}
```

---

## §36. Table with Summary Box (auto-replicates table)

```javascript
function addTableWithSummarySlide(pres, title, headers, tableData, summaryTitle, summaryText, C) {
  const MAX_ROWS = 6;
  const chunks = chunkArray(tableData, MAX_ROWS);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const isLast = ci === chunks.length - 1;
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    let startX = 0.5, startY = 1.3, tableW = 12.33;
    let numCols = Math.max(headers.length, 1);
    let colW = tableW / numCols;
    let rowH = 0.45;

    headers.forEach((header, i) => {
      slide.addShape('rect', { x: startX + (i * colW), y: startY, w: colW, h: rowH, fill: { color: C.primary }, line: { color: 'E0E0E0', width: 0.5 } });
      slide.addText(header, { x: startX + (i * colW), y: startY, w: colW, h: rowH, align: 'center', valign: 'middle', color: 'FFFFFF', bold: true, fontSize: 13, fontFace: detectFont(header, 'title') });
    });

    chunk.forEach((row, rowIndex) => {
      let y = startY + rowH + (rowIndex * rowH);
      let bgColor = (rowIndex % 2 === 0) ? (C.cardBg || 'FFFFFF') : (C.bgAlt || 'F0F8F0');
      let rowVals = Array.isArray(row) ? row : (row.values ? [row.label, ...row.values] : Object.values(row));
      
      for (let i = 0; i < numCols; i++) {
        slide.addShape('rect', { x: startX + (i * colW), y: y, w: colW, h: rowH, fill: { color: bgColor }, line: { color: C.border || 'E0E0E0', width: 0.5 } });
        let val = rowVals[i] !== undefined ? String(rowVals[i]) : '';
        slide.addText(val, { x: startX + (i * colW), y: y, w: colW, h: rowH, align: 'center', valign: 'middle', color: C.cardText || C.textDark, fontSize: 13, fontFace: detectFont(val) });
      }
    });

    // Summary box — only on the LAST chunk slide
    if (isLast && (summaryTitle || summaryText)) {
      let tableBottomY = startY + rowH + (chunk.length * rowH);
      let summaryY = tableBottomY + 0.3;
      let summaryH = Math.min(2.5, 6.8 - summaryY);

      if (summaryY < 6.0) {
        slide.addShape('roundRect', { x: startX, y: summaryY, w: tableW, h: summaryH, rectRadius: 0.08, fill: { color: C.cardBgAlt || 'E8F5E9' }, line: { color: C.primary, width: 1.5 } });
        if (summaryTitle) {
          slide.addText(summaryTitle, { x: startX + 0.3, y: summaryY + 0.15, w: tableW - 0.6, h: 0.4, align: 'left', color: C.primary, bold: true, fontSize: 16, fontFace: detectFont(summaryTitle, 'title') });
        }
        if (summaryText) {
          let textY = summaryTitle ? summaryY + 0.55 : summaryY + 0.2;
          slide.addText(summaryText, { x: startX + 0.3, y: textY, w: tableW - 0.6, h: summaryH - (summaryTitle ? 0.7 : 0.4), align: 'left', valign: 'top', color: C.cardText || C.textLight, fontSize: 14, fontFace: detectFont(summaryText) });
        }
      }
    }

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §37. Before / After Comparison Slide (NEW)

> Use this when showing a transformation, improvement, or change between two states. Features a red-tinted "Before" column on the left and a green-tinted "After" column on the right with a central divider.

```javascript
function addBeforeAfterSlide(pres, title, beforeTitle, beforeItems, afterTitle, afterItems, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const halfW = 6.16, leftX = 0.3, rightX = 6.87, contentY = 1.4, contentH = 5.5;

  // Before column (red-tinted). Keep the title band inset and rounded so it
  // cannot cover the rounded upper corners of the outer card.
  slide.addShape('roundRect', { x: leftX, y: contentY, w: halfW, h: contentH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: 'E57373', width: 2 } });
  slide.addShape('roundRect', { x: leftX + 0.1, y: contentY + 0.1, w: halfW - 0.2, h: 0.55, rectRadius: 0.06, fill: { color: 'E57373' }, line: { type: 'none' } });
  slide.addText('❌  ' + (beforeTitle || 'Before'), { x: leftX + 0.2, y: contentY + 0.13, w: halfW - 0.4, h: 0.49, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: 'FFFFFF', fontFace: detectFont(beforeTitle, 'title') });

  const beforeBullets = beforeItems.map(b => ({ text: '• ' + b + '\n', options: { fontSize: 16, color: C.cardText, fontFace: detectFont(b), breakLine: true } }));
  slide.addText(beforeBullets, { x: leftX + 0.3, y: contentY + 0.8, w: halfW - 0.6, h: contentH - 1.0, valign: 'top', lineSpacingMultiple: 1.4 });

  // Divider
  slide.addShape('rect', { x: 6.56, y: contentY + 0.2, w: 0.21, h: contentH - 0.4, fill: { color: C.accent }, line: { type: 'none' } });

  // After column (green-tinted). Apply the same inset rounded header treatment.
  slide.addShape('roundRect', { x: rightX, y: contentY, w: halfW, h: contentH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: '66BB6A', width: 2 } });
  slide.addShape('roundRect', { x: rightX + 0.1, y: contentY + 0.1, w: halfW - 0.2, h: 0.55, rectRadius: 0.06, fill: { color: '66BB6A' }, line: { type: 'none' } });
  slide.addText('✅  ' + (afterTitle || 'After'), { x: rightX + 0.2, y: contentY + 0.13, w: halfW - 0.4, h: 0.49, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: 'FFFFFF', fontFace: detectFont(afterTitle, 'title') });

  const afterBullets = afterItems.map(b => ({ text: '• ' + b + '\n', options: { fontSize: 16, color: C.cardText, fontFace: detectFont(b), breakLine: true } }));
  slide.addText(afterBullets, { x: rightX + 0.3, y: contentY + 0.8, w: halfW - 0.6, h: contentH - 1.0, valign: 'top', lineSpacingMultiple: 1.4 });

  return slide;
}
```

---

## §38. Numbered Step Cards (NEW)

> Use for tutorials, how-to guides, or procedural content where each step needs a number badge, title, and description in its own card.

```javascript
function addNumberedStepCards(pres, title, steps, C) {
  const MAX = 4;
  const chunks = chunkArray(steps, MAX);
  const slides = [];

  chunks.forEach((chunk, ci) => {
    const slideTitle = ci === 0 ? title : `${title} (cont.)`;
    let slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, slideTitle, C);

    const n = Math.min(steps.length, MAX), gap = 0.3;
    const cardW = (12.33 - gap * (n - 1)) / n;
    const cardH = 4.5, startY = 1.6;

    chunk.forEach((step, i) => {
      const globalIdx = ci * MAX + i;
      const x = 0.5 + i * (cardW + gap);

      // Card body
      slide.addShape('roundRect', { x, y: startY, w: cardW, h: cardH, rectRadius: 0.1, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 6, offset: 2 } });

      // Large number badge
      slide.addShape('ellipse', { x: x + cardW / 2 - 0.4, y: startY + 0.3, w: 0.8, h: 0.8, fill: { color: C.primary }, line: { type: 'none' }, shadow: { type: 'outer', color: C.primary, opacity: 0.2, blur: 6, offset: 0 } });
      slide.addText(String(globalIdx + 1), { x: x + cardW / 2 - 0.4, y: startY + 0.3, w: 0.8, h: 0.8, align: 'center', valign: 'middle', fontSize: 28, bold: true, color: C.textLight, fontFace: FONTS.englishTitle });

      // Step title
      const cleanTitle = step.title.replace(/^[\d০-৯]+[\.\)]\s*/, '');
      slide.addText(cleanTitle, { x: x + 0.15, y: startY + 1.3, w: cardW - 0.3, h: 0.6, align: 'center', fontSize: 16, bold: true, color: C.primary, fontFace: detectFont(cleanTitle, 'title') });

      // Step description
      slide.addText(step.desc, { x: x + 0.15, y: startY + 2.0, w: cardW - 0.3, h: cardH - 2.3, align: 'center', valign: 'top', fontSize: 14, color: C.cardText, fontFace: detectFont(step.desc), lineSpacingMultiple: 1.3 });

      // Connecting dashed line to next card
      if (i < n - 1) {
        slide.addShape('line', { x: x + cardW + 0.02, y: startY + cardH / 2, w: gap - 0.04, h: 0, line: { color: C.accent, width: 2, dashType: 'dash' } });
      }
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## §39. SWOT Analysis Grid (NEW)

> Use for SWOT analysis or any 2×2 matrix layout with four color-coded quadrants.

```javascript
function addSwotGrid(pres, title, quadrants, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const gridX = 0.5, gridY = 1.3, gridW = 12.33, gridH = 5.8;
  const halfW = gridW / 2 - 0.1, halfH = gridH / 2 - 0.1;
  const headerH = 0.6;

  const labels = quadrants.map(q => q.title) || ['Strengths', 'Weaknesses', 'Opportunities', 'Threats'];
  const colors = [C.primary, C.accent, C.headerBar, C.primary];
  const lightColors = [C.bgAlt, C.cardBgAlt, C.bgAlt, C.cardBgAlt];
  const positions = [
    { x: gridX, y: gridY },
    { x: gridX + halfW + 0.2, y: gridY },
    { x: gridX, y: gridY + halfH + 0.2 },
    { x: gridX + halfW + 0.2, y: gridY + halfH + 0.2 }
  ];

  quadrants.slice(0, 4).forEach((q, i) => {
    const pos = positions[i];

    // Quadrant background
    slide.addShape('roundRect', { x: pos.x, y: pos.y, w: halfW, h: halfH, rectRadius: 0.08, fill: { color: lightColors[i % lightColors.length] }, line: { color: colors[i % colors.length], width: 1.5 } });

    // Inset rounded header preserves the card's rounded outer corners.
    slide.addShape('roundRect', { x: pos.x + 0.08, y: pos.y + 0.08, w: halfW - 0.16, h: headerH - 0.08, rectRadius: 0.06, fill: { color: colors[i % colors.length] }, line: { type: 'none' } });

    const emoji = q.emoji || ['💪', '⚠️', '🌟', '🔥'][i] || '';
    slide.addText(emoji + '  ' + (q.title || labels[i]), { x: pos.x + 0.16, y: pos.y + 0.11, w: halfW - 0.32, h: headerH - 0.14, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.textLight, fontFace: detectFont(q.title, 'title'), fit: 'shrink', margin: 0 });

    // Bullet items
    const bullets = (q.items || []).map(item => '• ' + item).join('\n');
    slide.addText(bullets, { x: pos.x + 0.24, y: pos.y + headerH + 0.16, w: halfW - 0.48, h: halfH - headerH - 0.3, fontSize: 14, color: C.cardText, valign: 'top', fontFace: detectFont(bullets), lineSpacingMultiple: 1.35 });
  });

  return slide;
}
```

---

## §40. Two-Column Text Layout (NEW)

> Use when presenting two parallel blocks of written text side-by-side, each with its own title. Perfect for compare/contrast essays, bilingual content, or dual perspectives.

```javascript
function addTwoColumnText(pres, title, leftTitle, leftBody, rightTitle, rightBody, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addHeaderBar(slide, title, C);

  const colW = 5.9, colH = 5.3, gap = 0.35;
  const leftX = 0.5, rightX = leftX + colW + gap;
  const startY = 1.4, titleH = 0.6;

  // Left column
  slide.addShape('roundRect', { x: leftX, y: startY, w: colW, h: colH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.primary, width: 1.5 } });
  slide.addShape('roundRect', { x: leftX + 0.1, y: startY + 0.1, w: colW - 0.2, h: titleH - 0.05, rectRadius: 0.06, fill: { color: C.primary }, line: { type: 'none' } });
  slide.addText(leftTitle, { x: leftX + 0.2, y: startY + 0.13, w: colW - 0.4, h: titleH - 0.12, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.textLight, fontFace: detectFont(leftTitle, 'title'), fit: 'shrink', margin: 0 });
  slide.addText(leftBody, { x: leftX + 0.25, y: startY + titleH + 0.2, w: colW - 0.5, h: colH - titleH - 0.35, fontSize: 15, color: C.cardText, fontFace: detectFont(leftBody), valign: 'top', lineSpacingMultiple: 1.35 });

  // Right column
  slide.addShape('roundRect', { x: rightX, y: startY, w: colW, h: colH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.accent, width: 1.5 } });
  slide.addShape('roundRect', { x: rightX + 0.1, y: startY + 0.1, w: colW - 0.2, h: titleH - 0.05, rectRadius: 0.06, fill: { color: C.accent }, line: { type: 'none' } });
  slide.addText(rightTitle, { x: rightX + 0.2, y: startY + 0.13, w: colW - 0.4, h: titleH - 0.12, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.textLight, fontFace: detectFont(rightTitle, 'title'), fit: 'shrink', margin: 0 });
  slide.addText(rightBody, { x: rightX + 0.25, y: startY + titleH + 0.2, w: colW - 0.5, h: colH - titleH - 0.35, fontSize: 15, color: C.cardText, fontFace: detectFont(rightBody), valign: 'top', lineSpacingMultiple: 1.35 });

  return slide;
}
```

## 41. Executive Summary (3 KPIs + Introduction + Objectives + Quote)

Use this function with `data` in this shape: `{ kpis: [{ label, value, desc }, ...], intro: "...", objectives: ["...", ...], quote: "...", quoteAuthor: "..." }`. Pass all objectives in the one `objectives` array; the layout creates continuation slides automatically when needed.

```javascript
function addExecutiveSummarySlide(pres, title, data, C) {
  const kpis = Array.isArray(data.kpis) ? data.kpis : [];
  const objectives = Array.isArray(data.objectives) ? data.objectives : [];
  const objectivePages = chunkArray(objectives, 6);
  const pages = objectivePages.length ? objectivePages : [[]];
  const slides = [];

  pages.forEach((objectivePage, pageIndex) => {
    const slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, pageIndex === 0 ? title : `${title} (cont.)`, C);

    if (pageIndex === 0) {
      const intro = data.intro || '';
      slide.addShape('roundRect', { x: 0.5, y: 1.05, w: 12.33, h: 0.58, rectRadius: 0.06, fill: { color: C.bgAlt }, line: { color: C.border, width: 1 } });
      slide.addText(intro, { x: 0.72, y: 1.16, w: 11.89, h: 0.32, fontSize: 12, color: C.textBody, fontFace: detectFont(intro), valign: 'mid', fit: 'shrink', margin: 0 });

      // The first row is reserved for exactly three executive KPIs.
      const shownKpis = kpis.slice(0, 3), gap = 0.22, cardW = (12.33 - gap * 2) / 3;
      shownKpis.forEach((kpi, i) => {
        const x = 0.5 + i * (cardW + gap);
        slide.addShape('roundRect', { x, y: 1.86, w: cardW, h: 1.35, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', color: C.shadow, opacity: 0.1, blur: 4, offset: 1 } });
        slide.addShape('rect', { x, y: 1.86, w: cardW, h: 0.07, fill: { color: i === 1 ? C.accent : C.primary }, line: { type: 'none' } });
        slide.addText(kpi.label || '', { x: x + 0.16, y: 2.05, w: cardW - 0.32, h: 0.24, fontSize: 11, bold: true, color: C.primary, fontFace: detectFont(kpi.label), fit: 'shrink', margin: 0 });
        slide.addText(kpi.value || '', { x: x + 0.16, y: 2.35, w: cardW - 0.32, h: 0.46, fontSize: 24, bold: true, color: C.accent, fontFace: detectFont(kpi.value, 'title'), fit: 'shrink', margin: 0 });
        slide.addText(kpi.desc || '', { x: x + 0.16, y: 2.83, w: cardW - 0.32, h: 0.2, fontSize: 9, color: C.textMuted, fontFace: detectFont(kpi.desc), fit: 'shrink', margin: 0 });
      });

      const panelY = 3.55, panelH = 3.15, objectivesX = 0.5, objectivesW = 7.25;
      slide.addShape('roundRect', { x: objectivesX, y: panelY, w: objectivesW, h: panelH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 } });
      slide.addText('মূল উদ্দেশ্যসমূহ', { x: objectivesX + 0.22, y: panelY + 0.2, w: objectivesW - 0.44, h: 0.3, fontSize: 16, bold: true, color: C.primary, fontFace: FONTS.banglaTitle });
      const objectiveText = objectivePage.map((item) => `• ${typeof item === 'string' ? item : item.text || item.title || ''}`).join('\n');
      slide.addText(objectiveText, { x: objectivesX + 0.25, y: panelY + 0.68, w: objectivesW - 0.5, h: panelH - 0.9, fontSize: 13, color: C.cardText, fontFace: detectFont(objectiveText), fit: 'shrink', margin: 0.02, valign: 'top', lineSpacingMultiple: 1.18 });

      const quoteX = 8.0, quote = data.quote || data.quoteText || '', quoteAuthor = data.quoteAuthor || '';
      slide.addShape('roundRect', { x: quoteX, y: panelY, w: 4.83, h: panelH, rectRadius: 0.08, fill: { color: C.primary }, line: { color: C.accent, width: 1 } });
      slide.addText('“', { x: quoteX + 0.25, y: panelY + 0.15, w: 0.7, h: 0.6, fontSize: 36, bold: true, color: C.accent, fontFace: FONTS.englishTitle, margin: 0 });
      slide.addText(quote, { x: quoteX + 0.38, y: panelY + 0.86, w: 4.07, h: 1.55, fontSize: 16, italic: true, color: C.textLight, fontFace: detectFont(quote), align: 'center', valign: 'middle', fit: 'shrink', margin: 0.03, lineSpacingMultiple: 1.15 });
      // The attribution sits on a dark primary panel, so use the light text
      // token rather than textMuted (which is designed for light surfaces).
      if (quoteAuthor) slide.addText(`— ${quoteAuthor}`, { x: quoteX + 0.35, y: panelY + 2.55, w: 4.13, h: 0.25, fontSize: 10, color: C.textLight, transparency: 18, fontFace: detectFont(quoteAuthor), align: 'center', fit: 'shrink', margin: 0 });
    } else {
      const objectiveText = objectivePage.map((item) => `• ${typeof item === 'string' ? item : item.text || item.title || ''}`).join('\n');
      slide.addShape('roundRect', { x: 0.7, y: 1.35, w: 11.93, h: 5.55, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 } });
      slide.addText('মূল উদ্দেশ্যসমূহ', { x: 1.0, y: 1.7, w: 11.33, h: 0.35, fontSize: 19, bold: true, color: C.primary, fontFace: FONTS.banglaTitle });
      slide.addText(objectiveText, { x: 1.05, y: 2.25, w: 11.23, h: 4.2, fontSize: 17, color: C.cardText, fontFace: detectFont(objectiveText), fit: 'shrink', margin: 0.04, valign: 'top', lineSpacingMultiple: 1.3 });
    }
    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```

---

## 42. Multi-metric Dashboard (Native Charts & KPIs)

```javascript
function addDashboardSlide(pres, title, subtitle, data, C) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };

  const sidebarW = 3.8;
  
  // 1. Sidebar Background
  slide.addShape('rect', { x: 0, y: 0, w: sidebarW, h: 7.5, fill: { color: C.primary }, line: { type: 'none' } });

  // 2. Sidebar Text
  // Keep the title and subtitle in separate, generous zones. A 3.2" sidebar
  // cannot safely hold a 32pt Bangla/English title in only one inch of height.
  // The smaller responsive size plus shrink-to-fit prevents long bilingual
  // titles from colliding with the subtitle.
  const sidebarTitleSize = title && title.length > 34 ? 22 : title && title.length > 22 ? 25 : 28;
  slide.addText(title, { x: 0.3, y: 0.45, w: sidebarW - 0.6, h: 1.45, fontSize: sidebarTitleSize, bold: true, color: C.textLight, fontFace: detectFont(title, 'title'), valign: 'top', breakLine: false, fit: 'shrink', margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.3, y: 2.05, w: sidebarW - 0.6, h: 0.65, fontSize: 12, color: C.textLight, fontFace: detectFont(subtitle), valign: 'top', lineSpacingMultiple: 1.15, fit: 'shrink', margin: 0 });
  }
  slide.addShape('line', { x: 0.3, y: 3.0, w: sidebarW - 0.6, h: 0, line: { color: C.accent, width: 1.5 } });

  // 3. Sidebar KPI Pills
  if (data.sidebarMetrics && Array.isArray(data.sidebarMetrics)) {
    const kpiStartY = 3.3;
    const kpiGap = 1.15;
    data.sidebarMetrics.slice(0, 3).forEach((kpi, i) => {
      const y = kpiStartY + (i * kpiGap);
      // Pill Background
      slide.addShape('roundRect', { x: 0.3, y, w: sidebarW - 0.6, h: 1.0, rectRadius: 0.15, fill: { color: C.cardBg }, line: { type: 'none' }, shadow: { type: 'outer', blur: 6, offset: 2, opacity: 0.15 } });
      
      // Icon Box (Accent color)
      slide.addShape('roundRect', { x: 0.45, y: y + 0.15, w: 0.7, h: 0.7, rectRadius: 0.1, fill: { color: C.primary }, line: { type: 'none' } });
      if (kpi.icon) {
        slide.addText(kpi.icon, { x: 0.45, y: y + 0.15, w: 0.7, h: 0.7, align: 'center', valign: 'middle', fontSize: 22, fontFace: FONTS.englishBody });
      }

      // Value & Label
      slide.addText(kpi.value || '', { x: 1.3, y: y + 0.15, w: sidebarW - 1.8, h: 0.4, fontSize: 22, bold: true, color: C.primary, fontFace: detectFont(kpi.value, 'title'), valign: 'bottom' });
      slide.addText(kpi.label || '', { x: 1.3, y: y + 0.55, w: sidebarW - 1.8, h: 0.3, fontSize: 12, color: C.cardText, fontFace: detectFont(kpi.label), valign: 'top' });
    });
  }

  // 4. Main Area - Top Metrics
  const mainX = sidebarW + 0.3;
  const mainW = 13.33 - sidebarW - 0.6;
  const textColor = isLightBg(C) ? (C.textDark || '333333') : C.textLight;
  
  if (data.topMetrics && Array.isArray(data.topMetrics)) {
    const count = Math.min(data.topMetrics.length, 4);
    const gap = 0.2;
    const cardW = (mainW - (gap * (count - 1))) / count;
    const cardH = 1.3;
    
    data.topMetrics.slice(0, 4).forEach((metric, i) => {
      const x = mainX + (i * (cardW + gap));
      slide.addShape('roundRect', { x, y: 0.4, w: cardW, h: cardH, rectRadius: 0.1, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', blur: 5, offset: 2, opacity: 0.1 } });
      
      slide.addShape('rect', { x, y: 0.4, w: cardW, h: 0.08, fill: { color: i % 2 === 0 ? C.primary : C.accent }, line: { type: 'none' } });

      slide.addText(metric.label || '', { x: x + 0.1, y: 0.55, w: cardW - 0.2, h: 0.3, fontSize: 13, bold: true, color: C.primary, fontFace: detectFont(metric.label) });
      slide.addText(metric.desc || '', { x: x + 0.1, y: 0.85, w: cardW - 0.2, h: 0.3, fontSize: 10, color: C.cardText, fontFace: detectFont(metric.desc) });
      slide.addText(metric.value || '', { x: x + 0.1, y: 1.2, w: cardW - 0.2, h: 0.4, fontSize: 26, bold: true, color: C.accent, fontFace: detectFont(metric.value, 'title') });
    });
  }

  // 5. Main Area - Charts
  if (data.charts && Array.isArray(data.charts)) {
    const count = Math.min(data.charts.length, 2);
    const chartStartY = 1.9;
    const chartCardH = (7.5 - chartStartY - 0.4 - (0.3 * (count - 1))) / count;
    
    data.charts.slice(0, 2).forEach((chartObj, i) => {
      const y = chartStartY + (i * (chartCardH + 0.3));
      
      // Chart Card Background
      slide.addShape('roundRect', { x: mainX, y, w: mainW, h: chartCardH, rectRadius: 0.08, fill: { color: C.cardBg }, line: { color: C.border, width: 1 }, shadow: { type: 'outer', blur: 6, offset: 2, opacity: 0.1 } });
      
      // Chart Title
      slide.addText(chartObj.title || 'Chart', { x: mainX + 0.2, y: y + 0.1, w: mainW - 0.4, h: 0.4, fontSize: 16, bold: true, color: C.primary, fontFace: detectFont(chartObj.title, 'title') });
      slide.addShape('line', { x: mainX + 0.2, y: y + 0.5, w: mainW - 0.4, h: 0, line: { color: C.border, width: 1 } });
      
      // Chart Rendering
      if (chartObj.data && chartObj.type) {
        let chartType = pres.ChartType.bar;
        if (chartObj.type === 'doughnut' || chartObj.type === 'pie') chartType = pres.ChartType.doughnut;
        else if (chartObj.type === 'line') chartType = pres.ChartType.line;

        let opts = {
          x: mainX + 0.2, y: y + 0.6, w: mainW - 0.4, h: chartCardH - 0.7,
          showLegend: true, legendPos: 'r', legendColor: textColor, legendFontSize: 11,
          chartColors: [C.primary, C.accent, '#4CAF50', '#FFC107', '#F44336', '#9C27B0'],
          dataLabelColor: textColor, dataLabelFontSize: 10,
          valAxisLineShow: false, catAxisLineShow: true, catAxisLineColor: C.border,
          valAxisLabelColor: textColor, catAxisLabelColor: textColor,
          valGridLine: { color: C.border, style: 'dash' }
        };

        if (chartType === pres.ChartType.doughnut) {
           opts.holeSize = 65;
           opts.showValue = true;
           opts.dataLabelPosition = 'outEnd';
           opts.showTitle = false;
        } else {
           opts.barDir = 'col';
           opts.barGrouping = 'clustered';
        }

        try {
          slide.addChart(chartType, chartObj.data, opts);
        } catch(e) {
          slide.addText('Error rendering chart: ' + e.message, { x: mainX+0.5, y: y+1, w: mainW-1, h: 1, color: 'FF0000' });
        }
      }
    });
  }

  return slide;
}
```

---

## 43. Disease Signs, Symptoms & Prevention Cards

> Use this as the default clinical, veterinary, agriculture, or public-health layout whenever every condition needs **all four** of these elements: a condition title, a real image, signs/symptoms bullets, and prevention bullets. The visual language intentionally follows an awareness poster: a bold color title band, an image-led upper section, and two strongly labelled bullet areas below. Do not use this layout for generic feature cards.

Pass every condition in one array—never pre-split it. Each item uses this shape:

```javascript
const diseaseCardExample = {
  title: 'Newcastle Disease (ND)',
  imagePath: '/absolute/path/to/newcastle-disease-chicken.jpg',
  signs: ['Sudden death', 'Coughing or sneezing', 'Green diarrhea'],
  prevention: ['Vaccinate the flock', 'Maintain strict biosecurity', 'Avoid contact with wild birds']
};
```

```javascript
function addDiseaseSymptomsPreventionCards(pres, title, diseases, C) {
  const pages = chunkArray(Array.isArray(diseases) ? diseases : [], 3);
  const safePages = pages.length ? pages : [[]];
  const slides = [];

  safePages.forEach((page, pageIndex) => {
    const slide = pres.addSlide();
    slide.background = { color: C.bg };
    addHeaderBar(slide, pageIndex === 0 ? title : `${title} (cont.)`, C);
    addCornerAccents(slide, C);

    // Three balanced poster cards create an infographic-like disease overview.
    const cardY = 1.06, cardW = 3.96, cardH = 6.04, gap = 0.225;
    const bodyColor = C.cardText || C.textDark || '222222';

    page.forEach((disease, i) => {
      if (!disease.imagePath) {
        throw new Error(`addDiseaseSymptomsPreventionCards requires imagePath for "${disease.title || 'each condition'}".`);
      }

      const x = 0.5 + i * (cardW + gap);
      const bandColor = i === 0 ? C.primary : C.accent;
      const labelColor = i === 0 ? C.primary : C.accent;
      const signs = (disease.signs || disease.symptoms || []).map(item => `• ${item}`).join('\n');
      const prevention = (disease.prevention || []).map(item => `• ${item}`).join('\n');

      // One self-contained poster card keeps the image and both bullet groups inseparable.
      slide.addShape('roundRect', {
        x, y: cardY, w: cardW, h: cardH, rectRadius: 0.10,
        fill: { color: C.cardBg }, line: { color: C.border, width: 1 },
        shadow: { type: 'outer', color: C.shadow || '777777', opacity: 0.16, blur: 4, offset: 1 }
      });
      slide.addShape('roundRect', {
        x, y: cardY, w: cardW, h: 0.58, rectRadius: 0.10,
        fill: { color: bandColor }, line: { type: 'none' }
      });
      // Square the lower edge of the title band so it reads as a clean header strip.
      slide.addShape('rect', { x, y: cardY + 0.40, w: cardW, h: 0.18, fill: { color: bandColor }, line: { type: 'none' } });
      slide.addText(disease.title || 'Condition', {
        x: x + 0.16, y: cardY + 0.12, w: cardW - 0.32, h: 0.28,
        fontSize: 13, bold: true, color: C.textLight, align: 'center', valign: 'middle',
        fontFace: detectFont(disease.title || 'Condition', 'title'), fit: 'shrink', margin: 0
      });

      // The photo frame is deliberately 16:9 so landscape clinical images retain their intended composition.
      const imageW = cardW - 0.30, imageH = imageW * 9 / 16;
      slide.addShape('roundRect', { x: x + 0.13, y: cardY + 0.71, w: cardW - 0.26, h: imageH + 0.04, rectRadius: 0.05, fill: { color: C.bgAlt || C.cardBg }, line: { color: bandColor, width: 1.1 } });
      slide.addImage({ path: disease.imagePath, x: x + 0.15, y: cardY + 0.73, w: imageW, h: imageH, sizing: { type: 'cover', w: imageW, h: imageH } });
      slide.addShape('ellipse', { x: x + 0.24, y: cardY + 0.80, w: 0.42, h: 0.42, fill: { color: C.cardBg }, line: { color: bandColor, width: 1.1 } });
      slide.addText(String(pageIndex * 3 + i + 1), { x: x + 0.24, y: cardY + 0.87, w: 0.42, h: 0.15, align: 'center', fontSize: 9, bold: true, color: bandColor, fontFace: FONTS.englishTitle, margin: 0 });
      slide.addShape('line', { x: x + 0.20, y: cardY + 2.98, w: cardW - 0.40, h: 0, line: { color: C.border, width: 0.8 } });

      slide.addShape('rect', { x: x + 0.22, y: cardY + 3.17, w: 0.06, h: 0.26, fill: { color: labelColor }, line: { type: 'none' } });
      slide.addText('SIGNS & SYMPTOMS', {
        x: x + 0.36, y: cardY + 3.17, w: cardW - 0.58, h: 0.22,
        fontSize: 9.5, bold: true, color: labelColor, fontFace: FONTS.englishTitle, margin: 0
      });
      slide.addText(signs, {
        x: x + 0.30, y: cardY + 3.52, w: cardW - 0.56, h: 0.83,
        fontSize: 9.3, color: bodyColor, fontFace: detectFont(signs), valign: 'top',
        lineSpacingMultiple: 1.02, fit: 'shrink', margin: 0
      });

      slide.addShape('line', { x: x + 0.20, y: cardY + 4.52, w: cardW - 0.40, h: 0, line: { color: C.border, width: 0.65 } });
      slide.addShape('rect', { x: x + 0.22, y: cardY + 4.73, w: 0.06, h: 0.26, fill: { color: C.accent }, line: { type: 'none' } });
      slide.addText('PREVENTION', {
        x: x + 0.36, y: cardY + 4.73, w: cardW - 0.58, h: 0.22,
        fontSize: 9.5, bold: true, color: C.accent, fontFace: FONTS.englishTitle, margin: 0
      });
      slide.addText(prevention, {
        x: x + 0.30, y: cardY + 5.08, w: cardW - 0.56, h: 0.69,
        fontSize: 9.3, color: bodyColor, fontFace: detectFont(prevention), valign: 'top',
        lineSpacingMultiple: 1.02, fit: 'shrink', margin: 0
      });
    });

    slides.push(slide);
  });

  return slides.length === 1 ? slides[0] : slides;
}
```
