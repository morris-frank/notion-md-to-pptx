# Implementation spec: Notion-style markdown to PowerPoint (single-file script)

## 1. Scope

Build a single-file CLI script that converts a constrained Notion-style markdown dialect into a `.pptx` deck.

Target workflow:

- author slides in one markdown file
- run one command
- receive a PowerPoint deck with deterministic layout
- drive appearance from a theme file rather than hard-coded style

Out of scope for v1:

- full Markdown/CommonMark fidelity
- arbitrary HTML support beyond the specified `<aside>` and `<img>` patterns
- editable native Office equations
- smart multi-column layout inference
- round-tripping PPTX back into markdown
- speaker notes

## 2. Recommended implementation choice

### Chosen stack

Use a Node.js single-file CLI script with:

- `pptxgenjs` for PowerPoint generation
- `markdown-it` for markdown tokenisation
- `mathjax` for TeX/LaTeX to SVG conversion
- native `fetch` / `https` for remote asset download
- a JSON theme file

### Why this is the best fit

1. PPTX generation
  `PptxGenJS` is explicitly built to generate PowerPoint presentations and supports text, images, shapes, tables, and general OOXML-compatible `.pptx` output. ([github.com](https://github.com/gitbrent/PptxGenJS?utm_source=chatgpt.com))
2. Math rendering
  `MathJax` supports server-side conversion from TeX input to SVG output, which is the most robust route for equation fidelity inside PowerPoint. SVG output is high-quality and resolution-independent. ([docs.mathjax.org](https://docs.mathjax.org/en/v4.0/web/convert.html?utm_source=chatgpt.com))
3. Better fit than Python for this specific task
  `python-pptx` is fully viable for creating and updating PowerPoint files, but for this use case the Node ecosystem gives a cleaner path for markdown parsing and MathJax integration in one script. Python remains a fallback, not the preferred choice. ([python-pptx.readthedocs.io](https://python-pptx.readthedocs.io/?utm_source=chatgpt.com))

## 3. Deliverable

### CLI shape

```bash
node notion-md-to-pptx.mjs input.md --theme theme.json --out output.pptx
```

### Inputs

- `input.md`: markdown source file
- `theme.json`: theme contract
- optional local asset cache directory

### Output

- `output.pptx`
- optional `.cache/` folder for downloaded icons and rendered equation SVGs

## 4. Authoring contract

### 4.1 Slide boundaries

Rules:

- first `#`  heading is always the title slide
- title slide contains title only
- `---` starts a new slide
- all subsequent slides should normally start with `##`  as slide title

Example:

```md
# Deck title

---
## Slide A

Body

---
## Slide B
```

### 4.2 Supported constructs

Supported in v1:

- `# H1` only for deck title
- `## H2` for slide title
- `### H3` as emphasized subheading within slide body
- paragraphs
- bullet lists
- strong / emphasis
- inline code
- fenced code blocks
- inline LaTeX
- block LaTeX
- `<aside>...</aside>` blocks
- `<img src="..." ... />` inside aside blocks

Unsupported or downgraded:

- nested blockquotes
- tables
- arbitrary inline HTML
- embedded iframes
- arbitrary image layout outside defined rules

## 5. Parsing model

Implement as a two-stage pipeline.

### Stage A. Parse source into AST

Create an internal deck AST independent of PowerPoint.

```ts
interface Deck {
  title: string
  slides: SlideSpec[]
}

interface SlideSpec {
  title: string
  blocks: Block[]
}

type Block =
  | ParagraphBlock
  | BulletListBlock
  | CodeBlock
  | MathBlock
  | CalloutBlock
  | SubheadingBlock

interface ParagraphBlock {
  kind: 'paragraph'
  inlines: Inline[]
}

interface BulletListBlock {
  kind: 'bullets'
  items: Inline[][]
}

interface CodeBlock {
  kind: 'code-block'
  text: string
  language?: string
}

interface MathBlock {
  kind: 'math-block'
  tex: string
}

interface CalloutBlock {
  kind: 'callout'
  icon?: IconSpec
  blocks: Block[]
}

interface SubheadingBlock {
  kind: 'subheading'
  inlines: Inline[]
}

type Inline =
  | { kind: 'text'; text: string }
  | { kind: 'strong'; text: string }
  | { kind: 'em'; text: string }
  | { kind: 'code'; text: string }
  | { kind: 'math-inline'; tex: string }
```

### Stage B. Render AST into PPTX

Rendering must be deterministic and layout-driven by theme + content fitting rules.

## 6. Slide layout model

Use a single-content-column master layout for v1.

### 6.1 Base regions

For 16:9 slides:

- slide size: standard wide
- outer margin: fixed
- title area: fixed Y band
- content area: remaining box under title

Suggested default geometry:

- slide: 13.333 x 7.5 in
- left/right margin: 0.6 in
- top margin: 0.45 in
- title box height: 0.8 in
- content top gap after title: 0.18 in
- bottom margin: 0.4 in

### 6.2 Reading order

Render blocks top-to-bottom in source order.

### 6.3 Overflow policy

Default style is fixed. Only scale down when required.

Algorithm:

1. render against default token sizes
2. estimate total required height
3. if content fits, emit
4. else reduce body scale in steps, e.g. `1.00 -> 0.94 -> 0.89 -> 0.84`
5. preserve title size unless impossible
6. if still does not fit at `min_scale`, fail with explicit overflow error and slide index

This is preferable to silently clipping content.

## 7. Typography and fitting

### 7.1 Defaults

Per theme tokens, not hard-coded in script.

Required token classes:

- title
- subtitle or subheading
- body
- bullet
- inline code
- code block
- callout title/body
- math inline/block sizing multipliers

### 7.2 Scaling policy

Scale only these on overflow:

- paragraph body
- bullets
- callout body
- code block text
- subheadings optionally by smaller ratio

Do not scale below configured floor.

Suggested floors:

- title: 36 pt
- body: 18 pt
- code: 16 pt

### 7.3 Alignment rules

Default:

- slide title: left aligned
- body text: left aligned
- bullets: left aligned
- code blocks: left aligned
- block equations: centred by default
- callout content: left aligned

## 8. `<aside>` mapping to PowerPoint callouts

### 8.1 Authoring pattern

Supported pattern:

```html
<aside>
<img src="https://www.notion.so/icons/checkmark-square_orange.svg" alt="..." width="40px" />

### Callout heading

Body text
</aside>
```

### 8.2 Rendered form

Map `<aside>` to a rounded rectangle or theme-defined box with:

- background fill
- border
- internal padding
- optional icon on left
- text stacked on right or below depending on width

Recommended v1 layout:

- icon column on left
- content column on right
- minimum box height determined by icon or content, whichever is larger

### 8.3 Emoji and icon mapping

Rules:

1. if `<img src>` points to a downloadable SVG/PNG, download once and cache locally
2. preserve local cached copy for deterministic builds
3. if source is SVG, prefer using SVG directly if supported by renderer path; otherwise rasterise once into PNG
4. if no explicit image exists and callout begins with an emoji token, map using configurable emoji-to-asset table
5. if no mapping exists, render emoji as text glyph

### 8.4 Asset cache

Suggested cache scheme:

```text
.cache/
  assets/
    sha256_<url>.<ext>
  math/
    sha256_<tex>.svg
```

## 9. Inline text rendering rules

### 9.1 Mixed runs

Paragraphs and bullets should render as rich text runs where possible.

Handle:

- plain text
- bold
- italic
- inline code
- inline math as inline image insert or run surrogate

### 9.2 Inline code

Render inline code as:

- monospace font from theme
- tinted fill or subtle highlight background if feasible
- no syntax highlighting in v1

If run-level background is too awkward in chosen PPTX library, acceptable fallback is monospace + colour only.

### 9.3 Inline math

Render inline math by converting TeX fragment to SVG and inserting it inline-equivalent.

Boundary decision:

- do not attempt native editable Office equations in v1
- render math as image assets for fidelity and portability

## 10. Block code and block math

### 10.1 Code blocks

Render fenced code blocks inside a code container:

- slightly tinted background
- optional border radius approximation
- monospace font
- preserve whitespace
- wrap long lines only if configured; otherwise shrink or overflow error

### 10.2 Block math

Convert block TeX to SVG with MathJax and insert as centred image with theme-controlled max width.

## 11. Theme contract

Use JSON for v1.

Reason:

- trivial to parse in a single-file script
- no dependency on PowerPoint desktop tooling
- portable in CI
- easier to validate explicitly than a `.pptx` template

A `.pptx` template is still useful later for corporate masters, but for v1 JSON is the cleaner source of truth.

### 11.1 Required theme fields

```json
{
  "meta": {
    "name": "Soilytix Light",
    "version": 1
  },
  "slide": {
    "layout": "LAYOUT_WIDE",
    "background": "#F5F2EF"
  },
  "colors": {
    "dk1": "#29322E",
    "dk2": "#423B2D",
    "lt1": "#FEFFFF",
    "lt2": "#F5F2EF",
    "accent1": "#87EA21",
    "accent2": "#B1813E",
    "accent3": "#3E4641",
    "accent4": "#212121",
    "accent5": "#A02B93",
    "accent6": "#4EA72E",
    "hlink": "#467886",
    "folHlink": "#96607D"
  },
  "fonts": {
    "title": "Inter",
    "body": "Inter",
    "mono": "Aptos Mono",
    "fallbackSans": "Aptos",
    "fallbackDisplay": "Aptos Display"
  },
  "typeScale": {
    "titlePt": 44,
    "subheadingPt": 24,
    "bodyPt": 28,
    "bulletPt": 28,
    "inlineCodePt": 24,
    "codeBlockPt": 20,
    "calloutHeadingPt": 24,
    "calloutBodyPt": 24,
    "minBodyPt": 18,
    "minCodePt": 16
  },
  "layout": {
    "marginLeft": 0.6,
    "marginRight": 0.6,
    "marginTop": 0.45,
    "marginBottom": 0.4,
    "titleHeight": 0.8,
    "contentGapAfterTitle": 0.18,
    "paragraphGap": 0.14,
    "bulletGap": 0.08,
    "calloutGap": 0.18,
    "calloutPadding": 0.16,
    "calloutIconSize": 0.32,
    "calloutIconGap": 0.14,
    "codePadding": 0.14,
    "mathBlockGap": 0.14
  },
  "text": {
    "titleColor": "#29322E",
    "bodyColor": "#29322E",
    "mutedColor": "#423B2D",
    "inlineCodeColor": "#212121",
    "linkColor": "#467886"
  },
  "blocks": {
    "callout": {
      "fill": "#FEFFFF",
      "border": "#B1813E",
      "borderWidth": 1.5,
      "radius": 0.12
    },
    "code": {
      "fill": "#FEFFFF",
      "border": "#3E4641",
      "borderWidth": 1,
      "radius": 0.08
    },
    "inlineCode": {
      "fill": "#EDE8E2"
    }
  },
  "fit": {
    "scales": [1.0, 0.94, 0.89, 0.84],
    "minScale": 0.84
  },
  "emojiMap": {
    "✅": "./assets/checkmark-square_orange.svg"
  }
}
```

## 12. Rendering rules by block type

### 12.1 Title slide

Input:

- first H1 only

Render:

- one centred or left-aligned title textbox according to theme
- no footer, no body, no subtitle in v1

### 12.2 Standard slide title

Input:

- first H2 after slide break

Render:

- title textbox at top using title style

### 12.3 Paragraph

Render as rich-text runs in one textbox with calculated height.

### 12.4 Bullet list

Render as one textbox with paragraph bullets or one paragraph per bullet item.

For v1, only one nesting level is required.

### 12.5 Subheading

Render as semantically distinct text block, visually between title and body.

### 12.6 Callout

Render as shape + child content.

### 12.7 Code block

Render inside code panel.

### 12.8 Math block

Render as SVG image.

## 13. Height estimation strategy

A precise PowerPoint layout engine is not available directly. Use pragmatic estimation.

### 13.1 Estimation inputs

- font size
- text length
- average character width factor by font class
- line height multiplier
- box width
- bullet indent

### 13.2 Strategy

For each block type, compute estimated height before emitting.

Suggested line-height multipliers:

- title: 1.1
- body: 1.2
- code: 1.15
- callout body: 1.2

This does not need pixel-perfect fidelity; it needs consistent overestimation.

### 13.3 Validation mode

Provide `--debug-layout` to render invisible guide boxes or log per-slide height budget vs estimated height.

## 14. Error handling

Fail explicitly for:

- missing title H1
- empty slide after `---`
- unsupported nested structures inside `<aside>`
- non-downloadable or failed remote image when no fallback exists
- malformed LaTeX block if MathJax cannot convert it
- content overflow after minimum scale reached

Warnings, not failures:

- unknown emoji mapping, falling back to text
- unsupported markdown construct dropped
- remote asset downloaded with content type mismatch but usable bytes

## 15. Determinism and reproducibility

Required:

- stable slide ordering
- asset cache by content hash or URL hash
- math cache by TeX hash
- optional `--offline` mode using cache only
- no non-deterministic random placements

## 16. Suggested file structure

Even if implemented as one script, organise internally into sections:

- CLI parsing
- theme load/validate
- markdown preprocessing
- AST parse
- asset resolution
- math rendering
- layout estimation
- PPTX rendering
- write file

## 17. Single-file script architecture

In one `.mjs` file, structure as pure functions:

```ts
main()
loadTheme()
readMarkdown()
splitSlides()
parseSlide()
parseAside()
resolveAssets()
renderMathSvg()
layoutSlide()
renderDeck()
renderSlide()
renderBlock()
writePptx()
```

Keep global mutable state limited to:

- asset cache index
- math cache index
- pptx document instance

## 18. Example source -> interpretation

Given source:

```md
## Material

<aside>
<img src="https://www.notion.so/icons/checkmark-square_orange.svg" alt="https://www.notion.so/icons/checkmark-square_orange.svg" width="40px" />

### Material provenance is not assay execution

</aside>

**Final boundary**

- `SampleLineage` / provenance is the durable scientific root for material identity.
- Assay orders, analysis selections, and execution attempts operate on material; they do not define it.
```

Interpretation:

- slide title: `Material`
- callout with icon + subheading text
- strong paragraph `Final boundary`
- bullet list with two items, each containing inline code runs

## 19. Boundary decisions and justifications

### Decision 1. Use JSON theme, not `.pptx` template, in v1

Justification:

- easier validation
- easier diffing in Git
- no dependency on pre-authored Office binary assets
- simpler single-file implementation

### Decision 2. Equations render as SVG images, not native PowerPoint equations

Justification:

- higher fidelity
- fewer compatibility problems
- deterministic output across platforms
- much simpler implementation

### Decision 3. Slide fitting scales down within bounded steps, then fails

Justification:

- avoids silent clipping
- avoids illegible text collapse
- gives deterministic reviewable failure cases

### Decision 4. Only constrained Notion subset is supported

Justification:

- makes the script reliable
- matches actual authoring intent
- prevents uncontrolled layout complexity

### Decision 5. Download callout icons into local cache

Justification:

- reproducible builds
- offline rebuild support
- avoids dead-link regressions

## 20. Minimal acceptance criteria

The implementation is acceptable when all of the following hold:

1. A markdown file with `#` title + multiple `---` slides produces a valid `.pptx`.
2. Title slide contains only the title.
3. Standard slides render H2 title and body blocks in order.
4. `<aside>` blocks render as boxed callouts.
5. Remote Notion icon assets are downloaded once and reused locally.
6. Inline code is visually distinct from body text.
7. Inline and block LaTeX render successfully via SVG in PowerPoint.
8. Overflow triggers bounded scale-down and then explicit failure if still too large.
9. Theme changes via JSON affect output without code edits.
10. Running twice on same input and cache yields equivalent output structure.

## 21. Recommended next implementation order

1. parse slide boundaries and titles
2. render plain title/body/bullets
3. add theme loader
4. add callout parser and shape rendering
5. add inline code handling
6. add MathJax SVG rendering and caching
7. add height estimation and scale-down logic
8. add diagnostics and offline mode

## 22. Example starter theme for the provided Soilytix palette

```json
{
  "meta": { "name": "Soilytix Light", "version": 1 },
  "slide": { "layout": "LAYOUT_WIDE", "background": "#F5F2EF" },
  "colors": {
    "dk1": "#29322E",
    "dk2": "#423B2D",
    "lt1": "#FEFFFF",
    "lt2": "#F5F2EF",
    "accent1": "#87EA21",
    "accent2": "#B1813E",
    "accent3": "#3E4641",
    "accent4": "#212121",
    "accent5": "#A02B93",
    "accent6": "#4EA72E",
    "hlink": "#467886",
    "folHlink": "#96607D"
  },
  "fonts": {
    "title": "Inter",
    "body": "Inter",
    "mono": "Aptos Mono",
    "fallbackSans": "Aptos",
    "fallbackDisplay": "Aptos Display"
  },
  "typeScale": {
    "titlePt": 44,
    "subheadingPt": 24,
    "bodyPt": 28,
    "bulletPt": 28,
    "inlineCodePt": 24,
    "codeBlockPt": 20,
    "calloutHeadingPt": 24,
    "calloutBodyPt": 24,
    "minBodyPt": 18,
    "minCodePt": 16
  },
  "layout": {
    "marginLeft": 0.6,
    "marginRight": 0.6,
    "marginTop": 0.45,
    "marginBottom": 0.4,
    "titleHeight": 0.8,
    "contentGapAfterTitle": 0.18,
    "paragraphGap": 0.14,
    "bulletGap": 0.08,
    "calloutGap": 0.18,
    "calloutPadding": 0.16,
    "calloutIconSize": 0.32,
    "calloutIconGap": 0.14,
    "codePadding": 0.14,
    "mathBlockGap": 0.14
  },
  "text": {
    "titleColor": "#29322E",
    "bodyColor": "#29322E",
    "mutedColor": "#423B2D",
    "inlineCodeColor": "#212121",
    "linkColor": "#467886"
  },
  "blocks": {
    "callout": {
      "fill": "#FEFFFF",
      "border": "#B1813E",
      "borderWidth": 1.5,
      "radius": 0.12
    },
    "code": {
      "fill": "#FEFFFF",
      "border": "#3E4641",
      "borderWidth": 1,
      "radius": 0.08
    },
    "inlineCode": {
      "fill": "#EDE8E2"
    }
  },
  "fit": {
    "scales": [1.0, 0.94, 0.89, 0.84],
    "minScale": 0.84
  },
  "emojiMap": {
    "✅": "./assets/checkmark-square_orange.svg"
  }
}
```

## 23. One important unresolved choice

There are two valid inline-math strategies:

- render inline math as small SVG images inserted between text runs
- downgrade inline math to body text with a TeX-ish style unless the paragraph is math-heavy

Recommendation:

implement true inline SVG math from the start if the chosen PPTX library handles mixed run/image layout acceptably

That is the only materially uncertain part of the proposed v1 design.