# Helping AI Agents "See" Visual Output: Research Findings

**Context**: Clean Slides generates PowerPoint slides from YAML specs. When extending to complex elements (waterfall charts, icons, rich layouts), the agent struggles to verify its own visual output — it can't reliably spot overlapping elements, mispositioned labels, or text overflow, even with PNG rendering feedback. This doc surveys how the broader ecosystem handles this class of problem.

---

## 1. Why Vision Models Fail at Spatial Precision

### The architectural root cause

Vision-Language Models (VLMs) compress images through patch embeddings that **fundamentally destroy pixel-level spatial indexing**. A 1024×1024 image gets compressed into a 64×64 grid of tokens — a 256-fold compression. Recent research quantifies the damage:

> "Vision encoders cause 40-60% k-nearest neighbor divergence, meaning nearly half of the local geometric structure disappears during encoding." — QVLM (Massih & Cosatto, NEC Labs, Jan 2026)

This means **no amount of prompting can recover architecturally discarded spatial information**. The model that can eloquently describe a chart cannot reliably tell you that a label is 0.1" off-center.

**Key papers:**
- **QVLM** (arXiv 2601.13401, Jan 2026): Demonstrates that VLMs achieve only 37-42% accuracy on counting tasks despite strong qualitative understanding. Proposes code-generation architecture that decouples language from visual analysis.
- **SpatialVLM** (CVPR 2024, Google DeepMind): Confirms VLMs lack 3D/spatial reasoning — hypothesizes this is a training data limitation, not architectural, but their fix requires 2B synthetic training examples.
- **SiT-Bench** (arXiv 2601.03590, Jan 2026): Shows that converting visual scenes into "coordinate-aware textual descriptions" lets LLMs do spatial reasoning via text — models achieve **better** spatial reasoning from structured text than from pixels.
- **Emergent Mind survey** (Oct 2025): LLM spatial reasoning "rapidly deteriorates as problem scale or compositional/geometric complexity increases, with performance losses ranging from 42% to over 80%."

### Benchmark evidence for slide-like content

A 2026 visual reasoning benchmark (aimultiple.com) testing 9 multimodal models found:
- Even the best models (Gemini 2.5 Pro, GPT-5.2) only hit **62% overall** on chart understanding + visual logic
- The hardest task: reading clustered bar charts with 4 bars per group. Models could not accurately distinguish which bar belonged to which category or compare relative heights
- The easiest: single-bar-per-category charts with clear spacing — models excel at simple min/max identification

> "Current models lacked the pixel-precise perception needed to correctly measure and sequence densely packed bars, leading to systematic misidentification of trends."

**Implication for Clean Slides**: Expecting a vision model to verify a waterfall chart's label positions by looking at a 150 DPI PNG is unreliable. The model may say "looks good" when labels are overlapping — this isn't a resolution issue, it's an architectural limitation.

---

## 2. Analogous Domains & How They Solve It

### 2a. Frontend: Visual Regression Testing

The web development world has grappled with this exact problem and converged on several approaches:

**Pixel-diffing** (BackstopJS, Percy, Chromatic):
- Capture baseline screenshots, then compare new screenshots pixel-by-pixel after code changes
- Works well for regression (detecting change), not for *correctness* (detecting first-time layout bugs)
- Inherently unstable across environments — font rendering, GPU drivers, OS differences all cause false positives
- Vitest (2026): "Screenshots will look different on different machines because of font rendering, GPU drivers, hardware acceleration, whether you're running headless or not, browser versions"

**DOM/structure-based testing** (Galen Framework — the most relevant):
- Tests **relative spatial relationships between elements**, not pixels
- Uses a human-readable DSL to define layout constraints:
  ```
  header:
    inside screen 0px top, 0px left, 0px right
  menu:
    below header 0px
    inside screen 0px left right
  side-panel:
    width 300px
    near content 0px right
  ```
- Reports violations as structured text with highlighted screenshots
- **Key insight**: Instead of asking "does this look right?", it asks "does element X satisfy constraint Y relative to element Z?"

**Accessibility tree / computed layout** (Playwright, mobile QA):
- Instead of looking at pixels, analyzes the **structure** of UI elements
- For mobile: inspects View Hierarchy (iOS) or Accessibility Tree (Android)
- Gives you positions, sizes, relationships as structured data

> "Instead of comparing raw pixels, DOM-based testing analyzes the structure of the page (HTML elements, CSS properties, layout hierarchy)." — Virtuoso QA

**Key takeaway**: The frontend world learned that pixel comparison is fragile and DOM-based / constraint-based testing is robust. The analog for slides: don't compare rendered PNGs, verify spatial constraints on the OOXML shape tree.

### 2b. Chip Design: Design Rule Checking (DRC)

DRC is the closest structural analogy to what Clean Slides needs. In semiconductor manufacturing:

- A designer creates a layout (analogous to YAML spec → OOXML shapes)
- The layout gets "rendered" by a fab process (analogous to LibreOffice/PowerPoint rendering)
- You **cannot just look at the result** — it's a physical chip
- Instead, you verify **geometric constraints programmatically** before fabrication

DRC rules include:
- **Width**: minimum width of any shape
- **Spacing**: minimum distance between two adjacent objects
- **Enclosure**: one object must be covered by another with N margin
- **Minimum area**: shapes must be above a minimum size
- **Antenna rules**: complex ratios between layer areas

> "While design rule checks do not validate that the design will operate correctly, they are constructed to verify that the structure meets the process constraints." — Wikipedia, Design Rule Checking

**Relevance**: Clean Slides' `verification.py` is already a rudimentary DRC engine for tables (overlap, boundary, text-fit). The approach scales — define geometric constraints per element type.

### 2c. PowerPoint-Specific Tools

Commercial tools like **PPT Productivity** and **UpSlide** offer automated PowerPoint proofing:
- Font type/size consistency checks
- Bullet style and margin consistency
- Content outside slide boundaries detection
- Misaligned shape detection
- Large image detection

These are purely structural checks on the OOXML — no rendering or vision involved. They work reliably because they operate on the data model, not pixels.

### 2d. AI Agent Best Practices (Anthropic)

Anthropic's own best practices for Claude Code explicitly address this:

> "Claude performs dramatically better when it can verify its own work, like run tests, compare screenshots, and validate outputs. Without clear success criteria, it might produce something that looks right but actually doesn't work."

Their recommended pattern for UI verification:
1. Take a screenshot of the result
2. Compare it to the original/expected
3. List differences and fix them

But they also note this works best when combined with **programmatic verification** (test suites, linters, validators). The screenshot is the last resort, not the primary verification mechanism.

---

## 3. Research-Backed Approaches (Most → Least Proven)

### Approach A: Expanded Constraint Verification (DRC for Slides)

**Evidence**: DRC in chip design (40+ years of validation), Galen Framework for web layout (proven in production at scale), Clean Slides' own `verification.py` already works for tables.

**How it works**:
- Define spatial constraints per element type (waterfall bar, label, icon, connector)
- Verify constraints programmatically against the OOXML shape tree
- Output structured violation reports the agent can read and act on

**Example constraints for a waterfall chart**:
```
bar_label "Revenue":
  centered_horizontally_within bar "Revenue" ± 0.05in
  above bar "Revenue" by 0.05in to 0.15in
  text_fits_within_width bar "Revenue"

bar "Cost":
  adjacent_to bar "Revenue" with gap 0.1in to 0.3in
  top_aligned_with connector from bar "Revenue"

all bars:
  no_overlaps
  within content_area
  uniform_gap ± 0.02in
```

**Strengths**: Deterministic, instant, no rendering required, agent reads structured text.
**Weaknesses**: Must define rules per element type (one-time cost per type).

### Approach B: Structured Spatial Report

**Evidence**: SiT-Bench (Jan 2026) shows LLMs reason about space **better from structured text than from images**. Galen Framework generates structured layout reports.

**How it works**:
After generating a slide, produce a machine-readable spatial analysis:
```
SPATIAL REPORT — slide 1
══════════════════════════

ELEMENTS (8 shapes, sorted top→bottom, left→right):
  [0] Label "Revenue"     (2.30, 1.20) 1.10×0.25in
  [1] Bar "Revenue"       (2.30, 1.50) 0.80×3.20in  fill=#2563EB
  [2] Label "Cost"        (3.50, 1.20) 1.10×0.25in
  [3] Bar "Cost"          (3.50, 2.10) 0.80×2.60in  fill=#E5546C

PAIRWISE RELATIONSHIPS:
  Label[0] ↔ Bar[1]: centered ✓ (Δx=0.00in), above by 0.30in
  Label[2] ↔ Bar[3]: centered ✓ (Δx=0.00in), above by 0.90in ⚠ (expected 0.30in)
  Bar[1] ↔ Bar[3]: gap=0.40in, top offset=0.60in

ISSUES:
  ⚠ Label[2] "Cost" is 0.60in above its bar (expected ~0.30in)
```

**Strengths**: Agent can read and reason about it; catches issues without needing rules per element type.
**Weaknesses**: Still requires knowing which elements are related (needs semantic tagging).

### Approach C: Architectural Decoupling (QVLM pattern)

**Evidence**: QVLM (Jan 2026) achieves 42% vs 28% accuracy by having the LLM generate code that calls specialized vision tools, rather than "looking" at images directly.

**How it works for slides**:
Instead of: LLM sees PNG → tries to spot issues
Do: LLM generates verification code → code extracts positions from OOXML → code checks constraints → returns structured results to LLM

This is essentially Approach A, but framed as the agent *writing* its own verification rather than relying on pre-built checks.

**Strengths**: Flexible — agent can write ad-hoc checks for novel situations.
**Weaknesses**: Adds latency; agent must know what to check; the generated verification code itself could have bugs.

### Approach D: OCR-Based Render Gap Detection

**Evidence**: Tesseract/pytesseract can extract text bounding boxes from rendered PNGs. Modern OCR models (OlmOCR-2, Chandra, etc.) achieve 75-83% accuracy on document parsing.

**How it works**:
1. Render slide to PNG via LibreOffice/PowerPoint
2. Run OCR on the PNG to extract text positions
3. Compare OCR'd positions against expected positions from the OOXML spec
4. Discrepancies reveal where the renderer disagrees with your layout model

```
RENDER GAP ANALYSIS:
  "Revenue" — expected at (2.30, 1.20), OCR'd at (2.31, 1.22) — OK (0.02in drift)
  "Operating costs" — expected 1 line, OCR'd as 2 lines — MISMATCH
    → text_metrics underestimates width for this string at 8pt Arial
```

**Strengths**: Catches the LibreOffice/PowerPoint rendering mismatch that structural checks miss.
**Weaknesses**: OCR bounding boxes are imprecise (known issue — Azure, Tesseract both have accuracy problems), adds rendering dependency, slower.

### Approach E: SVG Intermediate Representation

**Evidence**: LibreOffice can export to SVG. SVG is a text-based format with exact element positions.

**How it works**:
- Export slide to SVG instead of (or in addition to) PNG
- Parse SVG to extract actual rendered positions of text and shapes
- Compare against expected positions

**Strengths**: More precise than OCR, text-based so easily parseable.
**Weaknesses**: LibreOffice SVG export fidelity isn't perfect; less available than PNG rendering; still requires rendering step.

### Approach F: Element-Level Cropping (SlideAgent Pattern)

**Evidence**: SlideAgent (Georgia Tech + JP Morgan, Oct 2025) demonstrates that VLMs perform dramatically better when individual elements are **cropped and analyzed separately** vs. viewing the full page.

> "When given the full page, the LLM miscounts the number of product mix categories. After isolating the chart, it correctly identifies all eight categories, highlighting the importance of accurate element parsing."

**How it works**:
- Render slide to high-DPI PNG
- For each element of interest, crop its bounding box (+ margin)
- Feed individual crops to the vision model for focused analysis
- Agent sees one element at a time rather than a busy full-page image

**Relevance**: This is essentially what `pptx crop` already enables. The missing piece is automation — the agent needs to know *what* to crop and *what to look for* in each crop.

**Strengths**: Works with existing vision capabilities; no new infrastructure.
**Weaknesses**: Still relies on vision model precision; doesn't scale to checking dozens of spatial relationships; agent needs to know where to look.

---

## 4. Design: Placement-Based Verification for Generated Slides

### The key insight

Clean Slides controls the generation pipeline. The renderer already computes every shape's position, size, role, and text content — it just throws that information away after writing OOXML. If we capture it, verification becomes trivial: run checks on a flat list of typed, positioned shapes.

No heuristics. No inference. No vision model. The renderer tells you what it meant; the verifier checks whether the geometry is sound.

### Scope: generation first, edit later

This design covers the `pptx generate` workflow:
1. Agent writes YAML → `pptx generate` → `.pptx` + **placement report**
2. Agent reads the report → sees issues (or "OK")
3. Agent fixes YAML → re-generates

The `pptx edit` workflow (agent tweaks individual shapes in an existing file) is a separate problem — it requires working from `inspect_slide()` output rather than renderer-emitted placements. That's deferred.

### Data model

```python
@dataclass
class Placement:
    """A shape placed by the renderer, with its intended role."""
    name: str        # shape name in the PPTX (already set by renderer)
    role: str        # "bar", "label", "connector", "cell", "header", "icon", ...
    box: Box         # (x, y, w, h) in EMU — already computed by renderer
    group: str = ""  # links related elements (see below)
    text: str = ""   # text content, for text-fit checks
    font: str = ""   # font name
    size_pt: int = 0 # font size in points
```

**Where this data comes from**: The renderer is already computing all of these values to build OOXML. Capturing them is a few `.append()` calls per element type — not a framework, not an abstraction layer.

**`group` — linking related shapes**: When the renderer places a label above a bar, it knows they're related. `group` is a shared string key (e.g. `"Revenue"`) that lets checks pair them without guessing by proximity. The renderer assigns it because it created both shapes together.

Example: a waterfall chart renderer produces:
```
Placement(name="bar_Revenue",       role="bar",       group="Revenue", box=...)
Placement(name="label_Revenue",     role="label",     group="Revenue", box=...)
Placement(name="connector_Rev_Cost", role="connector", group="",       box=...)
Placement(name="bar_Cost",          role="bar",       group="Cost",    box=...)
Placement(name="label_Cost",        role="label",     group="Cost",    box=...)
```

Then `check_group_alignment` pairs label↔bar by matching `group` values. Without `group`, you'd need heuristic proximity matching — exactly the black-box inference we're avoiding.

**When `group` isn't needed**: Generic checks (overlap, bounds, text-fit) don't use `group` at all. It's only needed for relationship checks like "is this label centered above its bar?" If a new element type doesn't have paired elements, just leave `group` empty.

### Checks: standalone functions, not a framework

Each check is a pure function: `list[Placement]` → `list[LayoutIssue]`. No classes, no inheritance, no registration. You read it, you see what it does.

```python
def check_overlaps(placements: list[Placement]) -> list[LayoutIssue]:
    """Any two non-background shapes overlap?"""
    # ~15 lines: pairwise bbox intersection test
    # Skips pairs where either role is "background" or "connector"

def check_bounds(placements: list[Placement], area: ContentArea) -> list[LayoutIssue]:
    """Any shape exceeds the content area?"""
    # ~10 lines: compare each box against area bounds

def check_text_fits(placements: list[Placement], metrics: TextMetrics) -> list[LayoutIssue]:
    """Text content overflows its bounding box?"""
    # ~20 lines: for placements with text, check longest-word width + content height

def check_group_alignment(placements: list[Placement]) -> list[LayoutIssue]:
    """Within each group, are labels centered over their anchors?"""
    # ~20 lines: group by `group` key, find role="label" + role="bar",
    # check horizontal center alignment within tolerance

def check_uniform_spacing(placements: list[Placement], role: str) -> list[LayoutIssue]:
    """Are shapes of a given role evenly spaced?"""
    # ~15 lines: filter by role, sort by x, compute pairwise gaps,
    # flag if max gap - min gap exceeds tolerance
```

**Adding a new check**: Write a function. Call it from `run_all()`. That's it.

**Adding a new element type**: The new renderer emits `Placement` entries with appropriate roles. Existing checks (overlap, bounds, text-fit) work automatically. If the element has novel relationships (e.g. "icon must be left-aligned with its row label"), write one new check function.

### Integration with existing code

The current pipeline:

```
YAML → TableSpec → ConstraintSolver.solve() → TableLayout + LayoutReport → renderer
```

The solver already returns a `LayoutReport`. Two options for where placements live:

**Option 1: Renderer returns placements**. After generating shapes, the renderer also returns `list[Placement]`. The CLI runs checks and merges issues into the report. This is the natural place — the renderer has all the information.

**Option 2: Solver builds placements from TableLayout**. For tables, the solver already has the cell grid. It can emit `Placement(role="cell", ...)` entries. This keeps verification in the solver (where it already is) and doesn't require renderer changes for tables.

**Recommendation**: Option 2 for tables (minimal change to existing code), Option 1 for new element types (renderer knows the semantics). Both feed into the same `list[Placement]` → same check functions.

### What the agent sees

The `pptx generate` output goes from:

```
input.yaml: OK
```
or:
```
input.yaml: 2 warning(s) (--detail for list)
```

To (with `--detail`):

```
input.yaml: 1 error, 1 warning

  ERROR  overlap: "label_Cost" overlaps "bar_Revenue"
         label_Cost (3.40, 1.20) 1.20×0.25in ∩ bar_Revenue (2.30, 1.10) 0.80×3.20in

  WARN   spacing: bars have uneven horizontal gaps
         bar_Revenue→bar_Cost: 0.40in, bar_Cost→bar_Margin: 0.18in (Δ0.22in)
```

Structured text. No screenshots needed. The agent can read this and know exactly what to fix in the YAML.

### What NOT to invest in

- **Better prompting for vision inspection**: The architectural limitations of VLMs are fundamental — no prompt will make them reliably notice a 0.1" misalignment
- **Comprehensive visual regression tests**: Too brittle for the combinatorial space of chart types × formatting options
- **ASCII rendering**: Slides are too visually complex for ASCII representation to be useful
- **Heuristic relationship inference**: Trying to guess which label belongs to which bar from an arbitrary PPTX. For generation, the renderer knows — use that knowledge directly

---

## 5. Sources

| Source | Key Finding |
|--------|-------------|
| QVLM (arXiv 2601.13401, Jan 2026) | VLMs fail at pixel-precision; code-generation architecture 50% better |
| SiT-Bench (arXiv 2601.03590, Jan 2026) | LLMs reason better about space from structured text than pixels |
| SpatialVLM (CVPR 2024) | VLMs lack spatial reasoning; 2B synthetic examples needed to improve |
| SlideAgent (arXiv 2510.26615, Oct 2025) | Element-level cropping dramatically improves VLM accuracy on slides |
| aimultiple.com benchmark (2026) | Best VLMs only 62% on chart understanding; fail on dense layouts |
| Emergent Mind spatial reasoning survey (Oct 2025) | 42-80% performance loss as spatial complexity increases |
| Galen Framework | DSL for relative spatial constraint testing in web layouts |
| Chromatic / Percy / Vitest | Pixel-diffing is inherently fragile across environments |
| Panto AI (2026 mobile QA guide) | Structure-based testing (accessibility tree) > pixel comparison |
| Wikipedia: Design Rule Checking | 40+ years of geometric constraint verification in chip design |
| PPT Productivity / UpSlide | Commercial PPTX proofing operates on OOXML structure, not pixels |
| Anthropic best practices (Sept 2025) | Agents work best with programmatic verification + visual spot-checks |
| Claude Code docs | "Claude performs dramatically better when it can verify its own work" |
