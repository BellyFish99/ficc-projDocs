# PowerPoint Tips — Claude Add-in & Style Extraction

## Contents

1. [Claude Add-in Use Cases & Prompts](#1-claude-add-in-use-cases--prompts)
2. [Extracting & Applying Style from a Company Template](#2-extracting--applying-style-from-a-company-template)
3. [Quick Reference Tables](#3-quick-reference-tables)

---

## 1. Claude Add-in Use Cases & Prompts

### How it works
Select content on a slide (or nothing for whole-slide context), open the Claude panel, type your prompt. Claude sees the selected text or slide content and responds — copy/paste the result back onto the slide.

---

### 1.1 Rewrite & Tighten Text

**Bullets too long:**
```
These bullets are too verbose for a CEO audience.
Rewrite each as a single punchy line under 10 words.
Keep all the numbers and specifics.
```

**Simplify technical content:**
```
This slide is for a CEO who is not technical.
Rewrite this content removing jargon, keeping the business
impact clear. Each point should answer "so what does this
mean for my business?"
```

**Make it more urgent:**
```
Rewrite this content to create urgency.
This is a sales pitch — the reader should feel
they cannot afford to delay. Keep it factual, no hype.
```

**Condense — too much text:**
```
This text box has too much content for one slide.
Cut it to the 3 most important points only.
Prioritize anything with a dollar figure or a risk.
```

---

### 1.2 Generate New Content

**Write a McKinsey-style action title:**
```
Write 5 alternative slide titles for this content.
Each title should be a complete sentence that states
the conclusion (McKinsey action title style), not just
a topic label. Max 12 words each.
```

**Write speaker notes:**
```
Write speaker notes for this slide.
The presenter is a sales VP pitching to a CIO.
Include: the key point to land, 1-2 questions to ask
the audience, and a transition to the next topic.
Under 150 words.
```

**Create a callout box / key insight:**
```
Based on this slide content, write a single "key insight"
callout box — 1 bold header (max 8 words) and
2 supporting sentences. This should be the one thing
the reader must remember from this slide.
```

**Convert prose to table:**
```
Convert this paragraph into a 3-column table with columns:
[Problem] [Current State] [Solution].
Keep each cell under 15 words.
```

---

### 1.3 Translate & Localize

**Chinese to English:**
```
Translate this slide content to English.
Keep all numbers and product names (e.g. POMS, IBOR, IFRS9)
unchanged. Use financial industry terminology.
Maintain the same structure and bullet format.
```

**English to Chinese:**
```
Translate to Simplified Chinese.
Keep technical terms like VaR, DV01, Sharpe Ratio,
CEP, IBOR, IFRS9 in English.
Use formal financial Chinese (金融专业术语).
```

**Adjust for a different audience:**
```
This content is written for a CTO.
Rewrite it for a CFO — same facts, but focus on
cost, ROI, risk reduction, and regulatory compliance
instead of technical architecture.
```

---

### 1.4 Strengthen Arguments

**Add evidence to claims:**
```
This slide makes claims but lacks evidence.
For each bullet point, suggest what data, benchmark,
or proof point would make it more credible.
Format as: [Original claim] → [Suggested evidence to add]
```

**Anticipate objections:**
```
Read this slide as a skeptical CIO who has heard
many vendor pitches. List the top 3 objections
they would raise, and for each, write a 1-sentence
rebuttal I can use in Q&A.
```

**Sharpen the value proposition:**
```
Rewrite the value proposition on this slide using
this structure:
[Who] [has this problem] [our solution] [delivers this outcome]
[unlike alternatives which do X].
Keep it under 3 sentences.
```

---

### 1.5 Structure & Flow

**Check if slide is MECE:**
```
Review these bullet points. Are they MECE
(mutually exclusive, collectively exhaustive)?
Identify any overlaps or gaps, and suggest how
to restructure them to be truly MECE.
```

**Create a slide from raw notes:**
```
I have these rough notes: [paste your notes]

Turn this into a slide with:
- 1 action title (the conclusion)
- 3-4 bullet points (the evidence)
- 1 callout box (the so-what)
Keep each bullet under 15 words.
```

**Write slide transitions:**
```
I have 3 slides in sequence. Here is the content of each:
Slide A: [paste]
Slide B: [paste]
Slide C: [paste]

Write a 1-sentence transition for the presenter
to say between each slide that makes the logical
flow feel natural.
```

---

### 1.6 Numbers & Data

**Sanity-check calculations:**
```
Check these numbers for consistency.
If AUM is 100B CNY and trading volume is 300B/year
(3x turnover), and we claim 1bp improvement saves 3000万,
verify that math and flag any inconsistencies.
```

**Make data more vivid:**
```
Rewrite these statistics to make them more vivid
and memorable for a non-technical executive.
Use analogies or comparisons where helpful.
Keep the actual numbers unchanged.
```

**Summarize a table into prose:**
```
Summarize the key insight from this table in
2 sentences. What is the single most important
thing a CEO should take away from this data?
```

---

### 1.7 Layout Guidance (add-in can advise, not execute)

**Describe layout rules:**
```
Look at this slide layout. What are the spacing rules
being used? Estimate: margin size, gap between elements,
title bar height, and column widths as a proportion
of slide width.
```

**Get font pairing suggestions:**
```
This slide uses [Font Name] as the heading font.
Suggest 3 system fonts available on Windows that have
a similar feel and would pair well with it for body text.
```

**Extract color description:**
```
Look at this slide. List every color used —
background, text, shapes, lines — as precisely
as you can describe them (dark navy, bright red, etc.).
I need to recreate this palette.
```

---

### 1.8 Tips for Better Add-in Results

Always include in your prompt:
1. **Who the audience is** — CEO, CTO, CFO, investor
2. **What you want them to feel or do** — approve budget, feel urgency, trust the vendor
3. **Constraints** — word count, must keep specific numbers, language

**Pro tip:** The more specific, the better.

| Vague | Specific |
|-------|----------|
| "Make this better" | "Cut to 3 bullets, each under 10 words, keep all dollar figures" |
| "Translate this" | "Translate to Chinese, keep POMS/VaR/IBOR in English, formal register" |
| "Simplify" | "Rewrite for a CFO with no technical background, focus on ROI and risk" |

---

## 2. Extracting & Applying Style from a Company Template

### What "Style" Means in PowerPoint

There are 5 layers — they're separate and need different approaches:

| Layer | What it controls | Where it lives |
|-------|-----------------|----------------|
| Theme Colors | The 10 accent colors in color picker | Design → Variants → Colors |
| Theme Fonts | Heading + Body font pair | Design → Variants → Fonts |
| Slide Master | Background, logo, footer, title bar layout | View → Slide Master |
| Slide Layouts | Individual layout templates (Title, Content, etc.) | Under Slide Master |
| Object Styles | Default shape/text box styling | Right-click → Set as Default |

---

### Method 1: Save & Apply a Theme File (30 seconds)

Copies: colors + fonts + effects. Does **not** copy backgrounds or logos.

**Extract from company deck:**
```
Open company PPT
→ Design tab
→ Variants dropdown (click bottom-right arrow)
→ Save Current Theme...
→ Save as: CompanyTheme.thmx
```

**Apply to new deck:**
```
Open your new PPT
→ Design tab
→ Themes dropdown
→ Browse for Themes...
→ Select CompanyTheme.thmx
```

---

### Method 2: Steal the Slide Master (Full Layout Transplant)

Copies: backgrounds, title bars, logos, layout templates — the full visual structure.

**Copy-paste master between files:**
```
Open company PPT
→ View → Slide Master
→ Right-click the master (top thumbnail in left panel)
→ Copy

Switch to your new PPT
→ View → Slide Master
→ Right-click below existing master
→ Paste
```

After pasting, reassign slide layouts:
```
View → Normal
→ Right-click any slide thumbnail
→ Change Layout
→ Pick a layout from the imported master
```

---

### Method 3: Format Painter (Object-by-Object)

Copies fill color, border, font, size, and effects from one object to another.

```
Click the source object (the one with the style you want)
→ Home → Format Painter (paintbrush icon)
→ Click the target object
```

**Tip:** Double-click Format Painter to apply to multiple objects in a row. Press Esc when done.

---

### Method 4: Extract Exact Values via Claude Code CLI

When you need the exact hex codes, font names, and sizes from the XML source:

```bash
# Unpack any PPTX to readable XML
python scripts/office/unpack.py "company.pptx" unpacked/

# Theme colors (hex codes, accent names):
cat unpacked/ppt/theme/theme1.xml

# Slide master fonts and layouts:
cat unpacked/ppt/slideMasters/slideMaster1.xml

# Any specific slide's styling:
cat unpacked/ppt/slides/slide1.xml
```

Then ask Claude Code CLI:
```
Read unpacked/ppt/theme/theme1.xml and extract:
- All theme color hex codes with their names (dk1, lt1, accent1–6)
- Heading and body font names
- Any gradient or background definitions
```

Example output:
```python
# Exact values extracted from theme XML:
DARK_NAVY    = RGBColor(0x03, 0x0E, 0x42)   # dk1
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)    # lt1
BRAND_RED    = RGBColor(0xF5, 0x4D, 0x61)   # accent1
BRAND_YELLOW = RGBColor(0xFF, 0xD0, 0x51)   # accent2
FONT_HEADING = "MiSans Normal"
FONT_BODY    = "MiSans Normal"
```

---

### Method 5: Apply Styles Programmatically (python-pptx)

Use when you need to change a property across all slides at once.

**Change font across entire deck:**
```python
from pptx import Presentation
from pptx.util import Pt

prs = Presentation("my_deck.pptx")

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = "MiSans Normal"

prs.save("my_deck_restyled.pptx")
```

**Replace a color everywhere:**
```python
from pptx.dml.color import RGBColor

OLD_COLOR = RGBColor(0xFF, 0x00, 0x00)   # old red
NEW_COLOR = RGBColor(0xF5, 0x4D, 0x61)   # brand red

for slide in prs.slides:
    for shape in slide.shapes:
        try:
            if shape.fill.fore_color.rgb == OLD_COLOR:
                shape.fill.fore_color.rgb = NEW_COLOR
        except:
            pass  # shape has no solid fill

prs.save("my_deck_recolored.pptx")
```

**Change all title font sizes:**
```python
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    if run.font.size and run.font.size > Pt(20):
                        run.font.size = Pt(24)  # standardize titles
```

---

### Recommended Workflow: Company Template → New Deck

**Step 1 — Extract theme (30 sec):**
```
Open company PPT → Design → Variants → Save Current Theme → CompanyTheme.thmx
```

**Step 2 — Apply theme (30 sec):**
```
Open new PPT → Design → Browse for Themes → CompanyTheme.thmx
```

**Step 3 — Import Slide Master (2 min):**
```
Company PPT → View → Slide Master → Copy top master
New PPT → View → Slide Master → Paste
```

**Step 4 — Fix remaining issues at scale:**
If specific colors or fonts are still off across many slides, ask Claude Code CLI to fix them programmatically — changing 40+ slides at once takes seconds vs. manual edits.

---

## 3. Quick Reference Tables

### Claude Add-in: Goal → Key Phrase

| Goal | Key phrase to use in prompt |
|------|-----------------------------|
| Shorter | "Cut to 3 points only, keep all numbers" |
| Simpler | "Rewrite for a CEO, remove jargon" |
| More urgent | "Make this more urgent, keep it factual" |
| McKinsey style | "Rewrite as an action title — state the conclusion" |
| Table | "Convert to a 3-column table: X / Y / Z" |
| Speaker notes | "Write speaker notes, under 150 words" |
| Translate CN→EN | "Translate to English, keep POMS/VaR/IBOR unchanged" |
| Translate EN→CN | "Translate to Simplified Chinese, formal financial register" |
| Objections | "What would a skeptical CIO push back on?" |
| Transitions | "Write a 1-sentence transition to the next slide" |
| MECE check | "Are these bullets MECE? Identify overlaps and gaps" |
| Vivid data | "Make these numbers more vivid, keep figures unchanged" |

---

### Style Extraction: Which Tool for What

| Task | Best Tool |
|------|-----------|
| Copy color + font theme | PowerPoint: Save/Apply Theme (.thmx) |
| Copy backgrounds + logos + title bars | PowerPoint: Copy Slide Master |
| Copy one object's style to another | PowerPoint: Format Painter |
| Get exact hex codes and font names | Claude Code CLI: unpack XML |
| Change a color across ALL slides | Claude Code CLI: python-pptx script |
| Change font across ALL slides | Claude Code CLI: python-pptx script |
| Understand spacing/layout rules | Claude add-in: describe the slide |
| Redesign a single slide layout | PowerPoint: manual + Format Painter |
| Apply theme from file | PowerPoint: Design → Browse for Themes |
| Full visual structure transplant | PowerPoint: Slide Master copy-paste |
