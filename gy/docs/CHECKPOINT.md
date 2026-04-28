# Project Checkpoint — 华锐技术 (Archforce Technology) GY

**Last updated:** 2026-04-24  
**Working directory:** `D:\Work\gy`

---

## Session 2026-04-24 — Brand Styling Sprint

### What We Did

#### 1. Learned the Company Brand
- Read and analysed `基于国内外FICC平台发展经验规划FICC平台建设路径0417.pdf` (58-slide FICC platform planning deck)
- Extracted the full Archforce Technology visual identity: colors, layout patterns, typography rules
- Saved the guide permanently at machine level (see §4 below)

#### 2. Styled CEO_Deck_v9.pptx
- Source: `CEO_Deck_v8.pptx` (43 slides)
- Output: `CEO_Deck_v9.pptx` ✅

| Old Color | Role | New Brand Color |
|-----------|------|-----------------|
| `#030E42` | Dark background | `#0F2060` |
| `#FFD051` (shape fill) | Accent shapes | `#E53935` (brand red) |
| `#FFD051` (text) | Yellow text | `#FFFFFF` (white) |
| `#56C4D0` | Teal elements | `#3B7DD8` |
| `#379DFF` | Bright blue | `#2B5EC7` |
| `#1E5BA8` | Medium blue | `#1B3275` |
| `#B0B8D0` | Gray-blue | `#D6E4F7` |

**Lesson learned:** Red text hides in three XML locations in PPTX — `<a:rPr>`, `<a:defRPr>`, and shape-fill contexts. Must fix all three or red text reappears on slides.

#### 3. Restyled All 14 Draw.io Diagrams
Applied Archforce brand to every `.drawio` file in `D:\Work\gy\`:

| File | Type | Content |
|------|------|---------|
| `diagram_1_投资经理的一天.drawio` | Timeline | Investment manager daily workflow |
| `diagram_3_924行情推演.drawio` | Flowchart | 924 market scenario analysis |
| `diagram_4_竞争定位.drawio` | Positioning | Competitive positioning map |
| `diagram_5_ROI决策框架.drawio` | Flowchart | ROI decision framework |
| `diagram_6_实施路线图.drawio` | Roadmap | Implementation roadmap |
| `diagram_7_CEP风控引擎.drawio` | Architecture | CEP risk control engine |
| `diagram_8_指令执行泳道.drawio` | Swimlane | Instruction execution flow (POMS) |
| `diagram_A_功能全景图.drawio` | Architecture | Full functional landscape |
| `diagram_B_系统组件图.drawio` | Component | System component diagram |
| `diagram_C_集成架构图.drawio` | Architecture | Integration architecture |
| `diagram_D_需求追溯矩阵.drawio` | Matrix | Requirements traceability |
| `diagram_E_数据流图.drawio` | Data flow | Data flow diagram |
| `diagram_F_信创技术栈.drawio` | Stack | 信创 domestic tech stack |
| `diagram_G_部署架构.drawio` | Deployment | Deployment architecture |

#### 4. Saved Brand Style at Machine Level (Permanent)
Brand now loads automatically in every Claude session — no need to re-explain:

| File | Scope | What it contains |
|------|-------|-----------------|
| `~/.claude/CLAUDE.md` | **Global — all projects, all sessions** | 6 core brand rules, auto-loaded |
| `~/.claude/skills/ppt-composer/brand-reference.md` | Machine-level skill | Full detail: color palette, PPTX mapping, draw.io rules for swimlane / architecture / flowchart |
| `~/.claude/projects/-mnt-d-Work-gy/memory/project_archforce_brand_style.md` | This project | Project-level backup with history |

---

## Brand Quick Reference

### Color Palette

| Role | Hex | Usage |
|------|-----|-------|
| Dark Navy | `#0F2060` | Title bars, diagram headers, section divider backgrounds |
| Deep Navy | `#1B3275` | Lane labels, borders, arrows, body text on light bg |
| Medium Blue | `#2B5EC7` | Primary node / card fill |
| Brand Blue | `#3B7DD8` | Secondary nodes, component borders |
| Accent Blue | `#4A90D9` | Tertiary nodes |
| Red Accent | `#E53935` | **Sparingly** — rejection paths and risk warnings ONLY |
| White | `#FFFFFF` | Text on any dark fill |
| Component Fill | `#D6E4F7` | Component box backgrounds |
| Canvas | `#F5F7FA` | Diagram canvas / slide background tint |

Lane tints (swimlane backgrounds): `#EEF3FA` · `#E8EFF8` · `#F0F5FF` · `#EBF1FA` · `#E5EDF8`

### Draw.io Rules at a Glance

- **Swimlane header:** `#0F2060` fill, white text
- **Lane label:** `#1B3275` fill, white text
- **Process nodes:** `#2B5EC7` fill, `#1B3275` stroke, white text, `rounded=1`
- **Arrows:** `strokeColor=#1B3275; strokeWidth=1.5`
- **Rejection path:** `strokeColor=#E53935; dashed=1` (only legitimate use of red)

### PPTX Rules at a Glance

- Dark slide backgrounds → `#0F2060`
- Red text must be fixed in `<a:rPr>` + `<a:defRPr>` + shape fills
- Section divider bars → `#E53935` (thin, left edge of title)
- Never use yellow, orange, green, or purple as slide accent colors

---

## Scripts & Tools

| Script | Where | How to use |
|--------|-------|------------|
| `brand_style.py` | `/tmp/brand_style.py` | Restyle any .drawio file to brand palette |
| `unpack.py` | `~/.claude/skills/pptx/scripts/office/unpack.py` | Unpack PPTX for XML editing |
| `pack.py` | `~/.claude/skills/pptx/scripts/office/pack.py` | Repack PPTX after editing |

```bash
# Restyle a new diagram
python3 /tmp/brand_style.py /mnt/d/Work/gy/new_diagram.drawio

# Edit a PPTX
python3 ~/.claude/skills/pptx/scripts/office/unpack.py input.pptx unpacked/
# ... edit XML in unpacked/ppt/slides/ ...
python3 ~/.claude/skills/pptx/scripts/office/pack.py unpacked/ output.pptx --original input.pptx
```

---

## Source Documents

| File | Purpose |
|------|---------|
| `基于国内外FICC平台发展经验规划FICC平台建设路径0417.pdf` | Brand source — all colors and styles extracted from here |
| `CEO_Deck_v8.pptx` | Original CEO deck (pre-brand, keep as baseline) |
| `CEO_Deck_v9.pptx` | Brand-styled output |
