# Project Checkpoint — 华锐技术 (Archforce Technology) GY

**Last updated:** 2026-04-28  
**Working directory:** `D:\work\FICC`  
**Git remote:** `git@github.com:BellyFish99/ficc-projDocs.git` (origin/main)  
**Latest commit:** `00e7587` — Replace S14 diagram: swap 1+16架构 for 功能全景图

---

## Session 2026-04-28 — PPT Revision, Diagram Embedding & GitHub Push

### What We Did

#### 1. PPT Text Revision (`gy/revise_ppt.py`)
Updated `gy/gy_ppt.pptx` (43 slides) to align with the 24-month roadmap:
- **S10:** Calypso/Murex 信创适配 marked "✗ 失格"; Takeaway updated with "综合评分 9/9"
- **S12:** "18个月黄金建设窗口" → "24个月黄金建设窗口，M6 Quick Win验收节点确认价值"
- **S33:** Phase headers updated — Phase 1 (M1–M12 筑基期), Phase 2 (M13–M18 赋能期), Phase 3 (M19–M24 进化期); M6 Quick Win milestone added
- **S34:** "Phase 1灯塔效应" → "M6 Quick Win效应" throughout

#### 2. Doc/PPT Sync Check (`gy/gy_solution.md`)
- Fixed §0 summary line 82: "3期18个月" → "3期24个月，M6 Quick Win验收节点"
- Fixed §0 line 658: `Phase 2（{{total_month}}月）` → `Phase 2（M13–M18）` (broken template variable)

#### 3. Diagram Embedding (`gy/embed_diagrams_ppt.py` + `gy/drawio_renderer.py`)
Built a custom drawio → PNG renderer (cairosvg + WenQuanYi Zen Hei font) since no drawio CLI is installed.
Embedded 7 diagrams into PPT:

| Slide | Diagram | Placement |
|-------|---------|-----------|
| S10 | `diagram_4_竞争定位` | (2.17", 1.10") 8.99"×5.26" |
| S14 | `diagram_A_功能全景图` | (2.83", 1.30") 7.67"×5.20" |
| S30 | `diagram_7_CEP风控引擎` | (0.77", 1.10") 11.78"×4.45" |
| S31 | `diagram_8_指令执行泳道` | (0.22", 1.10") 12.89"×4.55" |
| S33 | `diagram_6_实施路线图` | (2.81", 1.10") 7.72"×5.00" |
| S39 | `diagram_3_924行情推演` | (0.00", 1.10") 13.33"×4.28" |
| S40 | `diagram_5_ROI决策框架` | (1.39", 1.10") 10.56"×4.28" |

**Note:** S14 originally used `diagram_H_1+16架构`. Replaced with `diagram_A_功能全景图` (6-domain capability panorama) — more CEO-friendly as a platform intro.

#### 4. GitHub Remote Setup
- Added remote: `git@github.com:BellyFish99/ficc-projDocs.git`
- All commits pushed to `origin/main`

### Key Files
| File | Purpose |
|------|---------|
| `gy/gy_ppt.pptx` | Main deliverable — 43-slide CEO deck with embedded diagrams |
| `gy/gy_solution.md` | Solution document (source of truth) |
| `gy/revise_ppt.py` | Text revision script (run once, idempotent from backup) |
| `gy/embed_diagrams_ppt.py` | Diagram embedding script — always restores from `gy/backup/gy_ppt_pre_diagrams.pptx` first |
| `gy/drawio_renderer.py` | Custom drawio → SVG → PNG renderer |
| `gy/backup/gy_ppt_pre_diagrams.pptx` | Clean PPT snapshot (text revised, no diagrams) — used as base for re-runs |
| `gy/tmp_diagrams/` | Rendered PNG cache (7 files) |
| `gy/drawio/` | All 15 drawio source diagrams |

### Remaining Diagrams Not Yet Embedded
The following diagrams exist but are not currently placed in the PPT:

| Diagram | Content | Potential slide |
|---------|---------|-----------------|
| `diagram_H_1+16架构` | 1+16 engine structure (配得优/算得快/控得稳/连得通) | Could replace or complement S14 |
| `diagram_1_投资经理的一天` | Investment manager daily workflow | S34 (M6 Quick Win day-in-life) |
| `diagram_B_系统组件图` | System component diagram | Technical appendix |
| `diagram_C_集成架构图` | Integration architecture | Technical appendix |
| `diagram_D_需求追溯矩阵` | Requirements traceability matrix | Appendix |
| `diagram_E_数据流图` | Data flow | Technical appendix |
| `diagram_F_信创技术栈` | 信创 domestic tech stack | S8 or S11 |
| `diagram_G_部署架构` | Deployment architecture | Technical appendix |

### How to Re-run Diagram Embedding
If diagrams or placements need updating:
```bash
# Always restore from clean backup first to avoid double-embedding
cp gy/backup/gy_ppt_pre_diagrams.pptx gy/gy_ppt.pptx
python3 gy/embed_diagrams_ppt.py
```

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
