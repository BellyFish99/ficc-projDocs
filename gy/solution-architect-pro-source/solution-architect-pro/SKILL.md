---
name: solution-architect-pro
description: A comprehensive framework for solution architects to transform client requirements into technical skeletons and executive-level CEO presentations. Use when building complex solution proposals, strategic planning documents, or multi-asset platform designs (e.g., FICC, POMS).
---

# Solution Architect Pro (SSoT Edition)

## Overview
This skill codifies the senior solution architecture workflow using a **Single Source of Truth (SSoT)** model. It transforms raw requirements into a unified **Master Manifest** that drives both the long-form solution document and the high-impact **CEO Deck** simultaneously, ensuring 100% synchronization.

## Workflow: The SSoT Manifest Model

### 1. Discovery & Value Matrix
- **Action**: Analyze raw requests and map them to `key_values.txt`.
- **Constraint**: Every strategic goal must have a corresponding "Capability Leap" defined in the Master Manifest.

### 2. The Master Manifest (`solution_master_v[N].md`)
- **Structure**:
    - **YAML Frontmatter**: Define all project variables (AUM, ROI, Client Name, Target Yield) here. These act as the "Single Source of Truth."
    - **Long-form Narrative**: Detailed technical and business analysis for the final proposal.
    - **PPT Slide Blocks**: Embed `<!-- @slide:start -->` and `<!-- @slide:end -->` tags within the narrative to define exact PPT content.
- **Rule**: Never update a number in the text; update it in the YAML frontmatter.

### 3. Automated Asset Generation
- **PPT Generation**: Use the internal script `scripts/ppt_engine.py` to extract slides from the Master Manifest and generate the `.pptx`.
- **Logic**: The script automatically injects YAML variables into the slides and renders them using the professional McKinsey template.

### 4. Refinement Loop
- One change to the Master Manifest = All artifacts (Doc, PPT, ROI tables) stay in sync.
- Increment version numbers for every major strategy shift (e.g., v8 -> v9).

## Core Capabilities

### 1. Variable-Driven Architecture
The skill uses dynamic variables (e.g., `{{metrics.aum}}`) to ensure consistency. 
- **ROI Calculation**: If you adjust the "AUM" variable, all ROI estimates across the Doc and the Deck must be recalculated.

### 2. Executive "So What?" Filter
Content within `<!-- @slide -->` tags must strictly follow:
- **Action Titles**: Conclusions only (e.g., "Real-time Risk is the Survival Baseline").
- **MECE Logic**: Mutually Exclusive, Collectively Exhaustive points.
- **Speaker Notes**: Clear, concise cues for the presenter.

## Resources & Internal Tools
- **scripts/ppt_engine.py**: The unified engine for PPT generation.
- **assets/master_template.md**: Boilerplate for a new SSoT Master Manifest.
- **references/mckinsey_style.md**: Standard rules for high-impact titles and insights.
