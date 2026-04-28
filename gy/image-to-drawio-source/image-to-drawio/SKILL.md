---
name: image-to-drawio
description: Converts images of diagrams (flowcharts, architecture, sequence) into editable .drawio XML files. Use when a user provides a screenshot or image of a diagram and wants to edit it in Diagrams.net (Draw.io).
---

# Image to Draw.io

## Overview
This skill leverages vision capabilities to decompose a diagram image into its constituent parts (nodes, edges, labels) and reassembles them into a valid Draw.io XML format.

## Workflow

### 1. Vision Analysis
- Analyze the input image provided by the user.
- **Identify Nodes**: List every box, circle, or shape. Note its text, color, and approximate position.
- **Identify Edges**: List every arrow/line connecting nodes. Note the source node, target node, and any edge labels.
- **Categorize Shapes**: Map visual shapes to Draw.io types (e.g., Diamond -> Rhombus/Decision).

### 2. ID Mapping
- Create a mapping table where every visual element gets a unique integer ID (starting from 2).

### 3. XML Generation
- Use the structure defined in [references/drawio_xml_specs.md](references/drawio_xml_specs.md) to wrap the elements.
- For every node, create an `<mxCell>` with the appropriate `style` and `geometry`.
- For every edge, create an `<mxCell>` with `edge="1"` and the correct `source` and `target` IDs.

### 4. File Output
- Output the raw XML.
- If directed, use `write_file` to save the content as `[filename].drawio`.

## Guidelines
- **Precision**: Maintain the relative positions and sizes of shapes from the image.
- **Styling**: If the image has a specific color theme (e.g., Corporate Blue), try to match the hex codes in the XML `fillColor` and `strokeColor`.
- **Labels**: Ensure text inside boxes is accurately transcribed and centered.

## Example Request
"I'm attaching a screenshot of our legacy trade flow. Can you convert this to a .drawio file so I can modify it for the new POMS proposal?"
