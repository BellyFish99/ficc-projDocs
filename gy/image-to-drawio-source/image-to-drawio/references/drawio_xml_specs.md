# Draw.io XML Specification Cheat Sheet

Use these XML snippets to ensure the model generates high-fidelity, editable Draw.io files.

## 1. File Wrapper
Every `.drawio` file (which is just an XML file) must be wrapped in this structure:
```xml
<mxfile host="Electron" modified="2024-01-01T00:00:00.000Z" agent="Gemini-CLI" version="21.0.0" type="device">
  <diagram id="diagram_id" name="Page-1">
    <mxGraphModel dx="1000" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
        <!-- ELEMENTS GO HERE -->
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
```

## 2. Common Shapes
| Shape | XML Style String |
| :--- | :--- |
| **Rectangle** | `rounded=0;whiteSpace=wrap;html=1;` |
| **Rounded Rect** | `rounded=1;whiteSpace=wrap;html=1;` |
| **Diamond** | `rhombus;whiteSpace=wrap;html=1;` |
| **Cylinder (DB)** | `shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;backgroundOutline=1;size=15;` |
| **Actor** | `shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;outlineConnect=0;` |
| **Cloud** | `ellipse;shape=cloud;whiteSpace=wrap;html=1;` |

## 3. Connectors (Arrows)
```xml
<mxCell id="edge_id" value="Label" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;" edge="1" parent="1" source="source_id" target="target_id">
  <mxGeometry relative="1" as="geometry" />
</mxCell>
```

## 4. Positioning Guidelines
- `x`, `y`: Top-left coordinates.
- `width`, `height`: Dimensions.
- Increment `id` for every element starting from "2" (0 and 1 are reserved for the root).
