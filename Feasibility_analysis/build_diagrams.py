"""
build_diagrams.py
=================
Reads ficc_data.yaml and generates one draw.io XML file per module that has
`function_domains`. Output files go to Feasibility_analysis/drawio/ named
detail_{no:02d}_{cn}.drawio (e.g. detail_07_组合管理+绩效归因.drawio).

Can be run standalone or imported by build_all.py:
    from build_diagrams import main as build_diagrams_main
"""

import html
import os
import yaml

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
YAML_PATH  = os.path.join(SCRIPT_DIR, 'ficc_data.yaml')
OUT_DIR    = os.path.join(SCRIPT_DIR, 'drawio')

# ---------------------------------------------------------------------------
# Layout constants
# ---------------------------------------------------------------------------
CANVAS_W = 1100   # total diagram width
HEADER_H = 44     # system title bar height
LAYER_H  = 32     # data_inputs / foundation strip height
DOMAIN_H = 36     # domain header row height
FUNC_H   = 28     # each function row height
X0, Y0   = 10, 10 # top-left origin

# ---------------------------------------------------------------------------
# Archforce brand styles
# ---------------------------------------------------------------------------
STYLE_HEADER = (
    "rounded=0;whiteSpace=wrap;html=1;"
    "fillColor=#0F2060;fontColor=#FFFFFF;strokeColor=#0F2060;"
    "fontStyle=1;fontSize=12;align=center;verticalAlign=middle;"
)
STYLE_DATA_LAYER = (
    "rounded=0;whiteSpace=wrap;html=1;"
    "fillColor=#F5F7FA;fontColor=#1B3275;strokeColor=#CCCCCC;"
    "fontSize=10;align=left;verticalAlign=middle;"
)
STYLE_DOMAIN = (
    "rounded=0;whiteSpace=wrap;html=1;"
    "fillColor=#2B5EC7;fontColor=#FFFFFF;strokeColor=#1B3275;"
    "fontStyle=1;fontSize=10;align=center;verticalAlign=middle;"
)
STYLE_FUNC = (
    "rounded=1;whiteSpace=wrap;html=1;"
    "fillColor=#D6E4F7;fontColor=#1B3275;strokeColor=#1B3275;"
    "fontSize=9;align=center;verticalAlign=middle;"
)
STYLE_FOUNDATION = (
    "rounded=0;whiteSpace=wrap;html=1;"
    "fillColor=#1B3275;fontColor=#FFFFFF;strokeColor=#1B3275;"
    "fontSize=10;align=left;verticalAlign=middle;"
)
STYLE_EMPTY = (
    "rounded=1;whiteSpace=wrap;html=1;"
    "fillColor=#F5F7FA;fontColor=#F5F7FA;strokeColor=#D6E4F7;"
    "fontSize=9;"
)

# ---------------------------------------------------------------------------
# Cell helper
# ---------------------------------------------------------------------------

def make_cell(cid, value, style, x, y, w, h):
    """Return a dict representing one mxCell."""
    return dict(id=cid, value=value, style=style, x=x, y=y, w=w, h=h)


# ---------------------------------------------------------------------------
# XML builder
# ---------------------------------------------------------------------------

def build_drawio(cells, canvas_w, total_h):
    """Render list of cell-dicts to a draw.io XML string."""
    pw = canvas_w + 20
    ph = total_h + 20
    lines = []
    lines.append(
        f'<mxGraphModel dx="1422" dy="762" grid="0" gridSize="10" guides="1" '
        f'tooltips="1" connect="1" arrows="1" fold="1" page="0" pageScale="1" '
        f'pageWidth="{pw}" pageHeight="{ph}" math="0" shadow="0">'
    )
    lines.append('  <root>')
    lines.append('    <mxCell id="0"/>')
    lines.append('    <mxCell id="1" parent="0"/>')
    for c in cells:
        val = html.escape(str(c['value']))
        lines.append(
            f'    <mxCell id="{c["id"]}" value="{val}" style="{c["style"]}" '
            f'vertex="1" parent="1">'
        )
        lines.append(
            f'      <mxGeometry x="{c["x"]:.1f}" y="{c["y"]:.1f}" '
            f'width="{c["w"]:.1f}" height="{c["h"]}" as="geometry"/>'
        )
        lines.append('    </mxCell>')
    lines.append('  </root>')
    lines.append('</mxGraphModel>')
    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# Per-module diagram builder
# ---------------------------------------------------------------------------

def build_module_diagram(m):
    """
    Build the draw.io XML for one module.
    Returns None if the module has no function_domains.
    """
    domains = m.get('function_domains')
    if not domains:
        return None

    n_domains = len(domains)
    col_w     = CANVAS_W / n_domains
    max_funcs = max(len(d.get('functions', [])) for d in domains)

    # Total diagram height
    total_h   = HEADER_H
    has_data  = bool(m.get('data_inputs'))
    has_found = bool(m.get('foundation_layer'))
    if has_data:  total_h += LAYER_H
    total_h += DOMAIN_H + max_funcs * FUNC_H
    if has_found: total_h += LAYER_H

    cell_id = 2   # ids 0 and 1 are reserved by draw.io
    cells   = []

    # 1. System header (full width)
    header_text = (
        f"{m['cn']}  |  {m['priority']}  |  "
        f"{m['duration_m']}月  |  {m['team_sz']}人"
    )
    cells.append(make_cell(cell_id, header_text, STYLE_HEADER,
                           X0, Y0, CANVAS_W, HEADER_H))
    cell_id += 1
    cur_y = Y0 + HEADER_H

    # 2. Data inputs strip (optional)
    if has_data:
        text = "数据层：" + " · ".join(m['data_inputs'])
        cells.append(make_cell(cell_id, text, STYLE_DATA_LAYER,
                               X0, cur_y, CANVAS_W, LAYER_H))
        cell_id += 1
        cur_y += LAYER_H

    # 3. Domain header row
    domain_y = cur_y
    for i, domain in enumerate(domains):
        x     = X0 + i * col_w
        label = f"{domain['id']}. {domain['name']}"
        cells.append(make_cell(cell_id, label, STYLE_DOMAIN,
                               x, domain_y, col_w, DOMAIN_H))
        cell_id += 1
    cur_y = domain_y + DOMAIN_H

    # 4. Function rows — pad shorter columns with invisible placeholders
    for row_idx in range(max_funcs):
        for i, domain in enumerate(domains):
            x     = X0 + i * col_w
            funcs = domain.get('functions', [])
            if row_idx < len(funcs):
                fn = funcs[row_idx]
                cells.append(make_cell(cell_id, fn['name'], STYLE_FUNC,
                                       x, cur_y, col_w, FUNC_H))
            else:
                cells.append(make_cell(cell_id, "", STYLE_EMPTY,
                                       x, cur_y, col_w, FUNC_H))
            cell_id += 1
        cur_y += FUNC_H

    # 5. Foundation strip (optional)
    if has_found:
        text = "量化底座：" + " · ".join(m['foundation_layer'])
        cells.append(make_cell(cell_id, text, STYLE_FOUNDATION,
                               X0, cur_y, CANVAS_W, LAYER_H))
        cell_id += 1

    return build_drawio(cells, CANVAS_W, total_h)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    with open(YAML_PATH, encoding='utf-8') as f:
        data = yaml.safe_load(f)

    generated = []
    skipped   = []

    for m in data['modules']:
        if not m.get('function_domains'):
            skipped.append(m['cn'])
            continue

        xml_str  = build_module_diagram(m)
        no       = m['no']
        cn       = m['cn']
        filename = f"detail_{no:02d}_{cn}.drawio"
        out_path = os.path.join(OUT_DIR, filename)

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(xml_str)

        generated.append(filename)
        print(f"  [OK] {filename}")

    print(f"\nGenerated {len(generated)} diagram(s).")
    if skipped:
        print(f"Skipped (no function_domains): {', '.join(skipped)}")

    return generated


if __name__ == '__main__':
    main()
