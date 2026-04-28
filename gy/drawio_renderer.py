"""
Minimal drawio → PNG renderer for brand-consistent diagrams.
Handles: rectangles, ellipses, HTML-formatted text, parent offsets.
"""

import re
import html as html_mod
import xml.etree.ElementTree as ET
import cairosvg

FONT_FAMILY = "WenQuanYi Zen Hei,Arial,sans-serif"
PADDING = 8


def parse_style(s):
    d = {}
    for part in (s or "").split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            d[k.strip()] = v.strip()
    return d


def strip_html(text):
    """Strip HTML tags, decode entities, preserve line-breaks."""
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</p>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = html_mod.unescape(text)
    return text.strip()


def get_parent_abs(cells, pid, depth=0):
    if depth > 10 or pid in ("0", "1", None):
        return 0.0, 0.0
    p = cells.get(pid)
    if p is None:
        return 0.0, 0.0
    geo = p.find("mxGeometry")
    if geo is None:
        return 0.0, 0.0
    px = float(geo.get("x", 0) or 0)
    py = float(geo.get("y", 0) or 0)
    gpx, gpy = get_parent_abs(cells, p.get("parent", "1"), depth + 1)
    return px + gpx, py + gpy


def esc(text):
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def svg_text_block(lines, tx, ty, w, h, font_size, font_color, font_weight,
                   anchor, valign, spacing_top, spacing_left, spacing_right):
    if not lines:
        return ""
    lh = font_size * 1.35
    total_h = len(lines) * lh

    if valign == "top":
        start_y = ty + spacing_top + font_size
    elif valign == "bottom":
        start_y = ty + h - total_h + font_size
    else:
        start_y = ty + h / 2 - total_h / 2 + font_size

    if anchor == "start":
        text_x = tx + spacing_left
    elif anchor == "end":
        text_x = tx + w - spacing_right
    else:
        text_x = tx + w / 2

    parts = []
    for i, line in enumerate(lines):
        yl = start_y + i * lh
        parts.append(
            f'<text x="{text_x:.1f}" y="{yl:.1f}" '
            f'font-family="{FONT_FAMILY}" font-size="{font_size:.0f}" '
            f'font-weight="{font_weight}" fill="{font_color}" '
            f'text-anchor="{anchor}">{esc(line)}</text>'
        )
    return "\n".join(parts)


def render_to_svg(drawio_path):
    tree = ET.parse(drawio_path)
    root = tree.getroot()
    cells = {c.get("id"): c for c in root.findall(".//mxCell")}

    rects = []
    for c in cells.values():
        if c.get("vertex") != "1":
            continue
        geo = c.find("mxGeometry")
        if geo is None:
            continue
        lx = float(geo.get("x", 0) or 0)
        ly = float(geo.get("y", 0) or 0)
        w = float(geo.get("width", 0) or 0)
        h = float(geo.get("height", 0) or 0)
        if w == 0 and h == 0:
            continue
        ox, oy = get_parent_abs(cells, c.get("parent", "1"))
        rects.append((lx + ox, ly + oy, w, h, c.get("style", ""), c.get("value", "")))

    if not rects:
        return None, 0, 0

    min_x = min(r[0] for r in rects)
    min_y = min(r[1] for r in rects)
    max_x = max(r[0] + r[2] for r in rects)
    max_y = max(r[1] + r[3] for r in rects)

    vw = max_x - min_x + 2 * PADDING
    vh = max_y - min_y + 2 * PADDING

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{vw:.0f}" height="{vh:.0f}">',
        f'<rect width="{vw:.0f}" height="{vh:.0f}" fill="#F5F7FA"/>',
    ]

    for ax, ay, w, h, style_str, value in rects:
        st = parse_style(style_str)
        rx = ax - min_x + PADDING
        ry = ay - min_y + PADDING

        fill = st.get("fillColor", "#FFFFFF")
        stroke = st.get("strokeColor", "#000000")
        sw = float(st.get("strokeWidth", "1") or "1")
        fc = st.get("fontColor", "#000000")
        fs = float(st.get("fontSize", "11") or "11")
        fsty = int(st.get("fontStyle", "0") or "0")
        bold = (fsty & 1) != 0
        align = st.get("align", "center")
        valign = st.get("verticalAlign", "middle")
        rounded = st.get("rounded", "0") == "1"
        is_ellipse = "ellipse" in style_str

        fill_a = "none" if fill in ("none", "") else fill
        stroke_a = "none" if stroke in ("none", "") else stroke
        rr = "4" if rounded else "0"
        fw = "bold" if bold else "normal"
        anchor = {"left": "start", "right": "end"}.get(align, "middle")
        sl = float(st.get("spacingLeft", "4") or "4")
        sr = float(st.get("spacingRight", "4") or "4")
        st_sp = float(st.get("spacingTop", "4") or "4")

        if is_ellipse:
            cx, cy = rx + w / 2, ry + h / 2
            parts.append(
                f'<ellipse cx="{cx:.1f}" cy="{cy:.1f}" rx="{w/2:.1f}" ry="{h/2:.1f}" '
                f'fill="{fill_a}" stroke="{stroke_a}" stroke-width="{sw:.1f}"/>'
            )
        else:
            parts.append(
                f'<rect x="{rx:.1f}" y="{ry:.1f}" width="{w:.1f}" height="{h:.1f}" '
                f'fill="{fill_a}" stroke="{stroke_a}" stroke-width="{sw:.1f}" rx="{rr}"/>'
            )

        if value and st.get("noLabel") != "1":
            plain = strip_html(value)
            lines = [l for l in plain.split("\n") if l.strip()]
            if lines:
                parts.append(svg_text_block(
                    lines, rx, ry, w, h, fs, fc, fw,
                    anchor, valign, st_sp, sl, sr
                ))

    parts.append("</svg>")
    return "\n".join(parts), vw, vh


def convert(drawio_path, output_path, target_w_px=1300, target_h_px=None):
    """Render drawio file to PNG. Returns (actual_w, actual_h) in pixels."""
    svg_str, vw, vh = render_to_svg(drawio_path)
    if svg_str is None:
        raise ValueError(f"No renderable content in {drawio_path}")
    if target_h_px is None:
        target_h_px = round(target_w_px * vh / vw)
    png = cairosvg.svg2png(
        bytestring=svg_str.encode("utf-8"),
        output_width=target_w_px,
        output_height=target_h_px,
    )
    with open(output_path, "wb") as f:
        f.write(png)
    return target_w_px, target_h_px
