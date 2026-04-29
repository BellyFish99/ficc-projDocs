"""Fix corrupted/non-Windows fonts in gy_ppt.pptx.

Replaces fonts unavailable on Windows with standard equivalents:
  Chinese:  MiSans Normal/Medium, Noto Sans SC, PingFang SC, ui-sans-serif → 微软雅黑
  Latin:    Inter UI → Calibri
  Latin:    Helvetica Neue Medium, Helvetica Neue, Helvetica → Arial
"""

import os
import re
import shutil
import subprocess

PPTX = "/mnt/d/work/ficc/gy/gy_ppt.pptx"
UNPACK_DIR = "/tmp/gy_fix_fonts_unpacked"
OUTPUT_PPTX = "/mnt/d/work/ficc/gy/gy_ppt.pptx"
BACKUP_PPTX = "/mnt/d/work/ficc/gy/backup/gy_ppt_pre_fontfix.pptx"

FONT_MAP = {
    "MiSans Normal":         "微软雅黑",
    "MiSans Medium":         "微软雅黑",
    "MiSans":                "微软雅黑",
    "Noto Sans SC":          "微软雅黑",
    "PingFang SC Regular":   "微软雅黑",
    "PingFang SC":           "微软雅黑",
    "ui-sans-serif":         "微软雅黑",
    "Inter UI":              "Calibri",
    "Helvetica Neue Medium": "Arial",
    "Helvetica Neue":        "Arial",
    "Helvetica":             "Arial",
}

UNPACK_SCRIPT = os.path.expanduser("~/.claude/skills/pptx/scripts/office/unpack.py")
PACK_SCRIPT   = os.path.expanduser("~/.claude/skills/pptx/scripts/office/pack.py")


def fix_xml_file(path: str) -> int:
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    original = content
    count = 0
    for bad_font, good_font in FONT_MAP.items():
        pattern = f'typeface="{re.escape(bad_font)}"'
        replacement = f'typeface="{good_font}"'
        new_content, n = re.subn(pattern, replacement, content)
        count += n
        content = new_content

    if content != original:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
    return count


def main():
    os.makedirs(os.path.dirname(BACKUP_PPTX), exist_ok=True)
    shutil.copy2(PPTX, BACKUP_PPTX)
    print(f"Backup saved: {BACKUP_PPTX}")

    if os.path.exists(UNPACK_DIR):
        shutil.rmtree(UNPACK_DIR)
    subprocess.run(["python", UNPACK_SCRIPT, PPTX, UNPACK_DIR], check=True)

    total = 0
    changed_files = []
    dirs_to_scan = [
        os.path.join(UNPACK_DIR, "ppt", "slides"),
        os.path.join(UNPACK_DIR, "ppt", "slideMasters"),
        os.path.join(UNPACK_DIR, "ppt", "slideLayouts"),
        os.path.join(UNPACK_DIR, "ppt", "theme"),
    ]
    for d in dirs_to_scan:
        if not os.path.isdir(d):
            continue
        for fname in sorted(os.listdir(d)):
            if not fname.endswith(".xml"):
                continue
            fpath = os.path.join(d, fname)
            n = fix_xml_file(fpath)
            if n > 0:
                changed_files.append((fname, n))
                total += n

    print(f"\nFixed {total} font references across {len(changed_files)} files:")
    for fname, n in changed_files:
        print(f"  {fname}: {n} replacements")

    subprocess.run(
        ["python", PACK_SCRIPT, UNPACK_DIR, OUTPUT_PPTX, "--original", PPTX],
        check=True,
    )
    print(f"\nOutput written: {OUTPUT_PPTX}")

    # verify
    remaining = subprocess.run(
        ["grep", "-rl",
         "MiSans\|Noto Sans SC\|PingFang\|ui-sans-serif\|Inter UI\|Helvetica",
         os.path.join(UNPACK_DIR, "ppt", "slides")],
        capture_output=True, text=True,
    ).stdout.strip()
    if remaining:
        print(f"\nWARN: remaining occurrences in slides:\n{remaining[:500]}")
    else:
        print("Verified: no remaining corrupted fonts in slides.")


if __name__ == "__main__":
    main()
