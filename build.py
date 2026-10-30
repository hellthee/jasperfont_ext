#!/usr/bin/env python3
"""
build.py — build the JasperReports Chinese font extension JAR.

Produces dist/jasperreports-ext-fonts.jar: a JasperReports font extension
bundling a CJK-capable font that renders Simplified Chinese correctly in
PDF (iText) and at a line height consistent with .docx export.

Pipeline
--------
1. Obtain NotoSansCJK-{Regular,Bold}.ttc  (--ttc-dir, or download)
2. Extract face index 2 (Noto Sans CJK SC) from each collection
3. Convert CFF/OpenType outlines -> TrueType `glyf` outlines  (cu2qu)
4. Patch OS/2 metrics so Word matches iText's line height
5. Rename the font family (SIL OFL reserved-name compliance)
6. Package with fonts.xml + jasperreports_extension.properties

Why step 3 is mandatory
-----------------------
noto-cjk ships CFF-flavoured OpenType (sfnt magic 'OTTO'). iText 2.1.7 —
the PDF engine in JasperReports 6.6.0 — cannot embed CFF OpenType with
Identity-H encoding. It raises no error; the glyphs simply do not appear
in the PDF. Converting the cubic CFF charstrings to quadratic TrueType
outlines (sfnt magic 0x00010000, with `glyf` + `loca`) is what makes the
font usable. Conversion tolerance is 1.0 em unit at 1000 UPM, i.e.
visually lossless.

Fonts above 32767 glyphs cannot use `post` format 2.0 (its glyph-name
index is uint16). Noto Sans CJK SC has 65535 glyphs, so `post` is written
as format 3.0 (no glyph names). JasperReports and iText address glyphs by
cmap, not by name, so this is harmless here.

Why step 4 is needed
--------------------
iText derives line height from the `hhea` table: for this font
(ascent 1160, descent -288, UPM 1000) that is 1.448 x font size, i.e.
17.376pt at 12pt. Word, with USE_TYPO_METRICS clear, uses OS/2
sTypoAscender/sTypoDescender, which upstream are 880/-120 -> 12.000pt at
12pt. Aligning the typo metrics to `hhea` and setting USE_TYPO_METRICS
(which requires OS/2 >= v4) makes both engines agree.

Note this fixes the *font's* contribution only. The JasperReports .docx
exporter additionally hardcodes <w:spacing w:lineRule="auto" w:line="240"/>
on every paragraph, which overrides font metrics; see
java/DocxLineSpacingFixer.java.

Why step 5 is required
----------------------
Noto Sans SC is licensed under the SIL Open Font License 1.1. Sections 3
and 5 forbid distributing a *modified* font under its original name.
Steps 3 and 4 modify it, so the family is renamed. The JRXML-facing alias
"Chinese" is defined in fonts.xml and is unaffected.

Usage
-----
    python3 build.py                        # download the TTCs
    python3 build.py --ttc-dir /path/to/ttc # use local TTCs
    python3 build.py --family "My Font"     # override family name

Requires: fonttools >= 4.0  (pip install fonttools)
"""

from __future__ import annotations

import argparse
import io
import os
import shutil
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

try:
    from fontTools.ttLib import TTFont, TTCollection, newTable
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    from fontTools.pens.cu2quPen import Cu2QuPen
except ImportError:
    sys.exit("fonttools is required:  pip install fonttools")

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

REPO = Path(__file__).resolve().parent

JAR_NAME = "jasperreports-ext-fonts.jar"
DIST_DIR = REPO / "dist"
SRC_DIR = REPO / "src"
FONT_PKG = "net/sf/jasperreports/fonts/chinese"

# Default renamed family. Must not be a Noto reserved font name.
DEFAULT_FAMILY = "Noto Sans SC Jasper"

# Face index of "Noto Sans CJK SC" inside NotoSansCJK-*.ttc.
# Collection order: 0 JP, 1 KR, 2 SC, 3 TC, 4 HK, 5-9 Mono variants.
SC_FACE_INDEX = 2

# cu2qu conversion tolerance, in font units at 1000 UPM.
MAX_ERR = 1.0

WEIGHTS = {
    "Regular": "NotoSansCJK-Regular.ttc",
    "Bold": "NotoSansCJK-Bold.ttc",
}

# Upstream release assets (noto-cjk Sans v2.004).
DOWNLOAD_BASE = (
    "https://github.com/notofonts/noto-cjk/releases/download/Sans2.004/"
)
DOWNLOAD_ARCHIVE = "03_NotoSansCJK-OTC.zip"

# Common install locations for the TTCs, checked before downloading.
LOCAL_TTC_DIRS = [
    Path("/usr/share/fonts/opentype/noto"),
    Path("/usr/share/fonts/truetype/noto"),
    Path("/Library/Fonts"),
    Path(os.path.expandvars(r"%WINDIR%\Fonts")) if os.name == "nt" else None,
]


# --------------------------------------------------------------------------
# Font acquisition
# --------------------------------------------------------------------------


def find_local_ttcs(explicit_dir: str | None) -> dict[str, Path] | None:
    """Return {weight: path} if every needed TTC is present locally."""
    candidates = [Path(explicit_dir)] if explicit_dir else [
        d for d in LOCAL_TTC_DIRS if d is not None
    ]
    for d in candidates:
        if not d.is_dir():
            continue
        found = {w: d / fn for w, fn in WEIGHTS.items()}
        if all(p.is_file() for p in found.values()):
            print(f"  using local TTCs from {d}")
            return found
    if explicit_dir:
        sys.exit(
            f"--ttc-dir {explicit_dir!r} does not contain "
            + " and ".join(WEIGHTS.values())
        )
    return None


def download_ttcs(workdir: Path) -> dict[str, Path]:
    """Download and unpack the upstream OTC archive."""
    url = DOWNLOAD_BASE + DOWNLOAD_ARCHIVE
    print(f"  downloading {url}")
    archive = workdir / DOWNLOAD_ARCHIVE
    try:
        with urllib.request.urlopen(url, timeout=300) as r, open(archive, "wb") as f:
            shutil.copyfileobj(r, f)
    except Exception as e:
        sys.exit(
            f"download failed: {e}\n"
            f"Fetch {url} manually, unpack it, and re-run with "
            f"--ttc-dir <dir containing the .ttc files>."
        )

    extract_dir = workdir / "ttc"
    extract_dir.mkdir(exist_ok=True)
    with zipfile.ZipFile(archive) as z:
        for member in z.namelist():
            base = os.path.basename(member)
            if base in WEIGHTS.values():
                with z.open(member) as srcf, open(extract_dir / base, "wb") as dstf:
                    shutil.copyfileobj(srcf, dstf)

    found = {w: extract_dir / fn for w, fn in WEIGHTS.items()}
    missing = [str(p.name) for p in found.values() if not p.is_file()]
    if missing:
        sys.exit(f"archive did not contain: {', '.join(missing)}")
    return found


# --------------------------------------------------------------------------
# Font transformation
# --------------------------------------------------------------------------


def extract_sc_face(ttc_path: Path) -> TTFont:
    """Pull the Simplified Chinese face out of a CJK collection."""
    coll = TTCollection(str(ttc_path))
    font = coll[SC_FACE_INDEX]
    name = font["name"].getDebugName(1)
    if "SC" not in (name or ""):
        sys.exit(
            f"{ttc_path.name} face {SC_FACE_INDEX} is {name!r}, not a "
            f"Simplified Chinese face — collection layout has changed."
        )
    return font


def cff_to_truetype(font: TTFont) -> None:
    """Replace CFF cubic outlines with TrueType quadratic `glyf` outlines."""
    if "glyf" in font:
        return  # already TrueType

    glyph_order = font.getGlyphOrder()
    glyph_set = font.getGlyphSet()

    glyf = newTable("glyf")
    glyf.glyphOrder = glyph_order
    glyf.glyphs = {}
    for name in glyph_order:
        pen = TTGlyphPen(glyph_set)
        glyph_set[name].draw(Cu2QuPen(pen, MAX_ERR, reverse_direction=True))
        glyf[name] = pen.glyph()

    font["glyf"] = glyf
    font["loca"] = newTable("loca")  # populated when glyf compiles

    maxp = font["maxp"]
    maxp.tableVersion = 0x00010000
    maxp.maxZones = 1
    maxp.maxTwilightPoints = 0
    maxp.maxStorage = 0
    maxp.maxFunctionDefs = 0
    maxp.maxInstructionDefs = 0
    maxp.maxStackElements = 0
    maxp.maxSizeOfInstructions = 0
    maxp.maxComponentElements = max(
        (len(g.components) if g.isComposite() else 0) for g in glyf.glyphs.values()
    )

    font["head"].glyphDataFormat = 0
    font["head"].indexToLocFormat = 0
    font.sfntVersion = "\x00\x01\x00\x00"

    for tag in ("CFF ", "VORG"):
        if tag in font:
            del font[tag]

    # post 2.0 indexes glyph names with uint16; CJK fonts exceed 32767 glyphs.
    post = font["post"]
    post.formatType = 3.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = None


def align_metrics(font: TTFont) -> None:
    """Make OS/2 typo metrics match `hhea` and enable USE_TYPO_METRICS."""
    hhea, os2 = font["hhea"], font["OS/2"]

    if os2.version < 4:
        os2.version = 4
        for attr, default in (
            ("sxHeight", 0),
            ("sCapHeight", 0),
            ("usDefaultChar", 0),
            ("usBreakChar", 32),
            ("usMaxContext", 0),
        ):
            if not hasattr(os2, attr):
                setattr(os2, attr, default)

    os2.sTypoAscender = hhea.ascent
    os2.sTypoDescender = hhea.descent
    os2.sTypoLineGap = hhea.lineGap
    os2.usWinAscent = hhea.ascent
    os2.usWinDescent = -hhea.descent  # stored unsigned
    os2.fsSelection |= 0x80  # USE_TYPO_METRICS


def rename_family(font: TTFont, family: str, weight: str) -> None:
    """Rewrite the `name` table so no Noto reserved font name remains."""
    subfamily = weight  # "Regular" | "Bold"
    full = family if subfamily == "Regular" else f"{family} {subfamily}"
    postscript = full.replace(" ", "")

    values = {
        1: family,       # family
        2: subfamily,    # subfamily
        3: f"{full}; built by jasperfont_ext",  # unique id
        4: full,         # full name
        6: postscript,   # PostScript name
        16: family,      # typographic family
        17: subfamily,   # typographic subfamily
    }

    name_table = font["name"]
    for record in list(name_table.names):
        if record.nameID in values:
            name_table.setName(
                values[record.nameID],
                record.nameID,
                record.platformID,
                record.platEncID,
                record.langID,
            )


def build_font(ttc_path: Path, weight: str, family: str, out_path: Path) -> None:
    print(f"  {weight}: extracting face {SC_FACE_INDEX} from {ttc_path.name}")
    font = extract_sc_face(ttc_path)

    print(f"  {weight}: converting {len(font.getGlyphOrder())} glyphs CFF -> TrueType")
    cff_to_truetype(font)

    align_metrics(font)
    rename_family(font, family, weight)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    font.save(str(out_path))
    print(f"  {weight}: wrote {out_path.name} ({out_path.stat().st_size / 1e6:.1f} MB)")


def verify(path: Path) -> None:
    """Fail loudly if the output is not a usable TrueType CJK font."""
    magic = path.open("rb").read(4)
    if magic != b"\x00\x01\x00\x00":
        sys.exit(f"{path.name}: sfnt magic {magic!r}, expected TrueType")

    font = TTFont(str(path))
    tables = set(font.reader.keys())
    for required in ("glyf", "loca", "cmap", "head", "hhea", "hmtx", "maxp", "OS/2"):
        if required not in tables:
            sys.exit(f"{path.name}: missing required table {required!r}")

    cmap = font.getBestCmap()
    for ch in "中文汇款甲乙":
        if ord(ch) not in cmap:
            sys.exit(f"{path.name}: no glyph for U+{ord(ch):04X} ({ch})")

    glyph = font["glyf"][cmap[0x4E2D]]  # 中
    if glyph.isComposite() or not len(glyph.coordinates):
        sys.exit(f"{path.name}: U+4E2D has no outline points")

    hhea, os2 = font["hhea"], font["OS/2"]
    if (os2.sTypoAscender, os2.sTypoDescender) != (hhea.ascent, hhea.descent):
        sys.exit(f"{path.name}: OS/2 typo metrics not aligned to hhea")
    if not os2.fsSelection & 0x80:
        sys.exit(f"{path.name}: USE_TYPO_METRICS not set")
    if "Noto Sans CJK" in (font["name"].getDebugName(1) or ""):
        sys.exit(f"{path.name}: family still carries a Noto reserved name")

    cjk = sum(1 for cp in cmap if 0x4E00 <= cp <= 0x9FFF)
    print(
        f"  {path.name}: OK — TrueType, {cjk} CJK glyphs, "
        f"hhea {hhea.ascent}/{hhea.descent}, USE_TYPO_METRICS set"
    )


# --------------------------------------------------------------------------
# Packaging
# --------------------------------------------------------------------------


def package(ttf_paths: dict[str, Path], jar_path: Path) -> None:
    jar_path.parent.mkdir(parents=True, exist_ok=True)

    manifest = (
        "Manifest-Version: 1.0\r\n"
        "Created-By: jasperfont_ext build.py\r\n"
        "Implementation-Title: JasperReports Chinese Font Extension\r\n"
        "\r\n"
    )

    with zipfile.ZipFile(jar_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("META-INF/MANIFEST.MF", manifest)
        z.write(SRC_DIR / "jasperreports_extension.properties",
                "jasperreports_extension.properties")
        z.write(SRC_DIR / FONT_PKG / "fonts.xml", f"{FONT_PKG}/fonts.xml")
        for weight, p in ttf_paths.items():
            z.write(p, f"{FONT_PKG}/NotoSansSC-{weight}.ttf")

    print(f"\n  {jar_path.relative_to(REPO)}  ({jar_path.stat().st_size / 1e6:.1f} MB)")


def check_family_matches_xml(family: str) -> None:
    """Warn if fonts.xml does not declare the family being built."""
    xml = (SRC_DIR / FONT_PKG / "fonts.xml").read_text(encoding="utf-8")
    if f'"{family}"' not in xml:
        print(
            f"  WARNING: fonts.xml does not declare fontFamily name=\"{family}\".\n"
            f"           JRXMLs referencing that family will not resolve."
        )


# --------------------------------------------------------------------------


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--ttc-dir", help="directory holding NotoSansCJK-*.ttc")
    ap.add_argument("--family", default=DEFAULT_FAMILY,
                    help=f"font family name (default: {DEFAULT_FAMILY!r})")
    ap.add_argument("--keep-ttf", action="store_true",
                    help="also leave the converted .ttf files in dist/")
    args = ap.parse_args()

    print("JasperReports Chinese font extension\n")
    check_family_matches_xml(args.family)

    with tempfile.TemporaryDirectory(prefix="jasperfont_") as tmp:
        tmpdir = Path(tmp)

        print("[1/3] source fonts")
        ttcs = find_local_ttcs(args.ttc_dir) or download_ttcs(tmpdir)

        print("\n[2/3] convert, patch, rename")
        out_dir = DIST_DIR if args.keep_ttf else tmpdir / "ttf"
        ttfs = {}
        for weight, ttc in ttcs.items():
            dest = out_dir / f"NotoSansSC-{weight}.ttf"
            build_font(ttc, weight, args.family, dest)
            verify(dest)
            ttfs[weight] = dest

        print("\n[3/3] package")
        package(ttfs, DIST_DIR / JAR_NAME)

    print("\nDone.")


if __name__ == "__main__":
    main()
