"""
ë°ì´í„° ëª¨ë¸ ì •ì˜
"""
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from utils import clean_text


@dataclass
class BomRow:
    category: str
    product: str
    material_name: str
    supplier_article_number: str
    usage: str
    quality_details: str
    supplier: str
    # key: formatted color header text (multi-line), value: color cell value
    colors: Dict[str, str]
    # PNG bytes for BOM Details 'Image' column (optional)
    image_png: Optional[bytes] = None
    # PNG bytes for color/print/graphic thumbnails inside color columns (Graphic section)
    color_images: Dict[str, bytes] = field(default_factory=dict)


def section_from_cell_text(s: str) -> Optional[str]:
    """
    Detect section header like 'Fabric (5)', 'Trim (6)', 'Graphic (1)', 'Packaging and Labels (10)'.
    Must match the section header pattern, not merely contain the word.
    """
    t = clean_text(s).lower()
    patterns = [
        (r"^fabric\s*\(\d+\)\s*$", "Fabric"),
        (r"^trim\s*\(\d+\)\s*$", "Trim"),
        (r"^graphic\s*\(\d+\)\s*$", "Graphic"),
        (r"^packaging\s+and\s+labels\s*\(\d+\)\s*$", "Packaging and Labels"),
        (r"^wash\s*\(\d+\)\s*$", "Wash"),
    ]
    for rx, name in patterns:
        if re.match(rx, t, flags=re.IGNORECASE):
            return name
    return None


def group_rows_by_material(rows: List[BomRow]) -> List[BomRow]:
    """
    Same material -> one row; colors spread to the right as separate columns.
    key = Category + Product + Material Name + Supplier Article Number + Usage + Quality Details + Supplier
    """
    from typing import Tuple
    grouped: Dict[Tuple[str, str, str, str, str, str, str], BomRow] = {}
    order: List[Tuple[str, str, str, str, str, str, str]] = []

    for r in rows:
        key = (
            r.category,
            r.product,
            r.material_name,
            r.supplier_article_number,
            r.usage,
            r.quality_details,
            r.supplier,
        )
        if key not in grouped:
            grouped[key] = BomRow(
                category=r.category,
                product=r.product,
                material_name=r.material_name,
                supplier_article_number=r.supplier_article_number,
                usage=r.usage,
                quality_details=r.quality_details,
                supplier=r.supplier,
                colors={},
                image_png=None,
                color_images={},
            )
            order.append(key)

        for h, v in (r.colors or {}).items():
            if not h:
                continue
            if h not in grouped[key].colors:
                grouped[key].colors[h] = v
            else:
                if (not grouped[key].colors[h]) and v:
                    grouped[key].colors[h] = v

        # image: keep first non-empty
        if grouped[key].image_png is None and getattr(r, "image_png", None):
            grouped[key].image_png = r.image_png

        # graphic color images: merge missing ones
        if getattr(r, "color_images", None):
            if grouped[key].color_images is None:
                grouped[key].color_images = {}
            for hk, bv in r.color_images.items():
                if hk and bv and hk not in grouped[key].color_images:
                    grouped[key].color_images[hk] = bv

    return [grouped[k] for k in order]
