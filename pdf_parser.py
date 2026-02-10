"""
PDF íŒŒì‹± ëª¨ë“ˆ
- Master ì •ë³´ ì¶”ì¶œ (Design Number, BOM Number ë“±)
- BOM Details í…Œì´ë¸” í–‰ ì¶”ì¶œ
- BOMColorMatrix ì»¬ëŸ¬ í—¤ë” ì¶”ì¶œ

ì»¬ëŸ¬ê°€ ë§Žì„ ê²½ìš° BOM Details í…Œì´ë¸”ì´ ê°€ë¡œë¡œ í• ë˜ì–´ ì—¬ëŸ¬ íŽ˜ì´ì§€ì— ê±¸ì¹¨:
  Page N  : Product~Supplier + ì²« Nê°œ ì»¬ëŸ¬  (full-header table)
  Page N+1: ì¶”ê°€ ì»¬ëŸ¬ë“¤                     (color continuation table)
  Page N+2: ì¶”ê°€ ì»¬ëŸ¬ë“¤ + Comment           (color continuation table)
ì´ ëª¨ë“ˆì€ ì´ëŸ° ë¶„í• ì„ ì˜¬ë°”ë¥´ê²Œ ì²˜ë¦¬í•¨.
"""
import re
from typing import Dict, List, Optional, Tuple

import pdfplumber

from utils import clean_text, normalize_header, clean_text_keep_newlines, format_color_header_text
from models import BomRow, section_from_cell_text
from image_handler import extract_bom_image_map_from_pdf, extract_graphic_color_cell_images_from_pdf, extract_continuation_graphic_images


def parse_master_from_pdf(pdf_path: str) -> Dict[str, str]:
    """
    Extract master fields from PDF text using label-based regex.
    """
    with pdfplumber.open(pdf_path) as pdf:
        target_text = ""
        first_page_text = ""
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if i == 0:
                first_page_text = t
            if "Design Number" in t and "BOM Number" in t:
                target_text = t
                break
        if not target_text:
            target_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

    def rx(label: str, stop_labels: List[str]) -> str:
        stop = "|".join([re.escape(s) for s in stop_labels])
        pattern = rf"{re.escape(label)}\s+(.*?)(?=\s+(?:{stop}))"
        m = re.search(pattern, target_text, flags=re.DOTALL)
        if m:
            captured = clean_text(m.group(1))
            if len(captured) < 200:
                return captured
        
        m2 = re.search(rf"{re.escape(label)}\s+([^\n]+(?:\n[^\n]+)?)", target_text)
        if m2:
            result = clean_text(m2.group(1))
            for stop_word in stop_labels:
                if stop_word in result:
                    result = result.split(stop_word)[0].strip()
            return result
            
        return ""

    master = {
        "design_number": rx("Design Number", [
            "Design Concept", "Description", "Category", "BOM Number", "Tech Pack"
        ]),
        "description": rx("Description", [
            "Category", "BOM Number", "Design BOM", "Design Type", "Tech Pack"
        ]),
        "bom_number": rx("BOM Number", [
            "Sub-", "SubCategory", "Sub-Category", "Design BOM", "Tech Pack BOM", 
            "Category", "Legacy", "Status"
        ]),
        "legacy_style_numbers": rx("Legacy Style Numbers", [
            "Carryover", "Hang/Fold", "Season Planning", "Brand/Division", 
            "Booking", "Good/Better", "Supplier", "Hard Tag", "RFID"
        ]),
        "hang_fold_instructions": rx("Hang/Fold Instructions", [
            "Booking Track", "Season Planning", "Brand/Division", "Department", 
            "Collection", "BOM Comments", "Revision", "Good/Better"
        ]),
    }
    
    # Design Number ì¶”ê°€ ê²€ìƒ‰
    if not master.get("design_number"):
        tech_pack_pattern = r'Tech Pack[^\n]*?(D\d{5,6})'
        match = re.search(tech_pack_pattern, first_page_text or target_text, re.IGNORECASE)
        if match:
            master["design_number"] = match.group(1)
        else:
            first_part = (first_page_text or target_text)[:500]
            pattern = r'\b(D\d{5,6})\b'
            match = re.search(pattern, first_part)
            if match:
                master["design_number"] = match.group(1)
            else:
                match = re.search(pattern, target_text)
                if match:
                    master["design_number"] = match.group(1)

    # í›„ì²˜ë¦¬
    for k in list(master.keys()):
        value = master[k]
        
        if k == "bom_number" and value:
            match = re.search(r'(\d{8,})', value)
            if match:
                master[k] = match.group(1)
            else:
                master[k] = ""
        
        elif k == "legacy_style_numbers" and value:
            invalid_keywords = [
                "Material", "Supplier", "Approved", "Status", "Allocate",
                "Booking", "Track", "Good", "Better", "Best", "Priority",
                "Carryover", "Season", "Brand", "Division", "INTERNATIONAL",
                "HOLDINGS", "LTD", "CORP", "Master", "Primary", "RD", "Comment"
            ]
            
            if any(keyword in value for keyword in invalid_keywords):
                master[k] = ""
            else:
                match = re.search(r'(\d{6,})', value)
                if match:
                    master[k] = match.group(1)
                else:
                    master[k] = ""
        
        elif k == "hang_fold_instructions" and value:
            if "Tops-" in value or "Tops -" in value:
                match = re.search(r'Tops-?\s*(\w+)', value, re.IGNORECASE)
                if match:
                    second_word = match.group(1)
                    if second_word in ["Hang", "Fold", "Flat", "Roll"]:
                        master[k] = f"Tops- {second_word}"
                    else:
                        master[k] = "Tops-"
                else:
                    master[k] = "Tops-"
            
            invalid_keywords = [
                "Brand/Division", "BOM Comments", "Department", "Collection",
                "Season Planning", "Revision", "Modified", "Booking Track"
            ]
            
            if any(word in value for word in invalid_keywords):
                master[k] = ""
            elif len(value) > 50:
                master[k] = ""

    return master


def extract_color_headers_from_bom_colormatrix(pdf_path: str) -> List[str]:
    """
    BOMColorMatrix ì„¹ì…˜ì—ì„œ ì»¬ëŸ¬ í—¤ë”ë¥¼ ì¶”ì¶œ.
    
    ë‘ ê°€ì§€ ë°©ì‹ìœ¼ë¡œ ì‹œë„:
    1. í…ìŠ¤íŠ¸ ê¸°ë°˜: "|" êµ¬ë¶„ìžê°€ ìžˆëŠ” ê²½ìš°
    2. í…Œì´ë¸” ê¸°ë°˜: Components íŽ˜ì´ì§€ì˜ í…Œì´ë¸”ì—ì„œ CC Name / BOM CC Number ì»¬ëŸ¼ íŒŒì‹±
    """
    headers: List[str] = []
    skip_keywords = [
        "CC Name", "Component", "Type", "Status", "Created", "Modified",
        "BOM CC Number", "Product Sustainability", "HERALD", "Concept", "Adopted",
        "BOMColorMatrix", "Displaying",
    ]

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            if "BOMColorMatrix" not in text and "CC Name" not in text:
                continue

            # â”€â”€ ë°©ë²• 1: í…ìŠ¤íŠ¸ì— "|" êµ¬ë¶„ìžê°€ ìžˆìœ¼ë©´ í…ìŠ¤íŠ¸ íŒŒì‹± â”€â”€
            if "|" in text and "CC Name" in text:
                lines = text.split('\n')
                in_color_section = False
                for line in lines:
                    line_stripped = line.strip()
                    if "BOMColorMatrix" in line_stripped or "CC Name" in line_stripped:
                        in_color_section = True
                        continue
                    if in_color_section and ("Documents" in line_stripped or "Measurement" in line_stripped or
                                             "POM Name" in line_stripped or "Displaying" in line_stripped):
                        break
                    if in_color_section and line_stripped and "|" in line_stripped:
                        parts = [p.strip() for p in line_stripped.split("|")]
                        if len(parts) >= 3:
                            cc_name = clean_text_keep_newlines(parts[0])
                            bom_cc = clean_text(parts[2])
                            potential = format_color_header_text(f"{cc_name}\n{bom_cc}" if bom_cc else cc_name)
                            if potential and len(potential) > 5 and not any(kw in potential for kw in skip_keywords):
                                headers.append(potential)
                if headers:
                    return headers

            # â”€â”€ ë°©ë²• 2: í…Œì´ë¸” ê¸°ë°˜ íŒŒì‹± â”€â”€
            tables = page.extract_tables() or []
            for table in tables:
                if not table or len(table) < 3:
                    continue

                # CC Name / BOM CC Numberê°€ ìžˆëŠ” í–‰ ì°¾ê¸° (ìƒìœ„ 3í–‰ ë‚´)
                cc_name_idx = None
                bom_cc_idx = None
                header_row_idx = None

                for ri in range(min(3, len(table))):
                    row = table[ri]
                    if not row:
                        continue
                    row_text = " ".join([str(c) for c in row if c])
                    if "CC Name" in row_text:
                        for ci, ch in enumerate(row):
                            ch_text = clean_text(ch)
                            if "CC Name" in ch_text:
                                cc_name_idx = ci
                            elif "BOM CC Number" in ch_text:
                                bom_cc_idx = ci
                        header_row_idx = ri
                        break

                if cc_name_idx is None or header_row_idx is None:
                    continue

                # í—¤ë” ë‹¤ìŒ í–‰ë¶€í„° ë°ì´í„° ì¶”ì¶œ
                for row in table[header_row_idx + 1:]:
                    if not row or cc_name_idx >= len(row):
                        continue
                    raw_cc_name = clean_text(row[cc_name_idx])
                    if not raw_cc_name:
                        continue
                    bom_cc = clean_text(row[bom_cc_idx]) if (bom_cc_idx is not None and bom_cc_idx < len(row)) else ""

                    # pdfplumber ì…€ ë³‘í•© ì•„í‹°íŒ©íŠ¸ ë³´ì •:
                    # "MA STONES THROW" â†’ "A STONES THROW" (ì•ž ì…€ì˜ ë§ˆì§€ë§‰ ê¸€ìžê°€ ë¶™ëŠ” ê²½ìš°)
                    # "1/8/2026, 6:21 AMA STONES THROW" íŒ¨í„´ ì²˜ë¦¬
                    cc_name_clean = raw_cc_name
                    am_match = re.search(r'(?:AM|PM)([A-Z])', cc_name_clean)
                    if am_match:
                        cc_name_clean = cc_name_clean[am_match.start() + 2:]
                    # ??/?? ??? ??
                    date_prefix = re.match(r'^[\d/,:\s]+(?:AM|PM)\s*', cc_name_clean)
                    if date_prefix:
                        cc_name_clean = cc_name_clean[date_prefix.end():]
                    # OCR artifacts on first color name: "MA STONES..." / "AA STONES..."
                    cc_name_clean = re.sub(r'^\s*MA\s+(?=[A-Z])', 'A ', cc_name_clean)
                    cc_name_clean = re.sub(r'^\s*A{2,}\s+', 'A ', cc_name_clean)
                    cc_name_clean = cc_name_clean.strip()

                    potential = format_color_header_text(f"{cc_name_clean}\n{bom_cc}" if bom_cc else cc_name_clean)
                    if potential and len(potential) > 3 and not any(kw in potential for kw in skip_keywords):
                        headers.append(potential)

            if headers:
                return headers

    return headers


def extract_bom_rows_from_pdf(pdf_path: str) -> Tuple[List[BomRow], List[str]]:
    """
    Extract BOM Details table rows from PDF using pdfplumber.extract_tables().
    
    ê°€ë¡œ ë¶„í•  ì²˜ë¦¬:
    - full-header tableì—ì„œ ê° í–‰ì˜ raw index â†’ BomRow index ë§¤í•‘ ê¸°ë¡
    - continuation tableì—ì„œ ë™ì¼í•œ ë§¤í•‘ìœ¼ë¡œ ì»¬ëŸ¬ ë°ì´í„° ì •í™•ížˆ í• ë‹¹
    - ì»¬ëŸ¬ ìˆ˜ì— ë”°ë¼ 1~NíŽ˜ì´ì§€ì˜ continuationì„ ëª¨ë‘ ì²˜ë¦¬
    """
    rows: List[BomRow] = []
    color_headers_order: List[str] = []
    matrix_headers: List[str] = extract_color_headers_from_bom_colormatrix(pdf_path)

    # â”€â”€ í—¬í¼ í•¨ìˆ˜ë“¤ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _is_excluded_color_column_header(raw_header: str) -> bool:
        hn = normalize_header(raw_header)
        if hn in {"commonqty", "commoncolor"}:
            return True
        # "Only for Product Colors" 컬럼 자체는 컬러가 아님 (CC Number 없이 단독인 경우)
        if hn == "onlyforproductcolors":
            return True
        return False

    def _map_value_to_matrix_header(value: str) -> str:
        v = clean_text(value)
        if not v or not matrix_headers:
            return ""
        base = re.sub(r"\s+\d{2,4}$", "", v).strip()
        base_norm = normalize_header(base) if base else ""
        if base_norm:
            for h in matrix_headers:
                if base_norm and base_norm in normalize_header(h):
                    return h
        lv = v.lower()
        if "tango" in lv:
            for h in matrix_headers:
                hn = normalize_header(h)
                if ("seasalt" in hn and "blue" in hn) or ("seasaltwblue" in hn):
                    return h
        return ""

    def _extract_cc_number(text: str) -> str:
        m = re.search(r"\b(\d{9,})\b", clean_text(text))
        return m.group(1) if m else ""

    def _map_header_to_matrix_header(header_txt: str) -> str:
        if not header_txt or not matrix_headers:
            return ""
        cc = _extract_cc_number(header_txt)
        if cc:
            for h in matrix_headers:
                if _extract_cc_number(h) == cc:
                    return h
        hn = normalize_header(header_txt)
        if hn:
            for h in matrix_headers:
                mh = normalize_header(h)
                if hn in mh or mh in hn:
                    return h
        return ""

    def _sanitize_color_header(header_txt: str, allow_loose: bool = False) -> str:
        h = format_color_header_text(header_txt)
        if not h:
            return ""
        hc = clean_text(h)
        if re.fullmatch(r"\d{6,}", hc):
            return ""
        hn = normalize_header(h)
        if hn in {
            "product", "materialname", "supplierarticlenumber", "usage", "qualitydetails",
            "supplierallocate", "supplier", "comment", "image", "primaryrd",
            "commonsize", "commonqty", "gaugeends", "stitch", "onlyforproductcolors",
            "commoncolor",
        }:
            return ""
        if allow_loose:
            return h
        if matrix_headers:
            mapped = _map_header_to_matrix_header(h)
            return mapped or ""
        return h if _extract_cc_number(h) else ""

    def _resolve_graphic_header(header_txt: str, value_txt: str, color_pos: int) -> str:
        """
        Graphic should reuse existing color columns.
        Never create a new color header here.
        """
        raw_h = _sanitize_color_header(header_txt, allow_loose=True)
        if raw_h and raw_h in color_headers_order:
            return raw_h

        mapped_h = _map_header_to_matrix_header(raw_h) if raw_h else ""
        if mapped_h and mapped_h in color_headers_order:
            return mapped_h

        mapped_v = _map_value_to_matrix_header(value_txt) if value_txt else ""
        if mapped_v and mapped_v in color_headers_order:
            return mapped_v

        if 0 <= color_pos < len(color_headers_order):
            return color_headers_order[color_pos]

        return ""

    def _find_graphic_color_image(prod: str, material: str, *header_variants) -> Optional[bytes]:
        """
        Try multiple header text variants then CC-number fallback
        to find a graphic color image from graphic_color_images dict.
        """
        for htxt in header_variants:
            if htxt:
                b = graphic_color_images.get((prod, material, htxt))
                if b:
                    return b
        # CC number fallback: match by same CC number in any stored key
        for htxt in header_variants:
            if not htxt:
                continue
            cc_m = re.search(r'\b(\d{9,})\b', htxt)
            if cc_m:
                cc_num = cc_m.group(1)
                for (p, m, h), img_bytes in graphic_color_images.items():
                    if p == prod and m == material and cc_num in h:
                        return img_bytes
        return None

    def _is_full_header(header_norm: List[str]) -> bool:
        """í…Œì´ë¸”ì´ Product, Material Name ë“± ì „ì²´ í—¤ë”ë¥¼ ê°€ì§€ê³  ìžˆëŠ”ì§€"""
        need = {"product", "materialname", "supplierarticlenumber", "usage", "qualitydetails"}
        supplier_ok = ("supplierallocate" in header_norm) or ("supplier" in header_norm)
        return all(n in header_norm for n in need) and supplier_ok

    def _is_color_continuation_table(header: List[str], header_norm: List[str],
                                      tbl: List[List[Optional[str]]]) -> bool:
        """
        ê°€ë¡œ ë¶„í• ëœ ì»¬ëŸ¬ ì „ìš© continuation í…Œì´ë¸”ì¸ì§€ íŒë³„.
        
        ì¡°ê±´:
        1. ì´ì „ì— ì²˜ë¦¬ëœ ë¸”ë¡(current_block_rows)ì´ ì¡´ìž¬í•´ì•¼ í•¨
        2. í…Œì´ë¸”ì— 2í–‰ ì´ìƒ ìžˆì–´ì•¼ í•¨ (í—¤ë” + ìµœì†Œ 1 ë°ì´í„°)
        3. Product/MaterialName ê°™ì€ ê¸°ë³¸ ì»¬ëŸ¼ì´ ì—†ì–´ì•¼ í•¨
        4. í—¤ë”ì— 9ìžë¦¬ ì´ìƒ ìˆ«ìž(CC Number)ê°€ í¬í•¨ë˜ì–´ì•¼ í•¨
        5. ë°ì´í„° í–‰ ìˆ˜ê°€ ì›ë³¸ í…Œì´ë¸”ì˜ data row ìˆ˜ì™€ ìœ ì‚¬í•´ì•¼ í•¨
        """
        if not current_block_rows:
            return False
        if not tbl or len(tbl) < 2:
            return False
        if "product" in header_norm or "materialname" in header_norm:
            return False

        meaningful_headers = [h for h in header if clean_text_keep_newlines(h)]
        if not meaningful_headers:
            return False
        if len(meaningful_headers) == 1 and normalize_header(meaningful_headers[0]) == "comment":
            return False

        # CC Number (9ìžë¦¬+) ì¡´ìž¬ í™•ì¸
        if not any(re.search(r"\b\d{9,}\b", h) for h in meaningful_headers):
            return False

        # â˜… í•µì‹¬ ìˆ˜ì •: ì›ë³¸ í…Œì´ë¸”ì˜ ì „ì²´ raw í–‰ ìˆ˜ì™€ ë¹„êµ
        data_rows = len(tbl) - 1
        if last_full_table_raw_data_count > 0:
            return data_rows <= last_full_table_raw_data_count + 2
        else:
            return data_rows >= 1

    def _looks_like_noise_color_value(v: str) -> bool:
        vv = clean_text(v)
        if not vv:
            return True
        if len(vv) > 60:
            return True
        bad_keywords = [
            "displaying", "units:", "grade", "pom", "measurement", "tol fraction",
            "grading on this critical", "from center back", "high point shoulder",
        ]
        lvv = vv.lower()
        return any(bk in lvv for bk in bad_keywords)

    def _is_footer_or_noise_row(cells: List[str]) -> bool:
        joined = " ".join([clean_text(x) for x in cells]).lower()
        if not joined.strip():
            return True
        noise = [
            "displaying", "results", "page ", "centric", "production(",
            "units:", "grade rule", "measurement chart",
        ]
        return any(n in joined for n in noise)

    def _detect_section_from_row(cells: List[str]) -> Optional[str]:
        if cells:
            sec = section_from_cell_text(cells[0])
            if sec:
                return sec
        return None

    def _apply_continuation_colors(tbl: List[List[Optional[str]]],
                                    header: List[str],
                                    header_norm: List[str]) -> bool:
        """
        continuation í…Œì´ë¸”ì˜ ì»¬ëŸ¬ ë°ì´í„°ë¥¼ current_block_rowsì— ì •í™•ížˆ ë§¤í•‘.
        row_to_bomrow_mapì„ ì‚¬ìš©í•˜ì—¬ ì„¹ì…˜ í—¤ë”/ë¹ˆ í–‰ì„ ê±´ë„ˆë›°ê³ 
        ì˜¬ë°”ë¥¸ BomRowì— ì»¬ëŸ¬ ê°’ì„ í• ë‹¹í•¨.
        """
        comment_idx = header_norm.index("comment") if "comment" in header_norm else None
        color_col_indices: List[int] = []
        for ci, hn in enumerate(header_norm):
            if comment_idx is not None and ci == comment_idx:
                continue
            if clean_text_keep_newlines(header[ci]) and not _is_excluded_color_column_header(header[ci]):
                color_col_indices.append(ci)

        if not color_col_indices:
            return False

        appended_any = False

        for data_i in range(1, len(tbl)):
            raw_idx = data_i - 1  # 0-based data row index

            # â˜… row_to_bomrow_mapìœ¼ë¡œ ì •í™•í•œ BomRow ì¸ë±ìŠ¤ ì°¾ê¸°
            target_i = row_to_bomrow_map.get(raw_idx)
            if target_i is None:
                # ë§¤í•‘ì— ì—†ìœ¼ë©´ â†’ ì„¹ì…˜ í—¤ë”/ë¹ˆ í–‰/footer â†’ skip
                continue
            if target_i >= len(current_block_rows):
                continue

            r = tbl[data_i]
            if not r:
                continue

            for ci in color_col_indices:
                if ci >= len(r):
                    continue
                v = clean_text(r[ci])
                if not v or _looks_like_noise_color_value(v):
                    continue
                is_graphic = ((current_block_rows[target_i].category or "").lower() == "graphic")
                raw_header_txt = format_color_header_text(header[ci] if ci < len(header) else "")
                if is_graphic:
                    header_txt = _resolve_graphic_header(raw_header_txt, v, ci)
                else:
                    header_txt = _sanitize_color_header(raw_header_txt)
                if not header_txt:
                    continue
                if (not is_graphic) and header_txt not in color_headers_order:
                    color_headers_order.append(header_txt)
                current_block_rows[target_i].colors[header_txt] = v
                appended_any = True

                # Graphic: find color image with fallback matching
                brow = current_block_rows[target_i]
                if (brow.category or "").lower() == "graphic":
                    b = _find_graphic_color_image(brow.product, brow.material_name, header_txt, raw_header_txt)
                    if b:
                        brow.color_images[header_txt] = b
        return appended_any

    def _append_color_values_from_text_continuation(page_text: str) -> bool:
        """í…Œì´ë¸” ì¶”ì¶œ ì‹¤íŒ¨ ì‹œ í…ìŠ¤íŠ¸ ê¸°ë°˜ìœ¼ë¡œ ì»¬ëŸ¬ continuation ë³´ê°•"""
        if not current_block_rows:
            return False
        if not page_text:
            return False

        lines = (page_text or "").split("\n")
        start_idx = None
        for i, line in enumerate(lines):
            if "Comment" in line and " - " in line:
                start_idx = i
                break
            if re.search(r"\b\d{9,}\b", line) and " - " in line and not any(
                kw in line for kw in ["Product", "Material", "Supplier", "Quality", "Centric", "Production"]
            ):
                start_idx = i
                break
        if start_idx is None:
            return False

        header_line = clean_text_keep_newlines(lines[start_idx])
        chunk_re = re.compile(r"([A-Z0-9][A-Za-z0-9\s\-/]*?\b\d{9,}\b)")
        header_chunks = [_sanitize_color_header(c) for c in chunk_re.findall(header_line)]
        header_chunks = [h for h in header_chunks if h]
        if header_chunks:
            for h in header_chunks:
                if h not in color_headers_order:
                    color_headers_order.append(h)

        data_lines: List[str] = []
        for line in lines[start_idx + 1:]:
            l = clean_text(line)
            if not l:
                continue
            if l.lower().startswith("displaying") or "measurement" in l.lower():
                break
            data_lines.append(l)

        if not data_lines:
            return False

        token_re = re.compile(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3}\s+\d{2,4})\b")

        appended_any = False
        target_i = 0
        for dl in data_lines:
            if target_i >= len(current_block_rows):
                break
            tokens = token_re.findall(dl)
            if not tokens:
                continue
            if header_chunks:
                for col_i, t in enumerate(tokens[: len(header_chunks)]):
                    tv = clean_text(t)
                    if not tv or _looks_like_noise_color_value(tv):
                        continue
                    hk = header_chunks[col_i]
                    if current_block_rows[target_i].colors.get(hk) != tv:
                        current_block_rows[target_i].colors[hk] = tv
                        appended_any = True
            else:
                for t in tokens:
                    tv = clean_text(t)
                    if not tv or _looks_like_noise_color_value(tv):
                        continue
                    header_key = _map_value_to_matrix_header(tv)
                    if header_key and header_key not in color_headers_order:
                        color_headers_order.append(header_key)
                    if header_key and current_block_rows[target_i].colors.get(header_key) != tv:
                        current_block_rows[target_i].colors[header_key] = tv
                        appended_any = True
            target_i += 1

        return appended_any

    # â”€â”€ ì´ë¯¸ì§€ ë§µ ì‚¬ì „ ì¶”ì¶œ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    image_map = extract_bom_image_map_from_pdf(pdf_path)
    graphic_color_images = extract_graphic_color_cell_images_from_pdf(pdf_path)

    # â”€â”€ ë©”ì¸ íŒŒì‹± ë£¨í”„ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    last_valid_header_info = None
    current_block_rows: List[BomRow] = []
    row_to_bomrow_map: Dict[int, int] = {}       # â˜… raw_data_idx â†’ BomRow ì¸ë±ìŠ¤
    last_full_table_raw_data_count: int = 0       # â˜… ì›ë³¸ í…Œì´ë¸”ì˜ ì „ì²´ data í–‰ ìˆ˜ (header ì œì™¸)
    rows_per_page = {}

    with pdfplumber.open(pdf_path) as pdf:
        current_section: str = ""

        for page_num, page in enumerate(pdf.pages, 1):
            page_row_count = 0
            table_objs = page.find_tables() or []
            appended_continuation_this_page = False

            for tbl_idx, tbl_obj in enumerate(table_objs):
                tbl = tbl_obj.extract() or []
                if not tbl or len(tbl) < 1 or not tbl[0]:
                    continue

                header = [clean_text_keep_newlines(c) for c in tbl[0]]
                header_norm = [normalize_header(c) for c in header]

                # â”€â”€â”€ 1) ê°€ë¡œ ë¶„í•  ì»¬ëŸ¬ continuation í…Œì´ë¸” â”€â”€â”€
                if _is_color_continuation_table(header, header_norm, tbl):
                    if _apply_continuation_colors(tbl, header, header_norm):
                        appended_continuation_this_page = True

                    # â˜… continuation í…Œì´ë¸”ì—ì„œë„ Graphic ì´ë¯¸ì§€ ì¶”ì¶œ
                    if current_block_rows and row_to_bomrow_map:
                        cont_imgs = extract_continuation_graphic_images(
                            page, tbl_obj, row_to_bomrow_map,
                            current_block_rows, header, header_norm,
                        )
                        for (prod, mat, htxt), png_bytes in cont_imgs.items():
                            for brow in current_block_rows:
                                if brow.product == prod and brow.material_name == mat:
                                    # Map raw continuation header to color_headers_order entry
                                    resolved = _map_header_to_matrix_header(htxt) if htxt else ""
                                    final_key = resolved if (resolved and resolved in color_headers_order) else htxt
                                    if final_key and final_key not in brow.color_images:
                                        brow.color_images[final_key] = png_bytes
                    continue

                # â”€â”€â”€ 2) Full-header í…Œì´ë¸” ì²˜ë¦¬ â”€â”€â”€
                has_valid_header = _is_full_header(header_norm)
                
                if has_valid_header:
                    idx_product = header_norm.index("product")
                    idx_material = header_norm.index("materialname")
                    idx_supp_art = header_norm.index("supplierarticlenumber")
                    idx_usage = header_norm.index("usage")
                    idx_quality = header_norm.index("qualitydetails")
                    idx_supplier = header_norm.index("supplierallocate") if "supplierallocate" in header_norm else header_norm.index("supplier")

                    color_start = None
                    # "Only for Product Colors" 컬럼 찾기 - 정확히 일치 또는 부분 포함
                    ofpc_idx = None
                    for ci_h, hn_h in enumerate(header_norm):
                        if hn_h == "onlyforproductcolors":
                            ofpc_idx = ci_h
                            break
                        # pdfplumber가 셀을 병합한 경우: "onlyforproductcolors" + 첫 번째 컬러가 합쳐짐
                        if "onlyforproductcolors" in hn_h and hn_h != "onlyforproductcolors":
                            ofpc_idx = ci_h
                            break

                    if ofpc_idx is not None:
                        ofpc_raw = clean_text_keep_newlines(header[ofpc_idx]) if ofpc_idx < len(header) else ""
                        # 병합 여부 체크: 같은 셀에 CC Number(9자리+)가 포함되면 병합된 것
                        has_cc_in_ofpc = bool(re.search(r'\b\d{9,}\b', ofpc_raw))
                        if has_cc_in_ofpc:
                            # 병합됨 → 이 셀 자체가 첫 번째 컬러 컬럼
                            # 헤더 텍스트에서 "Only for Product Colors" 부분 제거
                            color_start = ofpc_idx
                            # header 텍스트 보정: "Only for Product Colors" 제거
                            cleaned = re.sub(
                                r'(?i)only\s+for\s+product\s+colors?\s*[\n\r]*',
                                '', ofpc_raw
                            ).strip()
                            if cleaned:
                                header[ofpc_idx] = cleaned
                        else:
                            # 별도 컬럼 → 다음부터 컬러
                            color_start = ofpc_idx + 1
                    else:
                        color_start = idx_supplier + 1
                    
                    # ★ 추가 안전장치: color_start 위치의 헤더가 "Only for Product Colors"이면 건너뛰기
                    while color_start < len(header):
                        cs_norm = normalize_header(header[color_start]) if color_start < len(header) else ""
                        if cs_norm == "onlyforproductcolors":
                            color_start += 1
                        else:
                            break

                    color_end = len(header)
                    if "comment" in header_norm:
                        color_end = header_norm.index("comment")
                    
                    last_valid_header_info = {
                        'idx_product': idx_product,
                        'idx_material': idx_material,
                        'idx_supp_art': idx_supp_art,
                        'idx_usage': idx_usage,
                        'idx_quality': idx_quality,
                        'idx_supplier': idx_supplier,
                        'color_start': color_start,
                        'color_end': color_end,
                        'num_cols': len(header)
                    }
                    
                    data_start = 1

                elif last_valid_header_info is not None:
                    first_row = tbl[0] if tbl else []
                    first_cell = clean_text(first_row[0] if first_row else "")
                    
                    if re.fullmatch(r"\d{5,}", first_cell):
                        idx_product = last_valid_header_info['idx_product']
                        idx_material = last_valid_header_info['idx_material']
                        idx_supp_art = last_valid_header_info['idx_supp_art']
                        idx_usage = last_valid_header_info['idx_usage']
                        idx_quality = last_valid_header_info['idx_quality']
                        idx_supplier = last_valid_header_info['idx_supplier']
                        color_start = last_valid_header_info['color_start']
                        color_end = last_valid_header_info['color_end']
                        
                        data_start = 0
                    else:
                        continue
                else:
                    continue

                # â”€â”€â”€ 3) ë°ì´í„° í–‰ íŒŒì‹± + row mapping êµ¬ì¶• â”€â”€â”€
                block_rows: List[BomRow] = []
                new_row_mapping: Dict[int, int] = {}  # raw_idx â†’ BomRow index

                for r_idx in range(data_start, len(tbl)):
                    raw_idx = r_idx - data_start  # 0-based data row index
                    r = tbl[r_idx]

                    if not r or len(r) <= max(idx_supplier, idx_quality, idx_material):
                        continue

                    prod = clean_text(r[idx_product] if idx_product < len(r) else "")
                    material = clean_text(r[idx_material] if idx_material < len(r) else "")
                    row_cells_text = [prod, material] + [clean_text(x) for x in r]

                    # footer/noise â†’ skip (ë§¤í•‘ ì•ˆ í•¨)
                    if _is_footer_or_noise_row(row_cells_text):
                        continue

                    # ì„¹ì…˜ í—¤ë” â†’ section ì—…ë°ì´íŠ¸, skip (ë§¤í•‘ ì•ˆ í•¨)
                    sec = _detect_section_from_row(row_cells_text)
                    if sec:
                        current_section = sec
                        continue

                    # Product ìœ íš¨ì„± ê²€ì‚¬
                    if not re.fullmatch(r"\d{5,}", prod):
                        if current_section.lower() == "graphic":
                            prod = prod or "GRAPHIC"
                        else:
                            continue

                    if not (prod or material):
                        continue

                    supp_art = clean_text(r[idx_supp_art] if idx_supp_art < len(r) else "")
                    usage = clean_text(r[idx_usage] if idx_usage < len(r) else "")
                    quality = clean_text(r[idx_quality] if idx_quality < len(r) else "")
                    supplier = clean_text(r[idx_supplier] if idx_supplier < len(r) else "")

                    # ì´ í…Œì´ë¸”ì— í¬í•¨ëœ ì»¬ëŸ¬ ì¶”ì¶œ
                    colors: Dict[str, str] = {}
                    actual_color_end = min(color_end, len(r))
                    color_col_position = 0  # color_start부터의 위치 (matrix_headers 매칭용)
                    for ci in range(color_start, actual_color_end):
                        v = clean_text(r[ci])
                        if ci < len(header) and _is_excluded_color_column_header(header[ci]):
                            color_col_position += 1
                            continue

                        is_graphic_row = ((current_section or "").lower() == "graphic")
                        raw_header_txt = format_color_header_text(header[ci] if ci < len(header) else "")
                        if is_graphic_row:
                            header_txt = _resolve_graphic_header(raw_header_txt, v, color_col_position)
                        else:
                            header_txt = _sanitize_color_header(raw_header_txt)

                        # ? ??? ???? ?: matrix_headers?? ?? ?? ?? ? ?? ??
                        if not header_txt and matrix_headers and (not is_graphic_row):
                            # ?? 1: ??? ??? matrix header ??
                            if v:
                                matched = _map_value_to_matrix_header(v)
                                if matched:
                                    header_txt = matched
                            # ?? 2: ?? ?? ?? (color_start?? ????)
                            if not header_txt and color_col_position < len(matrix_headers):
                                header_txt = matrix_headers[color_col_position]

                        color_col_position += 1

                        if is_graphic_row:
                            header_txt = _resolve_graphic_header(header_txt, v, color_col_position - 1)
                        else:
                            header_txt = _sanitize_color_header(header_txt)
                        if not header_txt:
                            continue
                        if not v:
                            continue
                        if (not is_graphic_row) and header_txt not in color_headers_order:
                            color_headers_order.append(header_txt)
                        colors[header_txt] = v

                    color_images: Dict[str, bytes] = {}
                    if (current_section or "").lower() == "graphic":
                        for htxt in list(colors.keys()):
                            b = _find_graphic_color_image(prod, material, htxt, raw_header_txt)
                            if b:
                                color_images[htxt] = b
                    bomrow_idx = len(block_rows)
                    new_row_mapping[raw_idx] = bomrow_idx  # â˜… ë§¤í•‘ ê¸°ë¡

                    block_rows.append(
                        BomRow(
                            category=current_section,
                            product=prod,
                            material_name=material,
                            supplier_article_number=supp_art,
                            usage=usage,
                            quality_details=quality,
                            supplier=supplier,
                            colors=colors,
                            image_png=image_map.get((current_section, prod, material)),
                            color_images=color_images,
                        )
                    )
                    page_row_count += 1

                if block_rows:
                    rows.extend(block_rows)
                    current_block_rows = block_rows
                    row_to_bomrow_map = new_row_mapping
                    last_full_table_raw_data_count = len(tbl) - data_start  # â˜… ì „ì²´ data í–‰ ìˆ˜ ê¸°ë¡

            # í…Œì´ë¸”ë¡œ ëª» ìž¡ì€ continuation â†’ í…ìŠ¤íŠ¸ fallback
            if current_block_rows and not appended_continuation_this_page:
                try:
                    page_text = page.extract_text() or ""
                    if _append_color_values_from_text_continuation(page_text):
                        appended_continuation_this_page = True
                except Exception:
                    pass

            rows_per_page[page_num] = page_row_count

    if not color_headers_order and matrix_headers:
        color_headers_order = matrix_headers.copy()

    # â˜… ìž˜ë¦° í—¤ë” ë³´ì •: BOMColorMatrixì—ì„œ ê°€ì ¸ì˜¨ ì „ì²´ í—¤ë”ì™€ ë§¤ì¹­
    if matrix_headers and color_headers_order:
        _fix_truncated_headers(color_headers_order, matrix_headers, rows)

    return rows, color_headers_order


def _fix_truncated_headers(color_headers_order: List[str],
                           matrix_headers: List[str],
                           rows: List[BomRow]) -> None:
    """
    pdfplumberê°€ ì¢ì€ ì»¬ëŸ¼ì—ì„œ ìž˜ë¼ë‚¸ í—¤ë”ë¥¼
    BOMColorMatrixì˜ ì „ì²´ í—¤ë”ë¡œ êµì²´.
    
    ë§¤ì¹­ ê¸°ì¤€: ê°™ì€ CC Number(9ìžë¦¬+) í¬í•¨ ì—¬ë¶€
    ë‹¨, BOM Detail í—¤ë”ê°€ ì´ë¯¸ ì™„ì „í•œ CC Numberë¥¼ ê°€ì§€ê³  ìžˆìœ¼ë©´ êµì²´í•˜ì§€ ì•ŠìŒ
    (BOMColorMatrix í…Œì´ë¸” íŒŒì‹± ì‹œ ì…€ ë³‘í•© ì•„í‹°íŒ©íŠ¸ë¡œ CC Nameì´ ì˜¤ì—¼ë  ìˆ˜ ìžˆìœ¼ë¯€ë¡œ)
    """
    def _extract_cc_number(h: str) -> str:
        m = re.search(r'(\d{9,})', h)
        return m.group(1) if m else ""

    def _has_complete_cc_number(h: str) -> bool:
        """í—¤ë”ê°€ ì™„ì „í•œ 9ìžë¦¬+ CC Numberë¥¼ í¬í•¨í•˜ëŠ”ì§€"""
        m = re.search(r'\d{9,}', h)
        if not m:
            return False
        # CC Number ë’¤ì— ìˆ«ìžê°€ ë” ìžˆìœ¼ë©´ ìž˜ë¦° ê²ƒì´ ì•„ë‹˜
        # "000003239937" â†’ ì™„ì „, "0" ë˜ëŠ” "00" â†’ ìž˜ë¦° ê²ƒ
        return len(m.group()) >= 9

    # matrix í—¤ë”ë¥¼ CC Numberë¡œ ì¸ë±ì‹±
    matrix_by_cc: Dict[str, str] = {}
    for mh in matrix_headers:
        cc = _extract_cc_number(mh)
        if cc:
            matrix_by_cc[cc] = mh

    if not matrix_by_cc:
        return

    # êµì²´ í•„ìš”í•œ í—¤ë” ì°¾ê¸° (CC Numberê°€ ìž˜ë¦° ê²ƒë§Œ)
    replacements: Dict[str, str] = {}  # old_header â†’ new_header
    for old_h in color_headers_order:
        if _has_complete_cc_number(old_h):
            # CC Numberê°€ ì´ë¯¸ ì™„ì „í•¨ â†’ êµì²´ ë¶ˆí•„ìš”
            continue
        # CC Number ë¶€ë¶„ ë§¤ì¹­: ìž˜ë¦° ìˆ«ìž(1~8ìžë¦¬)ë¡œ ì‹œìž‘í•˜ëŠ” ì™„ì „í•œ CC Number ì°¾ê¸°
        trailing = re.findall(r'(\d+)\s*$', old_h.replace('\n', ' '))
        if not trailing:
            continue
        partial = trailing[-1]
        for full_cc, full_h in matrix_by_cc.items():
            if full_cc.startswith(partial) and len(partial) < len(full_cc):
                replacements[old_h] = full_h
                break

    if not replacements:
        return

    # color_headers_order êµì²´
    for i, h in enumerate(color_headers_order):
        if h in replacements:
            color_headers_order[i] = replacements[h]

    # ëª¨ë“  BomRowì˜ colors/color_images í‚¤ êµì²´
    for row in rows:
        for old_h, new_h in replacements.items():
            if old_h in row.colors:
                row.colors[new_h] = row.colors.pop(old_h)
            if old_h in row.color_images:
                row.color_images[new_h] = row.color_images.pop(old_h)
