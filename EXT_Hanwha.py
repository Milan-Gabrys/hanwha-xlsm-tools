
# -*- coding: utf-8 -*-
"""
Hanwha XLSM odemknutí (Krok 1b):
- In-memory ZIP (.xlsm)
- Odstraní <sheetProtection .../> v daném sheet XML
- Odstraní <workbookProtection .../> v xl/workbook.xml
- -debug: rozbalí výsledek do temp složky pro kontrolu
"""

import sys
import io
import re
import argparse
import zipfile
from datetime import datetime
from pathlib import Path


def remove_tag_bytes(xml_bytes: bytes, tag_local_name: str) -> bytes:
    """
    Odstraní self-closing element <...tag_local_name.../> na úrovni bytes.
    - Žádné XML parsování: zachová se hlavička, whitespace, CRLF, pořadí atributů.
    """
    pattern = re.compile(rb'<[^>]*' + re.escape(tag_local_name.encode('utf-8')) + rb'[^>]*/>', re.IGNORECASE)
    return pattern.sub(b'', xml_bytes)


def process_xlsm_in_memory(input_xlsm: Path, sheet_xml_rel: str) -> bytes:
    """
    Načte .xlsm do paměti, upraví:
      - sheet XML (odstraní sheetProtection)
      - xl/workbook.xml (odstraní workbookProtection)
    a vytvoří nový ZIP v paměti.
    """
    original_bytes = input_xlsm.read_bytes()
    in_mem_zip = io.BytesIO(original_bytes)
    out_mem_zip = io.BytesIO()

    with zipfile.ZipFile(in_mem_zip, 'r') as zin, zipfile.ZipFile(out_mem_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)

            if info.filename == sheet_xml_rel:
                before, data = len(data), remove_tag_bytes(data, 'sheetProtection')
                after = len(data)
                print(f"[INFO] {sheet_xml_rel}: {before} → {after} bytes (sheetProtection removed: {before != after})")

            elif info.filename == 'xl/workbook.xml':
                before, data = len(data), remove_tag_bytes(data, 'workbookProtection')
                after = len(data)
                print(f"[INFO] xl/workbook.xml: {before} → {after} bytes (workbookProtection removed: {before != after})")

            zout.writestr(info, data)

    return out_mem_zip.getvalue()


def save_modified_xlsm(out_bytes: bytes, base_name: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = Path.cwd() / f"{base_name}_modified_{ts}.xlsm"
    out_path.write_bytes(out_bytes)
    return out_path


def debug_extract_zip(zip_bytes: bytes) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_dir = Path.cwd() / f"_tempHanwha_{ts}"
    temp_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes), 'r') as z:
        z.extractall(temp_dir)
    return temp_dir


def main():
    parser = argparse.ArgumentParser(description="Hanwha XLSM (Step 1b): remove sheet/workbook protection by bytes regex")
    parser.add_argument("input", help="Vstupní .xlsm soubor")
    parser.add_argument("--sheet-xml", default="xl/worksheets/sheet2.xml",
                        help="Relativní cesta k sheet XML (default: xl/worksheets/sheet2.xml)")
    parser.add_argument("-debug", action="store_true", help="Rozbalit upravený XLSM do temp složky pro kontrolu")
    args = parser.parse_args()

    input_xlsm = Path(args.input).resolve()
    base_name = input_xlsm.stem

    try:
        out_bytes = process_xlsm_in_memory(input_xlsm, args.sheet_xml)
        modified_path = save_modified_xlsm(out_bytes, base_name)
        print(f"[OK] Upravený XLSM uložen: {modified_path}")

        if args.debug:
            temp_dir = debug_extract_zip(out_bytes)
            print(f"[DEBUG] Upravený obsah rozbalen do: {temp_dir}")
            print(" → Zkontroluj:", temp_dir / args.sheet_xml)
            print(" → Zkontroluj:", temp_dir / "xl/workbook.xml")

    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
