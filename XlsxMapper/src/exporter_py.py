# -*- coding: utf-8 -*-

import hashlib

from pathlib import Path
from typing import Dict, Any


class PythonScriptExporter:
    """Generates a master Python script to mirror the original XLSX workbook."""

    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        # Stores unique styles: { "hash": "python_code" }
        self.style_registry = {"borders": {}, "fills": {}, "fonts": {}}

    def _generate_style_id(self, style_dict, prefix):
        """Generates a unique ID for a style based on its content."""
        style_str = str(sorted(style_dict.items()))
        style_hash = hashlib.md5(style_str.encode()).hexdigest()[:8]
        return f"{prefix}_{style_hash}".upper()

    def generate_full_workbook(self, workbook_data: Dict[str, Dict[str, Any]]) -> None:
        """
        Generates a master script to rebuild the workbook.
        Expected structure: { "SheetName": {"cells": [...], "dims": {...}, "assets": [...] } }
        """
        sheet_modules = []

        for sheet_name, content in workbook_data.items():
            clean_name = "".join(filter(str.isalnum, sheet_name))
            module_filename = f"sheet_{clean_name}.py"
            sheet_modules.append(clean_name)

            self._write_sheet_module(clean_name, sheet_name, content, module_filename)

        self._write_common_styles()
        self._write_main_reconstructor(sheet_modules)

    def _write_sheet_module(self, clean_name: str, original_name: str, content, filename: str):
        script = [
            "# -*- coding: utf-8 -*-\n",
            "from openpyxl.styles import Alignment, Font, PatternFill, Border, Side",
            "from openpyxl.drawing.image import Image as XLImage",
            "from datetime import datetime as dt, time as tm",
            "from pathlib import Path",
            "from common import *\n",
            f"def rebuild_{clean_name}(wb):",
            f"    ws = wb.create_sheet('{original_name}')"
        ]

        # Apply Dimensions
        dims = content.get("dims", {})
        for col, width in dims.get("cols", {}).items():
            script.append(f"    ws.column_dimensions['{col}'].width = {width}")
        for row, height in dims.get("rows", {}).items():
            script.append(f"    ws.row_dimensions[{row}].height = {height}")

        # Reconstruct Cells (Data & Styles)
        processed_merges = set()
        cells = content.get("cells", [])

        for c in cells:
            # Support both Dict (from JSON) and Dataclass
            data = c if isinstance(c, dict) else c.__dict__
            coord = data['coordinate']

            # Formulas and Values
            if data.get('formula'):
                script.append(f"    ws['{coord}'].value = '{data['formula']}'")
            elif data.get('value') is not None:
                val = data['value']
                formatted_val = f'"""{val}"""' if isinstance(val, str) else val
                script.append(f"    ws['{coord}'].value = {formatted_val}")

                # Alignment & Rotation
                align_data = {}
                if data.get('alignment') and data.get('alignment') != 'left':
                    align_data['horizontal'] = data['alignment']
                if data.get('vertical_align') and data.get('vertical_align') != 'bottom':
                    align_data['vertical'] = data['vertical_align']
                if data.get('text_rotation', 0) != 0:
                    align_data['text_rotation'] = data['text_rotation']

                if align_data:
                    a_id = self._generate_style_id(align_data, "ALIGN")

                    if a_id not in self.style_registry.get("alignments", {}):
                        # Se não existir a chave no dicionário principal, inicialize
                        if "alignments" not in self.style_registry:
                            self.style_registry["alignments"] = {}

                        a_args = [f"{k}={v}" if isinstance(v, int) else f"{k}='{v}'"
                                  for k, v in align_data.items()]
                        self.style_registry["alignments"][a_id] = f"Alignment({', '.join(a_args)})"

                    script.append(f"    ws['{coord}'].alignment = {a_id}")

                # Borders
                if data.get('borders'):
                    b_id = self._generate_style_id(data['borders'], "BORDER")
                    b_args = [f"{s}=Side(border_style='{d['style']}', color='{d['color']}')" for s,
                              d in data['borders'].items()]
                    self.style_registry["borders"][b_id] = f"Border({', '.join(b_args)})"
                    script.append(f"    ws['{coord}'].border = {b_id}")

                # Fill
                if data.get('fill_color'):
                    f_id = f"FILL_{data['fill_color']}"
                    self.style_registry["fills"][f_id] = f"PatternFill(start_color='{data['fill_color']}', fill_type='solid')"
                    script.append(f"    ws['{coord}'].fill = {f_id}")

                # Font
                if data.get('font_bold'):
                    font_id = "FONT_BOLD"
                    self.style_registry["fonts"][font_id] = "Font(bold=True)"
                    script.append(f"    ws['{coord}'].font = {font_id}")

                font_data = {}
                if data.get('font_bold'):
                    font_data['bold'] = True
                if data.get('font_italic'):
                    font_data['italic'] = True
                if data.get('font_color'):
                    font_data['color'] = data['font_color']
                if data.get('font_size'):
                    font_data['size'] = data['font_size']

                if font_data:
                    f_id = self._generate_style_id(font_data, "FONT")
                    if f_id not in self.style_registry["fonts"]:
                        f_args = [f"{k}={v}" if isinstance(v, (int, bool)) else f"{k}='{v}'"
                                  for k, v in font_data.items()]
                        self.style_registry["fonts"][f_id] = f"Font({', '.join(f_args)})"
                    script.append(f"    ws['{coord}'].font = {f_id}")

            # Merge Logic
            m_range = data.get('merge_range')
            if data.get('is_merged') and m_range not in processed_merges:
                script.append(f"    ws.merge_cells('{m_range}')")
                processed_merges.add(m_range)

        # Reconstruct Assets (Images)
        assets = content.get("assets", [])
        if assets:
            script.append(f"    # Assets for {original_name}")
            for img in assets:
                script.append("    try:")
                script.append(f"        img_path = Path(__file__).parent / 'images' / '{img['filename']}'")
                script.append("        img_obj = XLImage(img_path)")
                script.append(f"        img_obj.width, img_obj.height = {img['width']}, {img['height']}")
                script.append(f"        ws.add_image(img_obj, '{img['anchor']}')")
                script.append(
                    f"    except Exception as e: print(f' [!] Error loading image {img['filename']}: {{e}}')")

        script.append(f"    print(f\"   [+] Sheet {original_name} completed.\")\n")

        # Save Sheet module
        with open(self.output_dir / filename, "w", encoding="utf-8") as f:
            f.write("\n".join(script))

    def _write_common_styles(self):
        content = [
            "# -*- coding: utf-8 -*-\n",
            "from openpyxl.styles import Border, Side, PatternFill, Font, Alignment\n"
        ]

        for category in self.style_registry.values():
            for style_id, python_code in category.items():
                content.append(f"{style_id} = {python_code}")

        # Save Common module
        with open(self.output_dir / "common.py", "w", encoding="utf-8") as f:
            f.write("\n".join(content))

    def _write_main_reconstructor(self, sheet_modules):
        script = [
            "# -*- coding: utf-8 -*-\n",
            "import openpyxl",
            "import sys",
            "from pathlib import Path\n",
            "sys.path.append(str(Path(__file__).parent))\n"
        ]

        for mod in sheet_modules:
            script.append(f"from sheet_{mod} import rebuild_{mod}")

        script.append("\ndef main_rebuild():")
        script.append("    wb = openpyxl.Workbook()")
        script.append("    if 'Sheet' in wb.sheetnames: wb.remove(wb['Sheet'])")

        for mod in sheet_modules:
            script.append(f"    rebuild_{mod}(wb)")

        script.append("    output_file = 'full_reconstructed.xlsx'")
        script.append("    wb.save(output_file)")
        script.append("    print(f\"[!] Success: '{output_file}' generated.\")")

        script.append("\nif __name__ == '__main__':")
        script.append("    main_rebuild()")

        # Save Main module
        with open(self.output_dir / "main.py", "w", encoding="utf-8") as f:
            f.write("\n".join(script))


if __name__ == "__main__":
    print("Module to convert cell metadata or JSON, generated by the Analysis Module,"
          " into a Python script using Openpyxl."
          "\nUse:\n\tfrom exporter_py import PythonScriptExporter")
