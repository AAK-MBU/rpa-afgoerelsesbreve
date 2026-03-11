"""
Utility functions used by the skabelonmotor Excel parser.
"""

import re

import copy

from datetime import datetime

from io import BytesIO

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText


# Regex used to detect block headers such as:
# "Blok 1", "Blok 3.1", "Blok 7.2a"
BLOCK_HEADER_PATTERN = re.compile(r"^Blok\s+([0-9]+(?:\.\s*[0-9]+)?[a-zA-Z]?)")


def parse_date(value: str | None):
    """
    Convert a DD-MM-YYYY string into a datetime object.

    Used primarily for sorting transport rows by start/end date.

    Args:
        value (str | None): Date string in DD-MM-YYYY format.

    Returns:
        datetime: Parsed date or datetime.max if value is missing.
    """

    if not value:
        return datetime.max

    return datetime.strptime(value, "%d-%m-%Y")


def extract_cell_formatting(cell):
    """
    Convert Excel rich text content into HTML-like formatted text.

    Handles formatting such as bold, italic, underline, strike-through
    and color while preserving the original text structure.

    Args:
        cell: openpyxl cell object.

    Returns:
        str: HTML-like formatted text.
    """

    if cell is None or cell.value is None:
        return ""

    value = cell.value

    # ----------------------------------------
    # Rich formatted text (Excel rich text)
    # ----------------------------------------
    if isinstance(value, CellRichText):

        parts = []

        for block in value:

            text = block.text or ""
            font = block.font

            if not text:
                continue

            # Remove zero-width characters sometimes inserted by Excel
            text = text.replace("\u200b", "")

            prefix = ""
            suffix = ""

            if font:

                # Bold
                if font.b:
                    prefix += "<strong>"
                    suffix = "</strong>" + suffix

                # Italic
                if font.i:
                    prefix += "<em>"
                    suffix = "</em>" + suffix

                # Underline
                if font.u in ["single", "double", "singleAccounting", "doubleAccounting", True]:
                    prefix += "<u>"
                    suffix = "</u>" + suffix

                # Strikethrough
                if font.strike:
                    prefix += "<strike>"
                    suffix = "</strike>" + suffix

                # Text color
                if font.color and font.color.rgb:
                    rgb = font.color.rgb[-6:]

                    # Skip default black text
                    if rgb != "000000":
                        prefix += f'<span style="color:#{rgb}">'
                        suffix = "</span>" + suffix

            parts.append(f"{prefix}{text}{suffix}")

        return "".join(parts)

    # ----------------------------------------
    # Plain text cell (no formatting)
    # ----------------------------------------
    return str(value)


def parse_workbook(citizen_data: dict, input_dict: dict, binary_excel: bytes, block_metadata: dict) -> list[dict]:
    """
    Parse the Excel template into structured block definitions used by the skabelonmotor.

    The workbook is scanned for rows starting with "Blok X", which define logical
    text blocks in the template. Each block is converted into a dictionary containing
    a block id, mapping key, condition type and its possible text entries.

    Parsing happens in three main stages:
    1. Build lookup tables from `block_metadata` so each block knows its behavior
       (e.g. equals, has_value, custom, copy).
    2. Iterate through sheets and rows to detect block headers and collect their
       entries from the Excel template.
    3. Post-process blocks by executing custom handlers and resolving "copy"
       blocks that inherit content from other blocks.

    Args:
        citizen_data (dict): Citizen data used by custom block handlers.
        input_dict (dict): Additional input values used during block generation.
        binary_excel (bytes): Excel workbook content.
        block_metadata (dict): Configuration describing block conditions and handlers.

    Returns:
        list[dict]: Parsed block structures ready for the template engine.
    """

    wb = load_workbook(BytesIO(binary_excel), rich_text=True)

    parsed_blocks = []

    condition_lookup = {}
    custom_function_lookup = {}
    custom_key_lookup = {}
    copy_lookup = {}

    # ----------------------------------------
    # Build condition lookup tables
    # ----------------------------------------
    # Converts block_metadata into quick lookup dictionaries so each block can easily determine its behavior.
    for condition, blocks in block_metadata.items():

        if condition == "custom":

            for block_id, func in blocks.items():

                condition_lookup[block_id] = "custom"
                custom_function_lookup[block_id] = func

        elif condition == "custom_key":

            for block_id, value in blocks.items():

                condition_lookup[block_id] = "custom_key"
                custom_key_lookup[block_id] = value

        elif condition == "copy":

            for block_id, source_block in blocks.items():

                condition_lookup[block_id] = "copy"
                copy_lookup[block_id] = source_block

        else:

            for block_id in blocks:
                condition_lookup[block_id] = condition

    current_block = {}

    # ----------------------------------------
    # Parse workbook sheets
    # ----------------------------------------
    for sheet_name in wb.sheetnames:
        # Ignore sheets not starting with "Blok"
        if not sheet_name.startswith("Blok"):
            continue

        ws = wb[sheet_name]
        rows = list(ws.iter_rows())

        for i, row in enumerate(rows):

            col_a_cell = row[0] if len(row) > 0 else None
            col_b_cell = row[1] if len(row) > 1 else None

            col_a = col_a_cell.value if col_a_cell else None
            col_b = extract_cell_formatting(col_b_cell) if col_b_cell else None

            # ----------------------------------------
            # Detect block header
            # ----------------------------------------
            if isinstance(col_a, str):
                match = BLOCK_HEADER_PATTERN.match(col_a)

                if match:
                    # Finish previous block if it required custom processing
                    if current_block and current_block["condition"] == "custom":

                        func = custom_function_lookup.get(current_block["block_id"])

                        if func:
                            func(
                                citizen_data,
                                input_dict,
                                current_block
                            )

                    block_id = match.group(1).replace(" ", "").strip()

                    # Mapping key is defined in column C of the next row
                    next_row = rows[i + 1] if i + 1 < len(rows) else None
                    next_col_c = None

                    if next_row and len(next_row) > 2:
                        next_col_c = next_row[2].value

                    condition = condition_lookup.get(block_id, "equals")

                    mapping = str(next_col_c).strip() if next_col_c else None

                    # custom_key behaves like a custom block but with a predefined mapping value
                    if condition == "custom_key":
                        mapping = custom_key_lookup.get(block_id)
                        condition = "custom"

                    current_block = {
                        "block_id": block_id,
                        "title": col_a,
                        "mapping": mapping,
                        "condition": condition,
                        "entries": {}
                    }

                    parsed_blocks.append(current_block)

                    continue

            if not current_block:
                continue

            # ----------------------------------------
            # Parse block entries
            # ----------------------------------------
            if col_a and col_b:
                entry_text = col_b.strip()

                # Skip placeholder rows like "Ingen tekst"
                if normalize_key(entry_text) == "ingentekst":
                    continue

                key = str(col_a)

                current_block["entries"][key] = entry_text

        # ----------------------------------------
        # Finalize last block in sheet
        # ----------------------------------------
        if current_block and current_block["condition"] == "custom":
            func = custom_function_lookup.get(current_block["block_id"])

            if func:
                func(
                    citizen_data,
                    input_dict,
                    current_block
                )

    # ----------------------------------------
    # Apply "copy" block behavior
    # ----------------------------------------
    block_map = {b["block_id"]: b for b in parsed_blocks}

    for block in parsed_blocks:

        if block["condition"] == "copy":

            source_id = copy_lookup.get(block["block_id"])
            source_block = block_map.get(source_id)

            if source_block:

                block["mapping"] = source_block["mapping"]
                block["entries"] = copy.deepcopy(source_block["entries"])
                block["condition"] = source_block["condition"]

    return parsed_blocks


def normalize_key(value: str) -> str:
    """
    Normalize text for reliable key comparison.

    Removes whitespace, punctuation and Danish characters
    so template keys can be compared consistently.

    Args:
        value (str): Input string.

    Returns:
        str: Normalized key.
    """

    return (
        value.strip()
        .lower()
        .replace(" ", "")
        .replace(".", "")
        .replace("ø", "oe")
        .replace("å", "aa")
        .replace("æ", "ae")
    )
