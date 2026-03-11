"""
Docstring for afgørelsesbreve_folder.helper_functions
"""

import sys

import re

import copy

from datetime import datetime

from io import BytesIO

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText

BLOCK_HEADER_PATTERN = re.compile(r"^Blok\s+([0-9]+(?:\.\s*[0-9]+)?[a-zA-Z]?)")


def parse_date(value: str | None):
    """Convert DD-MM-YYYY string to datetime for sorting."""

    if not value:
        return datetime.max

    return datetime.strptime(value, "%d-%m-%Y")


def extract_cell_formatting(cell):
    """
    Convert Excel rich text into formatted HTML-like string.
    """

    if cell is None or cell.value is None:
        return ""

    value = cell.value

    # ----------------------------------------
    # Rich formatted text
    # ----------------------------------------

    if isinstance(value, CellRichText):

        parts = []

        for block in value:

            text = block.text or ""
            font = block.font

            if not text:
                continue

            # remove invisible characters Excel sometimes adds
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

                # Color
                if font.color and font.color.rgb:
                    rgb = font.color.rgb[-6:]

                    if rgb != "000000":
                        prefix += f'<span style="color:#{rgb}">'
                        suffix = "</span>" + suffix

            parts.append(f"{prefix}{text}{suffix}")

        return "".join(parts)

    # ----------------------------------------
    # Plain text cell
    # ----------------------------------------

    return str(value)


def parse_workbook(citizen_data: dict, input_dict: dict, binary_excel: bytes, block_metadata: dict) -> list[dict]:
    """
    Parse Excel workbook into structured block data.
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
    # Parse workbook
    # ----------------------------------------

    for sheet_name in wb.sheetnames:

        if not sheet_name.startswith("Blok"):
            continue

        ws = wb[sheet_name]

        rows = list(ws.iter_rows())

        for i, row in enumerate(rows):

            col_a_cell = row[0] if len(row) > 0 else None
            col_b_cell = row[1] if len(row) > 1 else None
            col_c_cell = row[2] if len(row) > 2 else None

            col_a = col_a_cell.value if col_a_cell else None
            col_b = extract_cell_formatting(col_b_cell) if col_b_cell else None

            # ----------------------------------------
            # Detect block header
            # ----------------------------------------

            if isinstance(col_a, str):

                match = BLOCK_HEADER_PATTERN.match(col_a)

                if match:

                    # ----------------------------------------
                    # Finish previous block before starting new
                    # ----------------------------------------

                    if current_block and current_block["condition"] == "custom":

                        func = custom_function_lookup.get(current_block["block_id"])

                        if func:
                            func(
                                citizen_data,
                                input_dict,
                                current_block
                            )

                    block_id = match.group(1).replace(" ", "").strip()

                    next_row = rows[i + 1] if i + 1 < len(rows) else None
                    next_col_c = None

                    if next_row and len(next_row) > 2:
                        next_col_c = next_row[2].value

                    condition = condition_lookup.get(block_id, "equals")

                    mapping = str(next_col_c).strip() if next_col_c else None

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
            # Parse entries
            # ----------------------------------------

            if col_a and col_b:

                entry_text = col_b.strip()

                if normalize_key(entry_text) == "ingentekst":
                    continue

                key = str(col_a)

                current_block["entries"][key] = entry_text

        # ----------------------------------------
        # After sheet ends, finalize last block
        # ----------------------------------------

        if current_block and current_block["condition"] == "custom":

            func = custom_function_lookup.get(current_block["block_id"])

            if func:
                func(
                    citizen_data,
                    input_dict,
                    current_block
                )

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


def replace_input_placeholders(letter_text: str, citizen_data: dict, input_data: dict):
    """
    Docstring for replace_input_placeholders

    :param letter_text: Description
    :type letter_text: str
    :param citizen_data: Description
    :type citizen_data: dict
    """

    barnets_fulde_navn = citizen_data.get("barnets_fulde_navn")

    barnets_fornavn = barnets_fulde_navn.split()[0] if barnets_fulde_navn else ""

    barnets_cpr = citizen_data.get("barnets_cpr")

    status = citizen_data.get("status")

    folkeregisteraddresse = citizen_data.get("folkeregisteraddresse")

    skole = citizen_data.get("skole")

    skolematrikel = citizen_data.get("skolematrikel")

    gaaafstand_km = citizen_data.get("gaaafstand_km")

    klasseart = citizen_data.get("klasseart")

    klassebetegnelse = citizen_data.get("klassebetegnelse")

    personligt_klassetrin = citizen_data.get("personligt_klassetrin")

    sfo = citizen_data.get("sfo")

    bopaelsdistrikt = citizen_data.get("bopaelsdistrikt")

    sagsbehandlingsdato = citizen_data.get("sagsbehandlingsdato")

    hjaelpemidler = citizen_data.get("hjaelpemidler")

    adresse_for_bevilling = citizen_data.get("adresse_for_bevilling")

    afstandskriterie_dato = citizen_data.get("afstandskriterie_dato")

    afstandskriterie_klassetrin = citizen_data.get("afstandskriterie_klassetrin")

    ansoeger_relation = citizen_data.get("ansoeger_relation")

    revurdering = citizen_data.get("revurdering")

    befordringsudvalg = citizen_data.get("befordringsudvalg")

    hjemmel = citizen_data.get("hjemmel")

    afgoerelsesbrev = citizen_data.get("afgoerelsesbrev")

    sagsbehandler = citizen_data.get("sagsbehandler")

    ppr_ansvarlig = citizen_data.get("ppr_ansvarlig")

    koerselsraekker = citizen_data.get("koerselsraekker")

    ophoersdato = input_data.get("ophoersdato")

    if not koerselsraekker:
        koersel_startdato = None
        koersel_slutdato = None

    elif len(koerselsraekker) == 1:
        first_value = next(iter(koerselsraekker.values()))

        koersel_startdato = first_value.get("bevilling_fra")
        koersel_slutdato = first_value.get("bevilling_til")

    else:
        koersel_startdato = min(
            koerselsraekker.values(),
            key=lambda k: parse_date(k.get("bevilling_fra"))
        ).get("bevilling_fra")

        koersel_slutdato = max(
            koerselsraekker.values(),
            key=lambda k: parse_date(k.get("bevilling_til"))
        ).get("bevilling_til")

    letter_text = letter_text.replace("Dato (modtagelse af ansøgning)", sagsbehandlingsdato)

    letter_text = letter_text.replace("Barnets fulde navn", barnets_fulde_navn)

    letter_text = letter_text.replace("Barnets fornavn", barnets_fornavn)

    letter_text = letter_text.replace("Barnets-cpr", barnets_cpr)

    letter_text = letter_text.replace("Nuværende klassetrin", input_data.get("nuvaerende_klassetrin"))

    letter_text = letter_text.replace("Skolens navn", skole)

    letter_text = letter_text.replace("D.D.", revurdering)

    letter_text = letter_text.replace("Kørsel startdato", koersel_startdato)

    letter_text = letter_text.replace("Kørsel slutdato", koersel_slutdato)

    letter_text = letter_text.replace("Dato for befordringsudvalgsmøde", befordringsudvalg)

    letter_text = letter_text.replace("Folkeregisteradresse", adresse_for_bevilling)

    letter_text = letter_text.replace("Afstandskriterie dato", afstandskriterie_dato)

    letter_text = letter_text.replace("Afstandskriterie klassetrin", afstandskriterie_klassetrin)

    letter_text = letter_text.replace("km gå", gaaafstand_km)

    letter_text = letter_text.replace("Sagsbehandlers navn", sagsbehandler)

    letter_text = letter_text.replace("{ophoersdato}", ophoersdato)

    return letter_text


def normalize_key(value: str) -> str:
    """
    Docstring for normalize_key

    :param value: Description
    :type value: str
    :return: Description
    :rtype: str
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
