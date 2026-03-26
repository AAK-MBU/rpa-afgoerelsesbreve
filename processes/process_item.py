"""Module to handle item processing"""
# from mbu_rpa_core.exceptions import ProcessError, BusinessError

import json
import logging
import re

import requests

from helpers import helper_functions, block_handlers

logger = logging.getLogger(__name__)

BLOCK_HEADER_PATTERN = re.compile(r"^Blok\s+([0-9]+(?:\.\s*[0-9]+)?[a-zA-Z]?)")


def process_item(item_data: dict, item_reference: str):
    """Function to handle item processing"""

    assert item_data, "Item data is required"
    assert item_reference, "Item reference is required"

    # Initialize an empty dict to contain key overrides
    custom_key_overrides = {}

    # Retrieve the childs full name and parse the first name - afterwards add it to the item_dict as it is used as a placeholder in the template letter texts
    barnets_fulde_navn = item_data.get("barnets_fulde_navn")
    barnets_fornavn = barnets_fulde_navn.split()[0] if barnets_fulde_navn else ""
    item_data["barnets_fornavn"] = barnets_fornavn

    # Retrieve the hjaelpemidler - the key is a string but we need to convert it to a list of hjaelpemidler so the skabelonmotor can properly identify necessary placeholder texts to include
    hjaelpemidler_raw = item_data.get("hjaelpemidler")
    hjaelpemidler = [item.strip() for item in hjaelpemidler_raw.split(",")] if hjaelpemidler_raw else []
    custom_key_overrides["hjaelpemidler"] = hjaelpemidler

    # The template texts sometimes use only the decision part of the afgoerelsesbrev key, therefore we extract it into a separate value - it's later used as a custom key for several blocks
    afgoerelsesbrev = item_data.get("afgoerelsesbrev")
    afgoerelsesbrev_decision = (
        afgoerelsesbrev.split(":", 1)[0].strip()
        if afgoerelsesbrev
        else None
    )

    # The snippet below is responsible for a couple things:
    # 1. We extract koerselsraekker and sort them by their start and end dates, so that we can initialize a koersel_startdato key, that is the start date of the earliest koerselstype
    # 2. We create a list of koerselstyper, that is used in the skabelonmotor to correctly identify which text snippets to use with regards to koerselstyper
    # 3. We do the same for koerselstype_tillaeg
    koerselsraekker = item_data.get("koerselsraekker") or {}
    sorted_koerselstyper = sorted(
        koerselsraekker.items(),
        key=lambda item: (
            helper_functions.parse_date(item[1].get("bevilling_fra")),
            helper_functions.parse_date(item[1].get("bevilling_til")),
            item[0].lower(),
        )
    )
    koerselstype = []
    koerselstype_tillaeg = []
    if sorted_koerselstyper:
        for i, (koerselstype_key, koerselstype_data) in enumerate(sorted_koerselstyper):
            if i == 0:
                # Here we set the koersel_startdato mentioned in point 1 above
                item_data["koersel_startdato"] = koerselstype_data.get("bevilling_fra")

            koerselstype.append(koerselstype_key)

            raw = koerselstype_data.get("koerselstype_tillaeg")
            if raw:
                koerselstype_tillaeg.extend(
                    item.strip() for item in raw.split(",")
                )
    # Here we set the koerselstype key override - by setting it like this, we ensure the skabelonmotor doesn't use the incorrectly formatted key from the citizen's data, but instead this properly formatted key
    custom_key_overrides["koerselstype"] = koerselstype
    # Here we do the same for koerselstype_tillaeg
    custom_key_overrides["koerselstype_tillaeg"] = koerselstype_tillaeg

    # We create 2 custom variables, used as custom keys to correctly handle block 9.1 and 9.2 in the template text data
    if "midlertidig" in str(afgoerelsesbrev).lower():
        klagevejledning = "Klagevejledning brækket ben ungdomsuddannelse"

    else:
        klagevejledning = "Klagevejledning"

    if afgoerelsesbrev == "Afslag: § 33, stk. 3 (ungdomsskolen)":
        regler = "Regler § 33, stk. 3 (ungdomsskoleloven)"

    elif "midlertidig" in str(afgoerelsesbrev).lower():
        regler = "Regler brækket ben ungdomssuddanelse"

    else:
        regler = "Regler standard"

    # This metadata is used to handle various scenarios where the template text data is not simply selected by mapping the mapping_key to a text entry
    block_metadata = {
        "has_value": [
            "1.2",
            "3.2",
            "4",
        ],
        "custom": {
            "3.1": block_handlers.handle_custom_koerselstyper,
        },
        "custom_key": {
            "5": afgoerelsesbrev_decision,
            "8": afgoerelsesbrev_decision,
            "9.1": klagevejledning,
            "9.2": regler,
        },
        "copy": {
            "7.3": "3.1",
        },
        "all": [
            "7.4",
        ],
    }

    request_data = item_data

    # This query is used to fetch the template data from our table of template data rows
    # We use an updated database instead of the actual docx/excel files to circumvent potential issues with regards to locked MSOffice files
    query = """
        SELECT TOP 1
            process_name,
            word_template,
            workbook_json
        FROM rpa.Templates
        WHERE process_name = :process_name
        ORDER BY last_updated DESC;
    """

    params = {
        "process_name": "afgoerelsesbreve"
    }

    df = helper_functions.read_sql(
        query=query,
        params=params,
        conn_string=helper_functions.get_db_connection_string()
    )

    if df.empty:
        raise Exception("No template found for process")

    row = df.iloc[0]

    # Retrieve the docx template and replace any placeholders
    template_binary_docx = row["word_template"]
    template_b64 = helper_functions.replace_template_placeholders(template_bytes=template_binary_docx, data=request_data)

    # Retrieve the template block data and handle any blocks that are specified in block_metadata dictionary
    blocks = json.loads(row["workbook_json"])
    resolved_blocks = helper_functions.resolve_blocks(blocks=blocks, block_metadata=block_metadata, item_data=item_data)

    # print()
    # print()
    # print()
    # print(resolved_blocks)
    # print()
    # print()
    # print()
    # sys.exit()

    print()

    for file_type in ["docx", "pdf"]:
        request = {
            "data": request_data,
            "block_data": resolved_blocks,
            "custom_key_overrides": custom_key_overrides,
            "file_type": file_type,
            "template_b64": template_b64,
        }

        url = "http://127.0.0.1:8000/skabelonmotor/api/letter_creation/create_text"

        response = requests.post(url, json=request, timeout=10)
        response.raise_for_status()

        file_bytes = response.content

        file_name = f"test_letter.{file_type}"

        with open(file_name, "wb") as f:
            f.write(file_bytes)
