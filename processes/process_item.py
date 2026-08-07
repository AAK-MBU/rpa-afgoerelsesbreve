"""Module to handle item processing"""
# from mbu_rpa_core.exceptions import ProcessError, BusinessError

import datetime
import json
import logging
import re

import requests

from mbu_msoffice_integration.sharepoint_class import Sharepoint

from helpers import config, helper_functions, block_handlers
# 🔥 TEMPORARY - remove together with helpers/mock_skabelonmotor.py when the API is live
from helpers import mock_skabelonmotor

logger = logging.getLogger(__name__)

BLOCK_HEADER_PATTERN = re.compile(r"^Blok\s+([0-9]+(?:\.\s*[0-9]+)?[a-zA-Z]?)")


def process_item(item_data: dict, item_reference: str):
    """Function to handle item processing"""

    assert item_data, "Item data is required"
    assert item_reference, "Item reference is required"

    item_data["barnets_cpr"] = helper_functions.format_cpr(item_data.get("barnets_cpr"))

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
    # 1. We extract koerselsraekker and sort them by their start and end dates, so that we can initialize a koersel_slutdato key, that is the end date of the latest koerselstype
    # 2. We create a list of koerselstyper, that is used in the skabelonmotor to correctly identify which text snippets to use with regards to koerselstyper
    # 3. We do the same for koerselstype_tillaeg
    # Note: koersel_startdato is a manual field supplied from the create-letter
    # modal and is intentionally NOT computed here.
    koerselsraekker = item_data.get("koerselsraekker") or []

    sorted_koerselsraekker = sorted(
        koerselsraekker,
        key=lambda row: (
            helper_functions.parse_date(row.get("bevilling_fra")),
            helper_functions.parse_date(row.get("bevilling_til")),
            str(row.get("koerselstype_key") or "").lower(),
            row.get("koersel_id") or 0,
        )
    )

    koerselstype_keys = []
    koerselstype_labels = []
    koerselstype_tillaeg = []

    if sorted_koerselsraekker:
        latest_koerselsraekke = max(
            sorted_koerselsraekker,
            key=lambda row: helper_functions.parse_date(row.get("bevilling_til"))
        )

        item_data["koersel_slutdato"] = latest_koerselsraekke.get("bevilling_til")

        for koerselsraekke in sorted_koerselsraekker:
            koerselstype_key = koerselsraekke.get("koerselstype_key")
            koerselstype_label = koerselsraekke.get("koerselstype")

            if koerselstype_key:
                koerselstype_keys.append(koerselstype_key)

            if koerselstype_label:
                koerselstype_labels.append(koerselstype_label)

            raw_tillaeg = koerselsraekke.get("koerselstype_tillaeg")

            if raw_tillaeg:
                koerselstype_tillaeg.extend(
                    item.strip()
                    for item in raw_tillaeg.split(",")
                    if item.strip()
                )


    # Used by the block engine for selecting snippets
    custom_key_overrides["koerselstype"] = koerselstype_keys
    custom_key_overrides["koerselstype_tillaeg"] = koerselstype_tillaeg

    # Used by normal placeholder replacement: {koerselstype}
    unique_koerselstype_labels = list(dict.fromkeys(koerselstype_labels))

    item_data["koerselstype"] = ", ".join(unique_koerselstype_labels)

    # Skolerejsekort details (transporttid i bus / antal skift) now live on the
    # koersel instead of being typed in at letter time. Take the first non-empty
    # value found across the koerselsraekker and expose each as a top-level
    # variable, so the template can use the {transporttid_i_bus} / {skift_med_bus}
    # placeholders.
    item_data["transporttid_i_bus"] = next(
        (
            str(koerselsraekke.get("transporttid_i_bus"))
            for koerselsraekke in sorted_koerselsraekker
            if koerselsraekke.get("transporttid_i_bus") not in (None, "")
        ),
        "",
    )

    # antal skift: take the first value present across the koerselsraekker.
    raw_skift = next(
        (
            koerselsraekke.get("skift_med_bus")
            for koerselsraekke in sorted_koerselsraekker
            if koerselsraekke.get("skift_med_bus") not in (None, "")
        ),
        None,
    )

    # The template reads "{skift_med_bus} skift". When nothing is entered, or 0
    # is entered, the caseworkers want it to read "uden skift" rather than
    # "0 skift" / a blank. Any positive count renders as the number ("2 skift").
    if raw_skift in (None, "", 0, "0"):
        item_data["skift_med_bus"] = "uden"
    else:
        item_data["skift_med_bus"] = str(raw_skift)

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
        ],
        "custom_key": {
            "1.1": item_data.get("brev_i_forbindelse_med"),
            "2.2": item_data.get("befordringsudvalg_resultat"),
            "5": afgoerelsesbrev_decision,
            "8": afgoerelsesbrev_decision,
            "9.1": klagevejledning,
            "9.2": regler,
        },
        "custom": {
            "3.1": block_handlers.handle_custom_koerselstyper,
            "4": block_handlers.handle_custom_sfo,
        },
        "copy": {
            "7.3": ["3.1", "3.2"],
        },
        "custom_contains": {
            "7.4": afgoerelsesbrev_decision,
        },
        "all": [
            "7.5",
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
        FROM
            rpa.Templates
        WHERE
            process_name = :process_name
        ORDER BY
            last_updated DESC;
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

    request_data["DOKUMENTNUMMER"] = "12325"
    request_data["dags_dato"] = datetime.datetime.now().strftime("%d-%m-%Y")
    request_data["skolens_navn"] = request_data.get("skole")

    # Format the child's CPR as XXXXXX-XXXX for the letter.
    request_data["barnets_cpr"] = helper_functions.format_cpr(request_data.get("barnets_cpr"))

    # All dates in the letter body should read "30. juli 2026" (Danish long
    # form). The header date in the main template is handled separately in the
    # template itself. Dates reach us in mixed formats (dd-mm-yyyy from the
    # views/dags_dato, ISO yyyy-mm-dd from the create-letter date pickers), so
    # format_danish_date parses both and leaves anything unparseable untouched.
    # This runs before the template placeholder replace AND before
    # resolve_blocks, so both the main template and the block texts get the
    # formatted dates.
    date_fields = (
        "modtagelsesdato",
        "sagsbehandlingsdato",
        "revurdering",
        "befordringsudvalg",
        "afstandskriterie_dato",
        "dags_dato",
        "koersel_startdato",
        "koersel_slutdato",
        "dato_for_seneste_bevilling",
        "dato_for_tidligere_afgoerelse",
        "ophoersdato",
    )

    for field in date_fields:
        if request_data.get(field):
            request_data[field] = helper_functions.format_danish_date(request_data[field])

    # NB: the kørselsrække start/end dates (bevilling_fra/bevilling_til) are
    # intentionally NOT reformatted here — they are still sorted with
    # parse_date (which expects dd-mm-yyyy) inside the kørselstype block
    # handler. They are formatted for display in block_handlers
    # (_format_koerselsraekke) instead.

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
    # import sys
    # sys.exit()

    # Initialize the SharePoint connection once - it is reused for every file we upload
    sharepoint = Sharepoint(**config.SHAREPOINT_KWARGS)

    for file_type in ["docx"]:
        file_name = f"{barnets_fulde_navn}_{request_data["dags_dato"]}.{file_type}"

        # ╔══════════════════════════════════════════════════════════════════╗
        # ║ 🔥 TEMPORARY MOCK - api-skabelonmotor is not yet live 🔥          ║
        # ║ While the API is not dockerised/online we build the letter        ║
        # ║ in-process via helpers.mock_skabelonmotor. When the API is         ║
        # ║ deployed, delete helpers/mock_skabelonmotor.py and restore the     ║
        # ║ HTTP call below.                                                   ║
        # ╚══════════════════════════════════════════════════════════════════╝
        file_bytes = mock_skabelonmotor.create_letter(
            data=request_data,
            block_data=resolved_blocks,
            custom_key_overrides=custom_key_overrides,
            file_type=file_type,
            file_name=file_name,
            template_b64=template_b64,
        )

        # --- ORIGINAL API CALL (restore when api-skabelonmotor is live) ---
        # request = {
        #     "data": request_data,
        #     "block_data": resolved_blocks,
        #     "custom_key_overrides": custom_key_overrides,
        #     "file_type": file_type,
        #     "file_name": file_name,
        #     "template_b64": template_b64,
        # }
        #
        # url = "http://localhost:8020/letter_creation/create_letter"
        #
        # response = requests.post(url, json=request, timeout=60)
        # response.raise_for_status()
        #
        # file_bytes = response.content

        # Upload the created letter to SharePoint instead of saving it locally
        sharepoint.upload_file_from_bytes(
            binary_content=file_bytes,
            file_name=file_name,
            folder_name=config.FOLDER_NAME,
        )
