"""Module to hande queue population"""

import sys

import re

import asyncio
import json
import logging

import requests

from automation_server_client import Workqueue

from mbu_dev_shared_components.database.connection import RPAConnection

from mbu_msoffice_integration.sharepoint_class import Sharepoint

from helpers import config, helper_functions, block_handlers

logger = logging.getLogger(__name__)

BLOCK_HEADER_PATTERN = re.compile(r"^Blok\s+([0-9]+(?:\.\s*[0-9]+)?[a-zA-Z]?)")

citizen_data = {
    "barnets_fulde_navn": "Kasper Hansentest",

    "barnets_cpr": "230115-5000",

    "status": "aktiv",

    "folkeregisteradresse": "Gade 1, 8000 Aarhus C",

    "skole": "Langagerskolen",

    "skolematrikel": "Kolt Østervej 45",

    "gaaafstand_km": "10",

    "klasseart": "modtagerklasse",

    "klassebetegnelse": "M1",

    "personligt_klassetrin": "2",

    "sfo": "SFO - Holme skole",

    "bopaelsdistrikt": "Holme skole",

    "sagsbehandlingsdato": "26-11-2025",

    "adresse_for_bevilling": "Gade 1, 8000 Aarhus C",

    "hjaelpemidler": "Kørestol, Magnetsele",

    "afstandskriterie_dato": "01-07-2026",

    "afstandskriterie_klassetrin": "3",

    "ansoeger_relation": "Forældremyndighed",

    "revurdering": "30-06-2026",

    "befordringsudvalg": "30-06-2026",

    "hjemmel": "§ 26, stk. 1 afstand",

    "afgoerelsesbrev": "Bevilling: § 26, stk. 1, nr. 1 (afstand)",

    "sagsbehandler": "Sofie Elrum",

    "ppr_ansvarlig": "Klaus",

    "koerselsraekker": {
        "rutekoersel": {
            "tidspunkt": "Morgen",
            "koerselstype_tillaeg": "Fast forsæde",
            "bevilget_koereafstand_pr_vej": "10",
            "dage": "Alle",
            "bevilling_fra": "01-01-2026",
            "bevilling_til": "01-01-2027",
            "taxa_id": "",
        },
        "skolerejsekort": {
            "tidspunkt": "Eftermiddag",
            "koerselstype_tillaeg": "Co-driver, Fast sæde",
            "bevilget_koereafstand_pr_vej": "10",
            "dage": "Alle",
            "bevilling_fra": "01-01-2026",
            "bevilling_til": "01-01-2027",
            "taxa_id": "",
        },
        # "test": {
        #     "tidspunkt": "morgen",
        #     "koerselstype_tillaeg": [""],
        #     "bevilget_koereafstand_pr_vej": "10",
        #     "dage": "tirsdag",
        #     "bevilling_fra": "01-01-2026",
        #     "bevilling_til": "01-01-2027",
        #     "taxa_id": "",
        # },
    },

    "modtagelsesdato": "21-11-2025"

}

input_dict = {

    # KUN HVIS DER ER TALE OM ET OPHØR
    "ophoers_dato": "",

}

custom_key_overrides = {}

barnets_fulde_navn = citizen_data.get("barnets_fulde_navn")

barnets_fornavn = barnets_fulde_navn.split()[0] if barnets_fulde_navn else ""
citizen_data["barnets_fornavn"] = barnets_fornavn

hjaelpemidler_raw = citizen_data.get("hjaelpemidler")
hjaelpemidler = [item.strip() for item in hjaelpemidler_raw.split(",")] if hjaelpemidler_raw else []
custom_key_overrides["hjaelpemidler"] = hjaelpemidler

revurdering = citizen_data.get("revurdering")

afgoerelsesbrev = citizen_data.get("afgoerelsesbrev")
afgoerelsesbrev_decision = (
    afgoerelsesbrev.split(":", 1)[0].strip()
    if afgoerelsesbrev
    else None
)

koerselsraekker = citizen_data.get("koerselsraekker") or {}
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
            citizen_data["koersel_startdato"] = koerselstype_data.get("bevilling_fra")

        koerselstype.append(koerselstype_key)

        raw = koerselstype_data.get("koerselstype_tillaeg")
        if raw:
            koerselstype_tillaeg.extend(
                item.strip() for item in raw.split(",")
            )
custom_key_overrides["koerselstype"] = koerselstype
custom_key_overrides["koerselstype_tillaeg"] = koerselstype_tillaeg

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


def retrieve_items_for_queue() -> list[dict]:
    """Function to populate queue"""

    data = []
    references = []

    print(koerselsraekker)

    sharepoint = Sharepoint(**config.SHAREPOINT_KWARGS)

    test_binary_excel = sharepoint.fetch_file_using_open_binary(
        file_name="Afgørelsesbreve.xlsx",
        folder_name=config.FOLDER_NAME
    )

    blocks = helper_functions.parse_workbook(citizen_data=citizen_data, input_dict=input_dict, binary_excel=test_binary_excel, block_metadata=block_metadata)

    # print()
    # print()
    # print()
    # print(blocks)
    # print()
    # print()
    # print()
    # sys.exit()

    print()

    request_data = {**citizen_data, **input_dict}

    for file_type in ["pdf", "docx"]:
        request = {
            "data": request_data,
            "block_data": blocks,
            "custom_key_overrides": custom_key_overrides,
            "file_type": file_type
        }

        url = "http://127.0.0.1:8000/skabelonmotor/api/create_text"

        response = requests.post(url, json=request, timeout=10)
        response.raise_for_status()

        file_bytes = response.content

        file_name = f"test_letter.{file_type}"

        with open(file_name, "wb") as f:
            f.write(file_bytes)

    sys.exit()

    items = [
        {"reference": ref, "data": d} for ref, d in zip(references, data, strict=True)
    ]

    return items


def create_sort_key(item: dict) -> str:
    """
    Create a sort key based on the entire JSON structure.
    Converts the item to a sorted JSON string for consistent ordering.
    """
    return json.dumps(item, sort_keys=True, ensure_ascii=False)


async def concurrent_add(workqueue: Workqueue, items: list[dict]) -> None:
    """
    Populate the workqueue with items to be processed.
    Uses concurrency and retries with exponential backoff.

    Args:
        workqueue (Workqueue): The workqueue to populate.
        items (list[dict]): List of items to add to the queue.

    Returns:
        None

    Raises:
        Exception: If adding an item fails after all retries.
    """
    sem = asyncio.Semaphore(config.MAX_CONCURRENCY)

    async def add_one(it: dict):
        reference = str(it.get("reference") or "")
        data = {"item": it}

        async with sem:
            for attempt in range(1, config.MAX_RETRIES + 1):
                try:
                    await asyncio.to_thread(workqueue.add_item, data, reference)
                    logger.info("Added item to queue with reference: %s", reference)
                    return True

                except Exception as e:
                    if attempt >= config.MAX_RETRIES:
                        logger.error(
                            "Failed to add item %s after %d attempts: %s",
                            reference,
                            attempt,
                            e,
                        )
                        return False

                    backoff = config.RETRY_BASE_DELAY * (2 ** (attempt - 1))

                    logger.warning(
                        "Error adding %s (attempt %d/%d). Retrying in %.2fs... %s",
                        reference,
                        attempt,
                        config.MAX_RETRIES,
                        backoff,
                        e,
                    )
                    await asyncio.sleep(backoff)

    if not items:
        logger.info("No new items to add.")
        return

    sorted_items = sorted(items, key=create_sort_key)
    logger.info(
        "Processing %d items sorted by complete JSON structure", len(sorted_items)
    )

    results = await asyncio.gather(*(add_one(i) for i in sorted_items))
    successes = sum(1 for r in results if r)
    failures = len(results) - successes

    logger.info(
        "Summary: %d succeeded, %d failed out of %d", successes, failures, len(results)
    )
