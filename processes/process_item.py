"""Module to handle item processing"""
# from mbu_rpa_core.exceptions import ProcessError, BusinessError

import datetime
import json
import logging
import re

import requests


from helpers import config, go_journalisering, helper_functions, block_handlers
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

    decision_lower = (afgoerelsesbrev_decision or "").lower()

    # A påtænkt afgørelse announces what we INTEND to decide; it grants nothing
    # and ends nothing. Several blocks are written for letters that actually
    # decide, so the two cases have to be told apart — see blocks 7.4 and 8 in
    # block_metadata below, and the letter title further down.
    er_paataenkt = "påtænkt" in decision_lower

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


    # Used by the block engine for selecting snippets.
    #
    # DE-DUPLICATED, because the engine's "equals" branch appends one text per
    # item in the list. A bevilling normally holds several kørselsrækker of the
    # same type — morning and afternoon are two rows — so without this the
    # kørselstype paragraph was printed once per række rather than once. The
    # same goes for a tillæg granted on more than one række.
    #
    # dict.fromkeys keeps the first occurrence, and the rækker are already
    # sorted by date, so the order the snippets appear in does not change.
    custom_key_overrides["koerselstype"] = list(dict.fromkeys(koerselstype_keys))
    custom_key_overrides["koerselstype_tillaeg"] = list(dict.fromkeys(koerselstype_tillaeg))

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

    # Egen befordring: the granted one-way driving distance ("Km i bil") lives
    # on the koersel as bevilget_koereafstand_pr_vej. From it we expose two
    # top-level placeholders for the egen-befordring block:
    #   {bevilget_koereafstand_pr_vej} — one-way distance ("Km i bil")
    #   {bevilget_koereafstand_pr_dag} — max per day ("Km pr dag (bil)")
    # Per day = one-way distance × the number of one-way trips the child is in
    # the car: 2 for "Morgen og eftermiddag", otherwise 1. Danish letters use a
    # comma as the decimal separator (6,1 — not 6.1).
    # Picked on the KØRSELSTYPE, not on "has a distance". The old test was
    #
    #     bevilget_koereafstand_pr_vej not in (None, "")
    #
    # and 0 passes it. A bevilling with an egen befordring række of 36,1 km
    # and a cykelbus række of 0 km therefore matched the cykelbus — which
    # sorts first — and the letter reported "opgjort til 0 km" with a blank
    # recipient, because a cykelbus række has neither.
    #
    # The label is compared with spaces removed as well as the key, because
    # the lookup is not consistent about them ("Egen befordring" in the
    # Befordringstype table, "Egenbefordring" elsewhere).
    def _er_egen_befordring(koerselsraekke: dict) -> bool:
        key = str(koerselsraekke.get("koerselstype_key") or "").replace("-", "_")
        label = str(koerselsraekke.get("koerselstype") or "").lower().replace(" ", "")

        return key == "egen_befordring" or label == "egenbefordring"

    egen_befordring_koersel = next(
        (
            koerselsraekke
            for koerselsraekke in sorted_koerselsraekker
            if _er_egen_befordring(koerselsraekke)
        ),
        None,
    )

    if egen_befordring_koersel:
        km_pr_vej = egen_befordring_koersel.get("bevilget_koereafstand_pr_vej")
        tidspunkt = str(egen_befordring_koersel.get("tidspunkt") or "").lower()
        trips_per_day = 2 if ("morgen" in tidspunkt and "eftermiddag" in tidspunkt) else 1

        item_data["bevilget_koereafstand_pr_vej"] = helper_functions.format_danish_number(km_pr_vej)

        try:
            item_data["bevilget_koereafstand_pr_dag"] = helper_functions.format_danish_number(
                float(km_pr_vej) * trips_per_day
            )
        except (TypeError, ValueError):
            item_data["bevilget_koereafstand_pr_dag"] = ""

        # Who the kørselsgodtgørelse is actually paid to. Set per
        # kørselsrække in befordringssystemet — either a Part or a
        # forælder — and resolved to name and CPR by
        # view_Letter_Koerselsraekker. Hoisted to item_data because the
        # sentence lives in template block 6.1, which is letter-level text
        # and cannot reach into a kørselsrække.
        item_data["koerselsgodtgoerelse_modtager"] = (
            egen_befordring_koersel.get("koerselsgodtgoerelse_modtager") or ""
        )
        # Formatted as DDMMYY-XXXX, like barnets_cpr. The column stores ten
        # bare digits, and the letter is read by a citizen — a CPR without the
        # dash is not how anyone writes one.
        item_data["koerselsgodtgoerelse_modtager_cpr"] = helper_functions.format_cpr(
            egen_befordring_koersel.get("koerselsgodtgoerelse_modtager_cpr")
        )
    else:
        item_data["bevilget_koereafstand_pr_vej"] = ""
        item_data["bevilget_koereafstand_pr_dag"] = ""
        item_data["koerselsgodtgoerelse_modtager"] = ""
        item_data["koerselsgodtgoerelse_modtager_cpr"] = ""

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
            # Block 8's only entry is keyed on the bare word "Påtænkt", but the
            # decision is always "Påtænkt afslag" / "Påtænkt ophør" / "Påtænkt
            # bevilling". custom_key is an exact match, so passing the decision
            # never matched and the block was silently dropped from EVERY
            # påtænkt letter. Pass the word the entry is actually keyed on.
            #
            # None for every other letter: the block then keeps its default
            # "equals" condition, which looks the full afgørelsesbrev text up
            # among block 8's entries, finds nothing, and appends nothing.
            "8": "Påtænkt" if er_paataenkt else None,
            "9.1": klagevejledning,
            "9.2": regler,
        },
        "custom": {
            "3.1": block_handlers.handle_custom_koerselstyper,
            "4": block_handlers.handle_custom_institution,
        },
        # "Herefter revurderes bevillingen." belongs on the end of the kørsel
        # sentence when there is only ONE kørselstype — as its own paragraph
        # under a single line it reads as a stray remark. Under a bulleted
        # list of several kørselstyper it stays a paragraph of its own, which
        # is why the merge names the variant it applies to.
        "merge_into": {
            "3.2": {"target": "3.1", "when_mapping": "Én kørselstype", "separator": " "},
        },

        "copy": {
            "7.3": ["3.1", "3.2"],
        },
        "custom_contains": {
            # Block 7.4 is "Alle bevillinger" — text that belongs on a letter
            # which actually GRANTS. The match is on any word of the decision
            # appearing in the entry key, so "Påtænkt bevilling" picked it up
            # through the word "bevilling" and was treated as a bevillingsbrev.
            # It is not one: nothing is granted until the påtænkt afgørelse has
            # been through partshøring and a real afgørelse follows.
            #
            # Empty mapping rather than a removed key: create_letter skips a
            # custom_contains block whose mapping is falsy.
            "7.4": "" if er_paataenkt else afgoerelsesbrev_decision,
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

    request_data["dags_dato"] = datetime.datetime.now().strftime("%d-%m-%Y")
    request_data["skolens_navn"] = request_data.get("skole")

    # Format the child's CPR as XXXXXX-XXXX for the letter.
    request_data["barnets_cpr"] = helper_functions.format_cpr(request_data.get("barnets_cpr"))

    # All dates in the letter body should read "30. juli 2026" (Danish long
    # form). NB: dags_dato (the letterhead date) is intentionally NOT in this
    # list — it stays dd-mm-yyyy. Dates reach us in mixed formats (dd-mm-yyyy
    # from the views, ISO yyyy-mm-dd from the create-letter date pickers), so
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
        "koersel_startdato",
        "koersel_slutdato",
        "dato_for_seneste_bevilling",
        "dato_for_tidligere_afgoerelse",
        "ophoersdato",
    )

    for field in date_fields:
        if request_data.get(field):
            request_data[field] = helper_functions.format_danish_date(request_data[field])

    # Same treatment for the numbers: Danish uses a comma as the decimal
    # separator, and the value arrives from the API as a float, so "6.8 km"
    # reached the letter instead of "6,8 km".
    #
    # gaaafstand_km is the only one that needs it — it is Elev.skoleafstand,
    # the sole FLOAT among the placeholders. transporttid_i_bus and
    # skift_med_bus are integers with no decimal to separate, and
    # bevilget_koereafstand_pr_vej / _pr_dag are already formatted where the
    # egen-befordring row is read.
    #
    # Before resolve_blocks for the same reason as the dates: both the main
    # template and the block texts have to see the formatted value.
    number_fields = ("gaaafstand_km",)

    for field in number_fields:
        if request_data.get(field) not in (None, ""):
            request_data[field] = helper_functions.format_danish_number(
                request_data[field]
            )

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

    # Letter title follows the decision type:
    #   Midlertidig kørsel -> "Afgørelse om midlertidig kørsel til NAVN"
    #   Påtænkt afgørelse  -> "Påtænkt afgørelse om kørsel til NAVN"
    #   Everything else    -> "Afgørelse om kørsel til NAVN"
    #
    # The date is appended to the file name so a second letter for the same
    # child does not overwrite the first. Without it the name was identical for
    # every letter of the same type to the same child, which is why only one
    # could exist at a time.
    #
    # Same-day letters still collide by design — the caseworker deletes the
    # earlier file when replacing it. Dropping seconds keeps the name readable
    # and matches how the files are looked for.
    if er_paataenkt:
        letter_title = f"Påtænkt afgørelse om kørsel til {barnets_fulde_navn}"
    elif "midlertidig" in decision_lower:
        letter_title = f"Afgørelse om midlertidig kørsel til {barnets_fulde_navn}"
    else:
        letter_title = f"Afgørelse om kørsel til {barnets_fulde_navn}"

    # dd-mm-yyyy, the same form as the letterhead date (dags_dato).
    file_name_date = datetime.datetime.now().strftime("%d-%m-%Y")

    for file_type in ["docx"]:
        file_name = f"{letter_title} - {file_name_date}.{file_type}"

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

        # Journalisér brevet på barnets sag i GO.
        #
        # Afløser en upload til et SharePoint-bibliotek: brevet hører til på
        # sagen, ikke i en mappe ved siden af den, og på sagen er det synligt
        # for alle der arbejder med barnet.
        #
        # Dokumentet FINALISERES IKKE. GO låser et færdiggjort dokument, og et
        # afgørelsesbrev skal kunne rettes bagefter — se
        # helpers/go_journalisering.py, hvor der slet ikke findes kode til at
        # finalisere.
        sags_nummer = str(item_data.get("sags_nummer") or "").strip()

        if not sags_nummer:
            raise ValueError(
                "Brevet kan ikke journaliseres: bevillingen har intet "
                "sags_nummer (esdh_noegle), så der er ingen sag i GO at lægge "
                "det på."
            )

        go_journalisering.journaliser_brev(
            case_id=sags_nummer,
            file_name=file_name,
            file_bytes=file_bytes,
            document_title=letter_title,
            document_date=file_name_date,
        )
