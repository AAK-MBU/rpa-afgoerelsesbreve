"""Custom block handler functions used by the skabelonmotor."""

from helpers import helper_functions


def _format_koerselsraekke(data: dict) -> str:
    """
    Format a single kørselsrække, e.g.:
        "Skånekørsel morgen [mandag, onsdag, fredag] fra 01-03-2026 til 01-07-2027"

    - Tidspunkt follows the kørselstype (lowercased) whenever it is filled.
    - Weekdays follow in brackets (lowercased), omitted when "Alle".
    """

    koerselstype = (
        data.get("koerselstype")
        or data.get("koerselstype_key")
        or "kørsel"
    )

    start = data.get("bevilling_fra")
    slut = data.get("bevilling_til")
    tidspunkt = data.get("tidspunkt")
    dage = data.get("dage")

    tidspunkt_text = ""
    if tidspunkt:
        tidspunkt_text = f" {tidspunkt.lower()}"

    dage_text = ""
    if dage and dage.lower() != "alle":
        dage_text = f" [{dage.lower()}]"

    return f"{koerselstype}{tidspunkt_text}{dage_text} fra {start} til {slut}"


def handle_custom_koerselstyper(item_data: dict, block: dict):
    """
    Generate dynamic text for the "Kørselstype" block based on the transport rows
    in item_data["koerselsraekker"].

    Supports multiple kørselsrækker with the same kørsels-/befordringstype.
    """

    koerselsraekker = item_data.get("koerselsraekker") or []

    # ----------------------------------------
    # Afslag overrides everything
    # ----------------------------------------
    afgoerelsesbrev = item_data.get("afgoerelsesbrev")
    afgoerelsesbrev_decision = (
        afgoerelsesbrev.split(":", 1)[0].strip()
        if afgoerelsesbrev
        else None
    )

    if afgoerelsesbrev_decision == "Afslag":
        block["mapping"] = "Afslag"

        return block

    # ----------------------------------------
    # Ophør
    # ----------------------------------------

    if item_data.get("ophoersdato"):
        text = f"Den nuværende kørsel ophører pr. {item_data['ophoersdato']}."

        block["mapping"] = "Ophør"
        block["entries"] = {"Ophør": text}

        return block

    antal = len(koerselsraekker)

    # ----------------------------------------
    # No transport rows
    # ----------------------------------------

    if antal == 0:
        block["mapping"] = "Ingen kørselstype"
        block["entries"] = {
            "Ingen kørselstype": ""
        }

        return block

    # Sort rows by start date, end date, type name, and ID
    sorted_koerselsraekker = sorted(
        koerselsraekker,
        key=lambda row: (
            helper_functions.parse_date(row.get("bevilling_fra")),
            helper_functions.parse_date(row.get("bevilling_til")),
            str(row.get("koerselstype") or row.get("koerselstype_key") or "").lower(),
            row.get("koersel_id") or 0,
        )
    )

    # ----------------------------------------
    # Single transport row
    # ----------------------------------------

    if antal == 1:
        text = f"Kørslen bevilges i form af {_format_koerselsraekke(sorted_koerselsraekker[0])}."

        block["mapping"] = "Én kørselstype"
        block["entries"] = {"Én kørselstype": text}

        return block

    # ----------------------------------------
    # Multiple transport rows
    # ----------------------------------------

    # Intro line, then one paragraph per kørselsrække. Each list item is
    # prefixed with the "[[LIST_ITEM]]" marker (and paragraphs are separated by
    # blank lines), which the skabelonmotor renderer turns into a real Word
    # bullet (the "List Bullet" style / punktopstilling). A literal "•" only
    # renders as a plain dot, not a Word list.
    parts = ["Kørslen bevilges i følgende form:"]

    for data in sorted_koerselsraekker:
        parts.append(f"[[LIST_ITEM]]{_format_koerselsraekke(data)}.")

    text = "\n\n".join(parts)

    block["mapping"] = "Flere kørselstyper"
    block["entries"] = {"Flere kørselstyper": text}

    return block


def handle_custom_blok_7_3(item_data: dict, block: dict):
    """Handle Blok 7.3.

    This block should always include the text for "Alle breve".

    If the afgørelsesbrev decision is "Bevilling", it should also include
    the text for "Alle bevillinger".
    """

    afgoerelsesbrev = item_data.get("afgoerelsesbrev")

    afgoerelsesbrev_decision = (
        afgoerelsesbrev.split(":", 1)[0].strip()
        if afgoerelsesbrev
        else None
    )

    entries = block.get("entries", {})

    selected_texts = []

    alle_breve_text = entries.get("Alle breve")
    alle_bevillinger_text = entries.get("Alle bevillinger")

    if alle_breve_text:
        selected_texts.append(alle_breve_text)

    if afgoerelsesbrev_decision == "Bevilling" and alle_bevillinger_text:
        selected_texts.append(alle_bevillinger_text)

    block["mapping"] = "Blok 7.3"
    block["entries"] = {
        "Blok 7.3": "\n\n".join(selected_texts)
    }

    return block
