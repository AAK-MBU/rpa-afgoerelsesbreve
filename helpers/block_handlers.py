"""Custom block handler functions used by the skabelonmotor."""

from helpers import helper_functions


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
        data = sorted_koerselsraekker[0]

        koerselstype = (
            data.get("koerselstype")
            or data.get("koerselstype_key")
            or "kørsel"
        )

        start = data.get("bevilling_fra")
        slut = data.get("bevilling_til")
        tidspunkt = data.get("tidspunkt")
        dage = data.get("dage")

        extras = []

        if tidspunkt and tidspunkt.lower() != "morgen og eftermiddag":
            extras.append(tidspunkt)

        if dage and dage.lower() != "alle":
            # Weekdays with a lowercase initial (mandag, onsdag, …).
            extras.append(dage.lower())

        extra_text = f" [{', '.join(extras)}]" if extras else ""

        text = (
            f"Kørslen bevilges i form af {koerselstype}"
            f"{extra_text} fra {start} til {slut}."
        )

        block["mapping"] = "Én kørselstype"
        block["entries"] = {"Én kørselstype": text}

        return block

    # ----------------------------------------
    # Multiple transport rows
    # ----------------------------------------

    lines = ["Kørslen bevilges i følgende form:"]

    for data in sorted_koerselsraekker:
        koerselstype = (
            data.get("koerselstype")
            or data.get("koerselstype_key")
            or "kørsel"
        )

        start = data.get("bevilling_fra")
        slut = data.get("bevilling_til")
        tidspunkt = data.get("tidspunkt")
        dage = data.get("dage")

        extras = []

        if tidspunkt and tidspunkt.lower() != "morgen og eftermiddag":
            extras.append(tidspunkt)

        if dage and dage.lower() != "alle":
            # Weekdays with a lowercase initial (mandag, onsdag, …).
            extras.append(dage.lower())

        extra_text = f" [{', '.join(extras)}]" if extras else ""

        lines.append(
            f"• {koerselstype}{extra_text} fra {start} til {slut}."
        )

    text = "\n".join(lines)

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
