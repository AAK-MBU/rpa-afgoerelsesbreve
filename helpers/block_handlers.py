"""Custom block handler functions used by the skabelonmotor."""

from helpers import helper_functions


def handle_custom_koerselstyper(citizen_data: dict, input_dict: dict, block: dict):
    """
    Generate dynamic text for the "Kørselstype" block based on the transport rows in citizen_data["koerselsraekker"].

    Args:
        citizen_data (dict): Citizen data containing transport rows.
        input_dict (dict): Additional inputs (e.g. termination date).
        block (dict): Parsed block that will be modified.

    Returns:
        dict: Updated block with generated mapping and entries.
    """

    # All configured transport rows for the child
    koerselsraekker = citizen_data.get("koerselsraekker", {})

    # ----------------------------------------
    # Ophør overrides everything
    # ----------------------------------------
    # If a termination date exists we ignore transport rows and generate a simple termination sentence.
    if input_dict.get("ophoers_dato"):

        text = f"Den nuværende kørsel ophører pr. {input_dict['ophoers_dato']}."

        block["mapping"] = "Ophør"
        block["entries"] = {"Ophør": text}

        return block

    antal = len(koerselsraekker)

    # ----------------------------------------
    # Single transport type
    # ----------------------------------------
    # If only one transport row exists, generate one sentence.
    if antal == 1:
        key, data = next(iter(koerselsraekker.items()))

        koerselstype = key
        start = data.get("bevilling_fra")
        slut = data.get("bevilling_til")
        tidspunkt = data.get("tidspunkt")
        dage = data.get("dage")

        extras = []

        # Include tidspunkt if it is not the default full-day transport
        if tidspunkt and tidspunkt.lower() != "morgen og eftermiddag":
            extras.append(tidspunkt)

        # Include specific days if transport is not valid for "Alle"
        if dage and dage.lower() != "alle":
            extras.append(dage)

        extra_text = f" [{', '.join(extras)}]" if extras else ""

        text = (
            f"Kørslen bevilges i form af {koerselstype}"
            f"{extra_text} fra {start} til {slut}."
        )

        block["mapping"] = "Én kørselstype"
        block["entries"] = {"Én kørselstype": text}

        return block

    # ----------------------------------------
    # Multiple transport types
    # ----------------------------------------
    # If several rows exist, create a bullet list describing each.
    lines = ["Kørslen bevilges i følgende form:"]

    # Sort rows by start date, end date, and type name
    sorted_koerselstyper = sorted(
        koerselsraekker.items(),
        key=lambda item: (
            helper_functions.parse_date(item[1].get("bevilling_fra")),
            helper_functions.parse_date(item[1].get("bevilling_til")),
            item[0].lower(),
        )
    )

    for key, data in sorted_koerselstyper:
        start = data.get("bevilling_fra")
        slut = data.get("bevilling_til")
        tidspunkt = data.get("tidspunkt")
        dage = data.get("dage")

        extras = []

        if tidspunkt and tidspunkt != "Morgen og Eftermiddag":
            extras.append(tidspunkt)

        if dage and dage.lower() != "alle":
            extras.append(dage)

        extra_text = f" [{', '.join(extras)}]" if extras else ""

        lines.append(
            f"• {key}{extra_text} fra {start} til {slut}."
        )

    text = "\n".join(lines)

    block["mapping"] = "Flere kørselstyper"
    block["entries"] = {"Flere kørselstyper": text}

    return block
