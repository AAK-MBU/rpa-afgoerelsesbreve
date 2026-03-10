"""Helper to have block handler functions"""

import sys

from datetime import datetime

from helpers import helper_functions


def handle_custom_koerselstyper(citizen_data: dict, input_dict: dict, block: dict):
    """
    test
    """

    koerselsraekker = citizen_data.get("koerselsraekker", {})

    # ----------------------------------------
    # Ophør overrides everything
    # ----------------------------------------

    if input_dict.get("ophoers_dato"):

        text = f"Den nuværende kørsel ophører pr. {input_dict['ophoers_dato']}."

        block["mapping"] = "Ophør"
        block["entries"] = {"Ophør": text}

        return block

    antal = len(koerselsraekker)

    # ----------------------------------------
    # Single transport type
    # ----------------------------------------

    if antal == 1:

        key, data = next(iter(koerselsraekker.items()))

        koerselstype = key
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

    lines = ["Kørslen bevilges i følgende form:"]

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
