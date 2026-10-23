"""Custom block handler functions used by the skabelonmotor."""

from helpers import helper_functions


def _format_koerselsraekke(data: dict) -> str:
    """
    Format a single kørselsrække, e.g.:
        "Skånekørsel fast forsæde mellem hjem og skole 12,5 km pr. vej
         [mandag, onsdag, fredag] fra 01-03-2026 til 01-07-2027"

    The order is Befordring's, from the "Flere kørselstyper" row of their
    template sheet:

        {koerselstype} {tillæg} {tidspunkt} [{bevilget_koereafstand_pr_vej}
        km pr. vej] [{dage}] fra {koersel_startdato} til {koerselstype_slutdato}

    EVERY part between the kørselstype and the dates is conditional. Only the
    type and the period are always known; a række may carry no tillæg, no
    rutetype, no distance, and "Alle" for its days. Each segment therefore
    carries its own leading space and collapses to "" when it has nothing to
    say, rather than being joined — joining is what leaves double spaces and
    stray brackets in a letter a family reads.

    Rutetype sits where TIDSPUNKT stands in the sheet. Befordring's own note
    on the restructure says why: "Tidspunkt er blevet overflødigt med
    ovenstående formuleringer" — the six rutetype names state where a trip
    starts and ends, and therefore when it happens. "Fra hjem til skole" IS
    the morning; "Mellem hjem og skole" is both legs. Printing both would say
    the same thing twice, and tidspunkt is on its way out of the database.

    Tillæg, rutetype and dage are lowercased because they sit inside a
    sentence. The kørselstype keeps its capital: every row is a bullet
    starting its own line. It was lowercased once, for a one-row sentence that
    ran on from "i form af" — that variant no longer exists.

    The tillæg here is NOT a duplicate of block 6.2. This states what was
    granted; 6.2 appends a paragraph explaining what each tillæg means. The
    two are the same word doing different jobs, and 6.2 de-duplicates across
    rækker where this is deliberately per-række.
    """

    koerselstype = (
        data.get("koerselstype")
        or data.get("koerselstype_key")
        or "kørsel"
    )

    start = helper_functions.format_danish_date(data.get("bevilling_fra"))
    slut = helper_functions.format_danish_date(data.get("bevilling_til"))
    tillaeg = data.get("koerselstype_tillaeg")
    rutetype = data.get("rutetype")
    dage = data.get("dage")

    # Already a comma-separated list per række — view_Letter_Koerselsraekker
    # STRING_AGGs the tillæg link table — so it needs no joining here.
    tillaeg_text = ""
    if tillaeg and str(tillaeg).strip():
        tillaeg_text = f" {str(tillaeg).strip().lower()}"

    rutetype_text = ""
    if rutetype:
        rutetype_text = f" {rutetype.lower()}"

    # Zero is treated as absent, not as a distance.
    #
    # It is the value a række carries when no distance applies — a cykelbus
    # række is the known case, and the same 0 once made block 6 pick the
    # cykelbus over a 36,1 km egen befordring række and report "opgjort til
    # 0 km". "0 km pr. vej" in an afgørelsesbrev is a statement about the
    # family's journey, and it would be a false one.
    km_text = ""
    km = helper_functions.format_danish_number(data.get("bevilget_koereafstand_pr_vej"))

    if km and km not in ("0", "0,0"):
        km_text = f" {km} km pr. vej"

    dage_text = ""
    if dage and dage.lower() != "alle":
        dage_text = f" [{dage.lower()}]"

    return (
        f"{koerselstype}{tillaeg_text}{rutetype_text}{km_text}{dage_text}"
        f" fra {start} til {slut}"
    )


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
    # The granted rows
    # ----------------------------------------

    # One row renders exactly like several: the intro line, then a bulleted
    # list. Befordring dropped the separate one-row sentence ("Kørslen
    # bevilges i form af ...") so that every bevillingsbrev reads the same way
    # whatever it grants — which also means the shape of the letter no longer
    # changes when a kørselsrække is added to or removed from a bevilling.
    #
    # Intro line, then one paragraph per kørselsrække. Each list item is
    # prefixed with the "[[LIST_ITEM]]" marker (and paragraphs are separated by
    # blank lines), which the skabelonmotor renderer turns into a real Word
    # bullet (the "List Bullet" style / punktopstilling). A literal "•" only
    # renders as a plain dot, not a Word list.
    parts = ["Kørslen bevilges i følgende form:"]

    for data in sorted_koerselsraekker:
        parts.append(f"[[LIST_ITEM]]{_format_koerselsraekke(data)}.")

    text = "\n\n".join(parts)

    # Still keyed "Flere kørselstyper" with one row. The name is Befordring's
    # own, from the template sheet, and it is the only granting variant left.
    block["mapping"] = "Flere kørselstyper"
    block["entries"] = {"Flere kørselstyper": text}

    return block


# The klub whose text differs from the generic SFO text. Compared
# case-insensitively against the item_data["institution"] value.


def _find_entry_key(entries: dict, token: str):
    """Return the first entry key that contains `token` (case-insensitive)."""

    for key in entries:
        if token in key.strip().lower():
            return key

    return None


def handle_custom_institution(item_data: dict, block: dict):
    """Handle Blok 4 (institution — SFO / klub).

    item_data["institution"] holds the institution the student is attached to
    (e.g. "SFO - Holme Skole" or "Klubben Holme Søndergaard"), or is empty when
    the student has no SFO/klub. A plain has_value check is no longer enough,
    because the block now has two different texts (SFO vs. Klubben Holme
    Søndergaard). Rules:

    - Afslag -> nothing. We never grant SFO/klub transport in a rejection.
    - No institution value -> nothing.
    - No kørselsrække marked "Kørsel til institution: Ja" -> nothing. The
      student may well have an SFO, but if no granted kørsel goes there the
      letter should not say anything about it.
    - institution names "Klubben Holme Søndergaard" -> the klub entry.
    - institution has any other value -> the generic SFO entry.

    Selection sets block["mapping"] to the matching entry key so the "custom"
    renderer appends that entry's text (or nothing when mapping is None).
    """

    afgoerelsesbrev = item_data.get("afgoerelsesbrev")
    afgoerelsesbrev_decision = (
        afgoerelsesbrev.split(":", 1)[0].strip()
        if afgoerelsesbrev
        else None
    )

    entries = block.get("entries", {})

    # Lower-case both sides for a lenient comparison, per the field's free-text
    # institution names.
    institution_value = (item_data.get("institution") or "").strip()
    normalized_institution = institution_value.lower()

    # Two conditions have to hold before the SFO/klub text goes in the letter:
    # the student must HAVE an institution, and the granted kørsel must actually
    # go TO it. Having an SFO is not on its own a reason to write about
    # transport to it.
    #
    # view_Letter_Koerselsraekker resolves koersel_til_institution to the
    # strings "Ja" / "Nej" / NULL, so this is a text comparison. NB: the field
    # is only filled for taxa-like kørselstyper — the application nulls it for
    # egenbefordring and skolerejsekort, so a bevilling made up entirely of
    # those has no "Ja" anywhere and gets no SFO block.
    koersel_til_institution = any(
        str(koerselsraekke.get("koersel_til_institution") or "").strip().lower() == "ja"
        for koerselsraekke in item_data.get("koerselsraekker") or []
    )

    # Still suppressed on Afslag: we never grant SFO/klub transport in a
    # rejection. "Nej" is treated as "no value" to match the template's
    # negative case.
    if (
        afgoerelsesbrev_decision == "Afslag"
        or not institution_value
        or normalized_institution == "nej"
        or not koersel_til_institution
    ):
        block["mapping"] = None

        return block

    klubben_holme = "klubben holme søndergaard"

    if klubben_holme in normalized_institution:
        block["mapping"] = _find_entry_key(entries, "klubben")

    else:
        block["mapping"] = _find_entry_key(entries, "sfo")

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
