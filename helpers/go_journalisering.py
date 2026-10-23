"""Journalisering af afgørelsesbreve i GetOrganized (GO).

Brevet lægges på barnets sag og markeres som sagsakt — og DER STOPPER DET.

Finalisering er med vilje udeladt. GO låser et færdiggjort dokument, og et
afgørelsesbrev skal kunne rettes af en sagsbehandler efter journaliseringen.
Der er derfor ingen funktion her, der kalder /_goapi/Documents/FinalizeMultiple
— ikke en parameter der som standard er slået fra, men slet ingen kode. Det
skal være et bevidst tilvalg at skrive den, ikke en værdi nogen kan komme til
at ændre.

De to kald er dem go_journalisering selv bruger:

    upload       POST /_goapi/Documents/AddToCase
    journaliser  POST /_goapi/Documents/MarkMultipleAsCaseRecord/ByDocumentId
"""

import logging
import time

import requests
from mbu_dev_shared_components.database.connection import RPAConnection
from mbu_dev_shared_components.getorganized import documents, objects

logger = logging.getLogger(__name__)


_UPLOAD_PATH = "/_goapi/Documents/AddToCase"
_JOURNALISER_PATH = "/_goapi/Documents/MarkMultipleAsCaseRecord/ByDocumentId"

# Afgørelsesbreve går til borgeren.
_KORRESPONDANCE = "Udgående"


# GO svarer en gang imellem 500 med {"Message":"An error has occurred."} på et
# kald, der virker et øjeblik senere. Beskeden siger intet, og den dækker både
# et sekunds utilgængelighed og en rigtig fejl — så det eneste, der skiller dem
# ad, er at prøve igen.
#
# Tre forsøg, 5 og derefter 10 sekunder imellem. Loftet er der, fordi en
# ProcessError stadig SKAL komme, når GO er nede for alvor: et brev, der ikke
# kom på sagen, må ikke se ud som om det gjorde, og emnet skal fejle synligt
# frem for at hænge i kø og vente på en server, der ikke kommer tilbage.
_FORSOEG = 3
_PAUSE_SEKUNDER = (5, 10)


def _credentials() -> tuple[str, str, str]:
    """(endpoint, brugernavn, kodeord) til GO, fra RPA-credential-storet.

    De samme tre værdier go_journalisering læser. Hentes pr. kørsel frem for
    at ligge i miljøet, så et skiftet kodeord virker uden en ny deployment.
    """

    with RPAConnection(db_env="PROD", commit=False) as rpa_conn:
        return (
            rpa_conn.get_constant("go_api_endpoint")["value"].rstrip("/"),
            rpa_conn.get_credential("go_api")["username"],
            rpa_conn.get_credential("go_api")["decrypted_password"],
        )


def _kald_go(beskrivelse: str, funktion, *args):
    """Kald GO, og prøv igen hvis svaret er en serverfejl.

    Args:
        beskrivelse:
            Hvad kaldet gør, som det skal stå i loggen ved et nyt forsøg.

        funktion:
            GO-funktionen fra mbu_dev_shared_components.

        *args:
            Argumenterne til den.

    Returns:
        Svaret fra det sidste forsøg. Kalderen afgør selv, om det er godt nok
        — den her rejser ikke på et dårligt svar, den prøver bare igen.

    Notes:
        KUN 5xx og netværksfejl prøves igen. En 4xx er deterministisk — et
        forkert kodeord, en sag der ikke findes, et dokument der er låst — og
        det svar bliver det samme om ti sekunder. At vente på det ville kun
        udskyde den fejl, en sagsbehandler skal se.

        Lagt om det enkelte HTTP-kald frem for om journaliser_brev som helhed.
        Slår journaliseringen fejl, er dokumentet allerede uploadet, og et nyt
        forsøg på HELE funktionen ville uploade det igen og hente
        legitimationen igen. Her prøves kun det kald, der faktisk fejlede.
    """

    for forsoeg in range(1, _FORSOEG + 1):
        sidste_forsoeg = forsoeg == _FORSOEG

        try:
            svar = funktion(*args)
        except requests.RequestException as fejl:
            # Timeout, afbrudt forbindelse, DNS. Samme slags flakken som en
            # 500, og uden et svar at kigge på.
            if sidste_forsoeg:
                raise

            grund = f"{type(fejl).__name__}: {fejl}"
        else:
            if svar.status_code < 500:
                return svar

            if sidste_forsoeg:
                return svar

            grund = f"{svar.status_code} — {svar.text[:200]}"

        pause = _PAUSE_SEKUNDER[min(forsoeg - 1, len(_PAUSE_SEKUNDER) - 1)]

        logger.warning(
            "%s fejlede (%s). Prøver igen om %d sekunder (forsøg %d af %d).",
            beskrivelse,
            grund,
            pause,
            forsoeg + 1,
            _FORSOEG,
        )

        time.sleep(pause)


def journaliser_brev(
    case_id: str,
    file_name: str,
    file_bytes: bytes,
    document_title: str,
    document_date: str,
) -> str:
    """Læg brevet på sagen i GO og markér det som sagsakt. Returnerer DocId.

    case_id er sagsnummeret som GO kender det — det samme som
    Bevilling.esdh_noegle, der følger med i brevets data som sags_nummer.

    overwrite=true: filnavnet indeholder allerede datoen, så to breve til det
    samme barn samme dag ER det samme brev lavet om. Uden overwrite ville det
    andet forsøg fejle i stedet for at erstatte.

    Det er også dét, der gør _kald_go's gentagelse ufarlig på uploaden. En 500
    fortæller ikke, om dokumentet nåede at blive oprettet, så et nyt forsøg kan
    ramme en fil, der allerede ligger der. Med overwrite bliver den erstattet
    af den samme fil; uden ville gentagelsen enten fejle eller lægge brevet på
    sagen to gange. Journaliseringen har samme egenskab af sig selv: at markere
    et dokument som sagsakt, der allerede ER markeret, ændrer ingenting.

    Rejser ved fejl frem for at logge og gå videre: et brev, der ikke kom på
    sagen, må ikke se ud som om det gjorde.
    """

    endpoint, brugernavn, kodeord = _credentials()

    document_handler = objects.DocumentJsonCreator()

    document_data = document_handler.document_data_json(
        case_id,
        "Dokumenter",
        "",
        file_name,
        '<z:row xmlns:z="#RowsetSchema" '
        f'ows_Dato="{document_date}" '
        f'ows_Title="{document_title}" '
        f'ows_Korrespondance="{_KORRESPONDANCE}" '
        "/>",
        "true",
        list(file_bytes),
    )

    svar = _kald_go(
        f"Upload af {file_name!r} til sag {case_id}",
        documents.upload_file_to_case,
        document_data,
        f"{endpoint}{_UPLOAD_PATH}",
        brugernavn,
        kodeord,
    )

    if not svar.ok:
        raise RuntimeError(
            f"Kunne ikke lægge {file_name!r} på sag {case_id} i GO: "
            f"{svar.status_code} — {svar.text[:300]}"
        )

    doc_id = svar.json()["DocId"]

    logger.info("Brevet er lagt på sag %s i GO som DocId %s.", case_id, doc_id)

    svar = _kald_go(
        f"Journalisering af DocId {doc_id} på sag {case_id}",
        documents.mark_file_as_case_record,
        [doc_id],
        f"{endpoint}{_JOURNALISER_PATH}",
        brugernavn,
        kodeord,
    )

    if not svar.ok:
        raise RuntimeError(
            f"Brevet blev lagt på sag {case_id} som DocId {doc_id}, men kunne "
            f"ikke journaliseres: {svar.status_code} — {svar.text[:300]}. "
            "Dokumentet ligger på sagen og skal journaliseres manuelt."
        )

    logger.info(
        "DocId %s er journaliseret på sag %s. Dokumentet er IKKE finaliseret "
        "og kan stadig rettes.",
        doc_id,
        case_id,
    )

    return str(doc_id)
