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

from mbu_dev_shared_components.database.connection import RPAConnection
from mbu_dev_shared_components.getorganized import documents, objects

logger = logging.getLogger(__name__)


_UPLOAD_PATH = "/_goapi/Documents/AddToCase"
_JOURNALISER_PATH = "/_goapi/Documents/MarkMultipleAsCaseRecord/ByDocumentId"

# Afgørelsesbreve går til borgeren.
_KORRESPONDANCE = "Udgående"


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

    svar = documents.upload_file_to_case(
        document_data, f"{endpoint}{_UPLOAD_PATH}", brugernavn, kodeord
    )

    if not svar.ok:
        raise RuntimeError(
            f"Kunne ikke lægge {file_name!r} på sag {case_id} i GO: "
            f"{svar.status_code} — {svar.text[:300]}"
        )

    doc_id = svar.json()["DocId"]

    logger.info("Brevet er lagt på sag %s i GO som DocId %s.", case_id, doc_id)

    svar = documents.mark_file_as_case_record(
        [doc_id], f"{endpoint}{_JOURNALISER_PATH}", brugernavn, kodeord
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
