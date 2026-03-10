"""Module for general configurations of the process"""

import os

MAX_RETRY = 10

# ----------------------
# Queue population settings
# ----------------------
MAX_CONCURRENCY = 100  # tune based on backend capacity
MAX_RETRIES = 3  # transient failure retries per item
RETRY_BASE_DELAY = 0.5  # seconds (exponential backoff)

# SharePoint stuff
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com"

SHAREPOINT_SITE_NAME = "MBURPA"

DOCUMENT_LIBRARY = "Delte dokumenter"

SHAREPOINT_KWARGS = {
    "tenant": os.getenv("TENANT"),
    "client_id": os.getenv("CLIENT_ID"),
    "thumbprint": os.getenv("APPREG_THUMBPRINT"),
    "cert_path": os.getenv("GRAPH_CERT_PEM"),
    "site_url": f"{SHAREPOINT_SITE_URL}/",
    "site_name": SHAREPOINT_SITE_NAME,
    "document_library": DOCUMENT_LIBRARY,
}

FOLDER_NAME = "Egenbefordring/Afgørelsesbreve"
