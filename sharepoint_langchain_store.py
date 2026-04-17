"""
SharePoint folder-based LangChain PGVector store.

Accepts a SharePoint folder path, resolves document IDs from the postgres
database that match that folder, and returns a configured PGVectorStore or
retriever scoped to those documents.

DB structure mirrors langchain_chat_components.py:
  view: langchainpg_chat_view
  key columns: id, text, embedding_chat, file_path, document_id, ...

Required env vars:
  LANGCHAIN_DB                    - postgres connection URL
  AZURE_OPENAI_ENDPOINT           - Azure OpenAI endpoint
  AZURE_OPENAI_KEY                - Azure OpenAI API key
  AZURE_OPENAI_EMBEDDING_DEPLOYMENT - embedding deployment name (default: text-embedding-3-small)
"""

import os
import logging
from typing import List, Optional

import psycopg
from langchain_core.documents import Document
from langchain_core.embeddings import Embeddings
from langchain_core.retrievers import BaseRetriever
from langchain_postgres import PGEngine, PGVectorStore

logger = logging.getLogger(__name__)

_CREATE_VIEW_SQL = """
CREATE VIEW IF NOT EXISTS langchainpg_chat_view AS
SELECT
    c.id,
    c.text,
    c.embedding_chat,
    c.chunk_index,
    c.tokens,
    c.metadata,
    d.id AS document_id,
    d.file_path,
    d.uploaded_in_session_id
FROM dispatch_brain_documentchunk c
INNER JOIN dispatch_brain_document d ON c.document_id = d.id
WHERE c.embedding_chat IS NOT NULL;
"""

class _EmptyRetriever(BaseRetriever):
    """Returned when a SharePoint folder has no embedded documents yet."""

    def _get_relevant_documents(self, query: str, **kwargs) -> List[Document]:
        return []

    async def _aget_relevant_documents(self, query: str, **kwargs) -> List[Document]:
        return []


_METADATA_COLUMNS = [
    "metadata",
    "document_id",
    "file_path",
    "chunk_index",
    "uploaded_in_session_id",
]


class AzureOpenAIEmbeddings(Embeddings):
    """LangChain-compatible wrapper for Azure OpenAI 1536d text-embedding-3-small."""

    def __init__(self):
        import openai

        self._client = openai.AzureOpenAI(
            api_key=os.environ["AZURE_OPENAI_KEY"],
            api_version="2024-02-01",
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        )
        self._deployment = os.environ.get(
            "AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "text-embedding-3-small"
        )
        logger.info("AzureOpenAIEmbeddings using deployment: %s", self._deployment)

    def embed_documents(self, texts: List[str]) -> List[List[float]]:
        response = self._client.embeddings.create(input=texts, model=self._deployment)
        return [item.embedding for item in response.data]

    def embed_query(self, text: str) -> List[float]:
        response = self._client.embeddings.create(input=[text], model=self._deployment)
        return response.data[0].embedding


def get_sharepoint_folder_path(folder: str) -> str:
    """
    Validate and return the SharePoint folder path.

    The path is prefix-matched against the file_path column in the DB, so pass
    the portion of the path that all documents in that folder share, e.g.:
      '/sites/MySite/Shared Documents/Reports'
      'https://company.sharepoint.com/sites/MySite/Shared Documents/Reports'

    Args:
        folder: SharePoint folder path or URL prefix.

    Returns:
        Normalised path string (trailing slash stripped).
    """
    path = folder.rstrip("/")
    if not path:
        raise ValueError("SharePoint folder path must not be empty")
    return path


def _ensure_view(db_url: str) -> None:
    with psycopg.connect(db_url) as conn:
        with conn.cursor() as cur:
            cur.execute(_CREATE_VIEW_SQL)
        conn.commit()


def _document_ids_for_folder(folder_path: str, db_url: str) -> List[str]:
    with psycopg.connect(db_url) as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT DISTINCT document_id FROM langchainpg_chat_view "
                "WHERE file_path LIKE %s",
                (f"{folder_path}%",),
            )
            rows = cur.fetchall()
    doc_ids = [str(row[0]) for row in rows]
    logger.info(
        "SharePoint folder '%s' → %d document(s) in DB", folder_path, len(doc_ids)
    )
    return doc_ids


def get_sharepoint_folder_vectorstore(
    folder: str,
    db_url: Optional[str] = None,
) -> PGVectorStore:
    """
    Return a PGVectorStore backed by all embedded documents under a SharePoint folder.

    Args:
        folder: SharePoint folder path (prefix-matched against file_path in DB).
        db_url: Postgres connection URL. Defaults to LANGCHAIN_DB env var.

    Returns:
        Configured PGVectorStore. The matched folder path and document IDs are
        attached as ``_sharepoint_folder_path`` and ``_sharepoint_document_ids``
        for callers that need to build their own filters.
    """
    folder_path = get_sharepoint_folder_path(folder)
    db_url = db_url or os.environ["LANGCHAIN_DB"]

    _ensure_view(db_url)

    doc_ids = _document_ids_for_folder(folder_path, db_url)
    if not doc_ids:
        logger.warning(
            "No embedded documents found under SharePoint folder: %s", folder_path
        )

    pg_engine = PGEngine.from_connection_string(url=db_url)
    embedding = AzureOpenAIEmbeddings()

    store = PGVectorStore.create_sync(
        engine=pg_engine,
        table_name="langchainpg_chat_view",
        embedding_service=embedding,
        id_column="id",
        content_column="text",
        embedding_column="embedding_chat",
        metadata_columns=_METADATA_COLUMNS,
    )

    # Attach resolved metadata so callers can reference it without re-querying
    store._sharepoint_folder_path = folder_path
    store._sharepoint_document_ids = doc_ids

    logger.info(
        "SharePoint vector store ready: folder=%s docs=%d", folder_path, len(doc_ids)
    )
    return store


def get_sharepoint_folder_retriever(
    folder: str,
    db_url: Optional[str] = None,
    top_k: int = 20,
    fetch_k: int = 100,
):
    """
    Return an MMR retriever scoped to documents under a SharePoint folder.

    Args:
        folder: SharePoint folder path.
        db_url: Postgres connection URL. Defaults to LANGCHAIN_DB env var.
        top_k: Final number of chunks returned.
        fetch_k: Candidate pool size for MMR diversity pass.

    Returns:
        LangChain retriever with document_id filter pre-applied.
    """
    store = get_sharepoint_folder_vectorstore(folder, db_url)
    doc_ids = store._sharepoint_document_ids

    if not doc_ids:
        logger.warning(
            "Returning empty retriever for SharePoint folder with no documents: %s",
            store._sharepoint_folder_path,
        )
        return _EmptyRetriever()

    return store.as_retriever(
        search_type="mmr",
        search_kwargs={
            "k": top_k,
            "fetch_k": fetch_k,
            "filter": {"document_id": {"$in": doc_ids}},
        },
    )
