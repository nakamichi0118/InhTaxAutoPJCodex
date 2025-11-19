"""Pydantic models for document processing endpoints."""
from __future__ import annotations

from typing import Dict, List, Literal, Optional

from pydantic import BaseModel, Field

DocumentType = Literal["bank_deposit", "land", "building", "transaction_history", "unknown"]


class TransactionLine(BaseModel):
    transaction_date: Optional[str] = Field(default=None, description="ISO 8601 date")
    description: Optional[str] = None
    withdrawal_amount: Optional[float] = None
    deposit_amount: Optional[float] = None
    balance: Optional[float] = None
    line_confidence: Optional[float] = None
    correction_note: Optional[str] = None


class AssetRecord(BaseModel):
    category: DocumentType
    type: Optional[str] = None
    source_document: str
    owner_name: List[str] = Field(default_factory=list)
    asset_name: Optional[str] = None
    location_prefecture: Optional[str] = None
    location_municipality: Optional[str] = None
    location_detail: Optional[str] = None
    identifier_primary: Optional[str] = None
    identifier_secondary: Optional[str] = None
    valuation_basis: Optional[str] = None
    valuation_currency: str = "JPY"
    valuation_amount: Optional[float] = None
    valuation_date: Optional[str] = None
    ownership_share: Optional[float] = None
    notes: Optional[str] = None
    transactions: List[TransactionLine] = Field(default_factory=list)

    def to_export_payload(self) -> dict:
        location = {
            "prefecture": self.location_prefecture,
            "municipality": self.location_municipality,
            "detail": self.location_detail,
        }
        identifiers = {
            "primary": self.identifier_primary,
            "secondary": self.identifier_secondary,
        }
        valuation = {
            "basis": self.valuation_basis,
            "currency": self.valuation_currency,
            "amount": self.valuation_amount,
            "date": self.valuation_date,
        }
        return {
            "category": self.category,
            "type": self.type,
            "source_document": self.source_document,
            "owner_name": self.owner_name,
            "asset_name": self.asset_name,
            "location": location,
            "identifiers": identifiers,
            "valuation": valuation,
            "ownership_share": self.ownership_share,
            "notes": self.notes,
            "transactions": [txn.model_dump() for txn in self.transactions],
        }


class DocumentAnalyzeResponse(BaseModel):
    status: Literal["ok"]
    document_type: DocumentType
    raw_lines: List[str]
    assets: List[AssetRecord]


class DocumentAnalyzeRequest(BaseModel):
    document_type: Optional[DocumentType] = None


class JobCreateResponse(BaseModel):
    status: Literal["accepted"]
    job_id: str


class JobStatusResponse(BaseModel):
    job_id: str
    status: Literal["pending", "running", "completed", "failed"]
    stage: Optional[str] = None
    detail: Optional[str] = None
    document_type: Optional[DocumentType] = None
    processed_chunks: Optional[int] = None
    total_chunks: Optional[int] = None
    files: Optional[Dict[str, str]] = None
    created_at: float
    updated_at: float


class JobResultResponse(BaseModel):
    status: Literal["ok"]
    job_id: str
    document_type: DocumentType
    files: Optional[Dict[str, str]] = None
    assets: List[AssetRecord]
    transactions: List["TransactionExport"] = Field(default_factory=list)


class TransactionExport(BaseModel):
    transaction_date: Optional[str] = None
    description: Optional[str] = None
    withdrawal_amount: int = 0
    deposit_amount: int = 0
    balance: Optional[int] = None
    memo: Optional[str] = None
