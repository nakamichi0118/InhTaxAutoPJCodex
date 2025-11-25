"""Pydantic models for document processing endpoints."""
from __future__ import annotations

from typing import Dict, List, Literal, Optional

from pydantic import BaseModel, Field, ConfigDict

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


class LedgerSessionRequest(BaseModel):
    app_id: Optional[str] = Field(default="ledger-app")
    session_token: Optional[str] = None


class LedgerSessionResponse(BaseModel):
    status: Literal["ok"]
    app_id: str
    user_id: str
    session_token: str


class LedgerAccountPayload(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    id: str
    case_id: Optional[str] = Field(default=None, alias="caseId")
    name: str
    number: Optional[str] = None
    holder_name: Optional[str] = Field(default=None, alias="holderName")
    order: int = 0
    user_id: Optional[str] = Field(default=None, alias="userId")
    created_at: Optional[str] = Field(default=None, alias="createdAt")
    updated_at: Optional[str] = Field(default=None, alias="updatedAt")


class LedgerAccountCreateRequest(BaseModel):
    name: str
    number: Optional[str] = None
    holder_name: Optional[str] = Field(default=None, alias="holderName")
    order: Optional[int] = None
    case_id: Optional[str] = Field(default=None, alias="caseId")


class LedgerAccountUpdateRequest(BaseModel):
    name: Optional[str] = None
    number: Optional[str] = None
    holder_name: Optional[str] = Field(default=None, alias="holderName")
    order: Optional[int] = None


class LedgerAccountOrderItem(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    account_id: str = Field(alias="id")
    order: int


class LedgerAccountReorderRequest(BaseModel):
    items: List[LedgerAccountOrderItem]


class LedgerTransactionPayload(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    id: str
    account_id: str = Field(alias="accountId")
    date: str
    withdrawal: int = 0
    deposit: int = 0
    memo: Optional[str] = None
    type: Optional[str] = None
    row_color: Optional[str] = Field(default=None, alias="rowColor")
    user_order: Optional[float] = Field(default=None, alias="userOrder")
    tags: List[str] = Field(default_factory=list)
    user_id: Optional[str] = Field(default=None, alias="userId")
    created_at: Optional[str] = Field(default=None, alias="createdAt")
    updated_at: Optional[str] = Field(default=None, alias="updatedAt")


class LedgerTransactionCreateRequest(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    account_id: str = Field(alias="accountId")
    date: str
    withdrawal: Optional[int] = 0
    deposit: Optional[int] = 0
    memo: Optional[str] = None
    type: Optional[str] = None
    user_order: Optional[float] = Field(default=None, alias="userOrder")
    row_color: Optional[str] = Field(default=None, alias="rowColor")
    tags: Optional[List[str]] = Field(default=None)


class LedgerTransactionUpdateRequest(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    account_id: Optional[str] = Field(default=None, alias="accountId")
    date: Optional[str] = None
    withdrawal: Optional[int] = None
    deposit: Optional[int] = None
    memo: Optional[str] = None
    type: Optional[str] = None
    row_color: Optional[str] = Field(default=None, alias="rowColor")
    user_order: Optional[float] = Field(default=None, alias="userOrder")
    tags: Optional[List[str]] = Field(default=None)


class LedgerTransactionOrderItem(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    transaction_id: str = Field(alias="id")
    user_order: float = Field(alias="userOrder")


class LedgerTransactionsReorderRequest(BaseModel):
    items: List[LedgerTransactionOrderItem]


class LedgerCasePayload(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    id: str
    name: str
    user_id: str = Field(alias="userId")
    created_at: Optional[str] = Field(default=None, alias="createdAt")
    updated_at: Optional[str] = Field(default=None, alias="updatedAt")


class LedgerCaseCreateRequest(BaseModel):
    name: str


class LedgerStateResponse(BaseModel):
    status: Literal["ok"]
    case: LedgerCasePayload
    cases: List[LedgerCasePayload]
    accounts: List[LedgerAccountPayload]
    transactions: List[LedgerTransactionPayload]


class LedgerImportRequest(BaseModel):
    case_id: str = Field(alias="caseId")
    accounts: List[LedgerAccountPayload]
    transactions: List[LedgerTransactionPayload]


class LedgerExportResponse(LedgerStateResponse):
    model_config = ConfigDict(populate_by_name=True)

    exported_at: Optional[str] = Field(default=None, alias="exportedAt")


class LedgerJobPreviewTransaction(BaseModel):
    transaction_date: Optional[str]
    description: Optional[str]
    withdrawal_amount: int = 0
    deposit_amount: int = 0
    memo: Optional[str] = None


class LedgerJobPreviewAccount(BaseModel):
    asset_id: str = Field(alias="assetId")
    account_name: Optional[str] = Field(alias="accountName", default=None)
    account_number: Optional[str] = Field(alias="accountNumber", default=None)
    owner_name: List[str] = Field(default_factory=list, alias="ownerName")
    transaction_count: int = Field(alias="transactionCount")
    total_withdrawal: int = Field(alias="totalWithdrawal")
    total_deposit: int = Field(alias="totalDeposit")
    sample_transactions: List[LedgerJobPreviewTransaction] = Field(default_factory=list, alias="sampleTransactions")


class LedgerJobPreviewResponse(BaseModel):
    status: Literal["ok"]
    job_id: str = Field(alias="jobId")
    accounts: List[LedgerJobPreviewAccount]


class LedgerJobImportMapping(BaseModel):
    asset_id: str = Field(alias="assetId")
    mode: Literal["new", "merge"]
    target_account_id: Optional[str] = Field(default=None, alias="targetAccountId")
    account_name: Optional[str] = Field(default=None, alias="accountName")
    account_number: Optional[str] = Field(default=None, alias="accountNumber")
    holder_name: Optional[str] = Field(default=None, alias="holderName")
    group_key: Optional[str] = Field(default=None, alias="groupKey")
    group_name: Optional[str] = Field(default=None, alias="groupName")
    group_number: Optional[str] = Field(default=None, alias="groupNumber")
    group_holder_name: Optional[str] = Field(default=None, alias="groupHolderName")


class LedgerJobImportRequest(BaseModel):
    case_id: Optional[str] = Field(default=None, alias="caseId")
    new_case_name: Optional[str] = Field(default=None, alias="newCaseName")
    mappings: List[LedgerJobImportMapping]
