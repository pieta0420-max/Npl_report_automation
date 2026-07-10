"""Canonical target schema for the 4 sheets we carry into the 투심 report.

Extracted by hand from WRB 2026-2 Program 투심보고서_VF.xlsx (Sheet A / Sheet
C-1 / Sheet E / Sheet F), which is the company's proven output format. Column
order matches the source workbook so the row layout looks the same to a
reviewer who is used to that file.

Two columns in the collateral (C-1) sheet -- `opb` and `claim_amount` -- do
NOT exist anywhere in the raw Data Disk. Cell-formula inspection of VF.xlsx
showed they are computed by allocating each borrower's Sheet-A-level OPB /
claim total across that borrower's properties, weighted by each property's
share of the borrower's total appraisal amount:

    property.opb          = borrower.opb_incl_prepaid * property.appraisal_amount_total
                             / sum(appraisal_amount_total for that borrower's properties)
    property.claim_amount = borrower.claim_total        * property.appraisal_amount_total
                             / sum(appraisal_amount_total for that borrower's properties)

normalizer.py implements this allocation; column_matcher.py must never try to
match a source DD column onto these two keys.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass(frozen=True)
class ColumnDef:
    key: str
    header_kr: str
    header_en: str
    derived: bool = False  # computed by normalizer, never matched to a source column
    category: str = ""  # top-level merged group header shown above the column, e.g.
                         # "Borrower Identification Information" (cosmetic, mirrors VF layout)


@dataclass(frozen=True)
class SheetSchema:
    type_id: str
    label_kr: str
    label_en: str
    title_keywords: List[str]
    columns: List[ColumnDef] = field(default_factory=list)

    def header_text(self) -> str:
        return f"{self.label_kr} ({self.label_en})"


BORROWER = SheetSchema(
    type_id="borrower",
    label_kr="차주일반정보",
    label_en="Borrower Information",
    title_keywords=["차주일반정보", "Borrower Information", "차주정보"],
    columns=[
        ColumnDef("serial_no", "일련번호", "Serial Number", category="Borrower Identification Information"),
        ColumnDef("pool_type", "Pool 구분", "Pool Type", category="Borrower Identification Information"),
        ColumnDef("asset_type", "자산유형", "Asset Type", category="Borrower Identification Information"),
        ColumnDef("borrower_control_no", "차주관리번호", "Borrower Control Number", category="Borrower Identification Information"),
        ColumnDef("borrower_name", "차주명", "Borrower Name", category="Borrower Identification Information"),
        ColumnDef("borrower_type", "차주형태", "Type of Borrower", category="Borrower Identification Information"),
        ColumnDef("related_borrower", "관련차주", "Related Borrower", category="Borrower Identification Information"),
        ColumnDef("opb_excl_prepaid", "환산된 미상환원금잔액(가지급금 제외)",
                   "Coverted Outstanding Principal Balance not including prepaid expense",
                   category="Summary of the Loans to be Sold"),
        ColumnDef("prepaid_expense", "가지급금 잔액", "Prepaid expense", category="Summary of the Loans to be Sold"),
        ColumnDef("opb_incl_prepaid", "환산된 미상환원금잔액(가지급금 포함)",
                   "Coverted Outstanding Principal Balance including prepaid expense",
                   category="Summary of the Loans to be Sold"),
        ColumnDef("accrued_interest", "미수이자 잔액", "Accrued Interest", category="Summary of the Loans to be Sold"),
        ColumnDef("claim_total", "채권액 합계", "Sum of OPB and Accrued Interest", category="Summary of the Loans to be Sold"),
        ColumnDef("mortgage_amount_converted", "차주별 환산 후 근저당권설정액",
                   "Converted Property Mortgage Amount", category="Mortgage Amount"),
        ColumnDef("senior_mortgage_amount", "차주별 선순위 근저당권 설정액",
                   "Converted Senior Property Claims- Senior Mortgage Amount", category="Mortgage Amount"),
        ColumnDef("pre_cutoff_court_deposit", "자산확정일 이전 공탁금", "Pre-Cut-off Date Court Deposit",
                   category="Description"),
        ColumnDef("remarks", "비고", "Deficiency Report", category="Description"),
    ],
)

COLLATERAL = SheetSchema(
    type_id="collateral",
    label_kr="담보재산정보_물건별",
    label_en="Property Information",
    title_keywords=["담보재산정보", "Property Information", "담보정보", "물건별"],
    columns=[
        ColumnDef("serial_no", "일련번호", "Serial Number", category="Borrower Identification Information"),
        ColumnDef("pool_type", "Pool 구분", "Pool Type", category="Borrower Identification Information"),
        ColumnDef("asset_type", "자산유형", "Asset Type", category="Borrower Identification Information"),
        ColumnDef("borrower_control_no", "차주관리번호", "Borrower Control Number", category="Borrower Identification Information"),
        ColumnDef("borrower_name", "차주명", "Borrower Name", category="Borrower Identification Information"),
        ColumnDef("property_serial_no", "Property 일련번호", "Property Serial Number", category="Property Basic Information"),
        ColumnDef("property_addr1", "담보소재지1(특별,광역시/도)", "Property Address 1 (Province)", category="Property Basic Information"),
        ColumnDef("property_type", "Property 담보물형태", "Property Type", category="Property Basic Information"),
        ColumnDef("appraisal_date", "감정평가일자", "Appraisal Date", category="Appraisal Information"),
        ColumnDef("appraisal_amount_total", "감정평가액 합계", "Appraisal Amount- Total", category="Appraisal Information"),
        ColumnDef("opb", "OPB", "OPB", derived=True, category="Claim Amount (안분계산)"),
        ColumnDef("claim_amount", "채권액", "Claim Amount", derived=True, category="Claim Amount (안분계산)"),
        ColumnDef("foreclosure_status", "경매개시 FiledNot Filed", "Foreclosure - Status", category="Foreclosure Information"),
    ],
)

GUARANTEE = SheetSchema(
    type_id="guarantee",
    label_kr="신용보증서 정보",
    label_en="Letter of Credit Guarantee Information",
    title_keywords=["신용보증서", "Letter of Credit Guarantee", "보증서"],
    columns=[
        ColumnDef("serial_no", "일련번호", "Serial Number", category="Borrower Identification Information"),
        ColumnDef("pool_type", "Pool 구분", "Pool Type", category="Borrower Identification Information"),
        ColumnDef("asset_type", "자산유형", "Asset Type", category="Borrower Identification Information"),
        ColumnDef("borrower_control_no", "차주관리번호", "Borrower Control Number", category="Borrower Identification Information"),
        ColumnDef("borrower_name", "차주명", "Borrower Name", category="Borrower Identification Information"),
        ColumnDef("loan_control_no", "관련 채권 일련번호", "Loan Control Number", category="Letter of Credit Guarantee Information"),
        ColumnDef("loan_account_no", "관련 대출채권 계좌번호", "Loan Account Number", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantor", "보증기관", "Guarantor", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantee_no", "보증서번호", "Credit Guarantee No", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantee_due_date", "보증기한", "Date of Credit Guarantee", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantee_ratio", "보증비율", "Guarantee Ratio", category="Letter of Credit Guarantee Information"),
        ColumnDef("initial_guarantee_amount", "보증금액", "Initial Guarantee Amount", category="Letter of Credit Guarantee Information"),
        ColumnDef("currency", "통화", "Original Currency", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantee_balance", "보증잔액", "Balance of Credit Guarantee", category="Letter of Credit Guarantee Information"),
        ColumnDef("guarantee_balance_converted", "환산후보증잔액", "Converted Balance of Credit Guarantee", category="Letter of Credit Guarantee Information"),
        ColumnDef("registration_date", "보증사고통지일", "Registration Date", category="Letter of Credit Guarantee Information"),
        ColumnDef("claim_date", "이행청구일", "Claim Date of Unpaid Letter of Credit Guarantee", category="Letter of Credit Guarantee Information"),
        ColumnDef("subrogation_date", "대위변제일", "Date of Subrogation", category="Letter of Credit Guarantee Information"),
        ColumnDef("remarks", "비고", "Deficiency Report", category="Description"),
    ],
)

REHAB = SheetSchema(
    type_id="rehab",
    label_kr="회생차주 일반정보",
    label_en="Borrower Information",
    title_keywords=["회생차주", "회생절차", "Rehabilitation"],
    columns=[
        ColumnDef("serial_no", "일련번호", "Serial Number", category="Borrower Identification Information"),
        ColumnDef("pool_type", "Pool 구분", "Pool Type", category="Borrower Identification Information"),
        ColumnDef("asset_type", "자산유형", "Asset Type", category="Borrower Identification Information"),
        ColumnDef("borrower_control_no", "차주관리번호", "Borrower Control Number", category="Borrower Identification Information"),
        ColumnDef("borrower_name", "차주명", "Borrower Name", category="Borrower Identification Information"),
        ColumnDef("approved", "인가/미인가", "Approved / Not Approved", category="General Understanding of the Subject Company"),
        ColumnDef("state", "세부 진행단계", "State", category="General Understanding of the Subject Company"),
        ColumnDef("court", "관할법원", "Court in Jurisdiction", category="General Understanding of the Subject Company"),
        ColumnDef("case_no", "회생사건번호", "Case Number", category="General Understanding of the Subject Company"),
        ColumnDef("case_borrower_name", "회생절차상 채무자 성명", "Borrower Name On Case", category="General Understanding of the Subject Company"),
        ColumnDef("claim_class", "회생담보권/회생채권/공익채권", "Secured/Unsecured Claims/General Claims", category="General Understanding of the Subject Company"),
        ColumnDef("hq_address", "본점 주소", "Location of Headquarter", category="General Understanding of the Subject Company"),
        ColumnDef("listed", "상장여부", "Listed / Unlisted", category="General Understanding of the Subject Company"),
        ColumnDef("industry_class", "표준산업분류", "Main Business", category="General Understanding of the Subject Company"),
        ColumnDef("major_products", "주요상품", "Major Products", category="General Understanding of the Subject Company"),
        ColumnDef("ceo_name", "대표이사", "Name of CEO", category="General Understanding of the Subject Company"),
        ColumnDef("lead_bank", "주거래은행", "Lead Bank", category="General Understanding of the Subject Company"),
        ColumnDef("employee_count", "종업원수", "Number of Employees", category="General Understanding of the Subject Company"),
        ColumnDef("established_date", "회사설립일", "Date of Company Establishment", category="General Understanding of the Subject Company"),
    ],
)

CANONICAL_SCHEMAS = {
    "borrower": BORROWER,
    "collateral": COLLATERAL,
    "guarantee": GUARANTEE,
    "rehab": REHAB,
}

# Sheet type -> output tab name used by writer.py
OUTPUT_SHEET_NAMES = {
    "borrower": "Sheet A (차주)",
    "collateral": "Sheet C-1 (담보)",
    "guarantee": "Sheet E (신용보증서)",
    "rehab": "Sheet F (회생)",
    "summary": "투심",
}
