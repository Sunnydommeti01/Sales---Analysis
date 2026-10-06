# ============================================================
# RENGY — CP + VENDOR LEAD-LEVEL MASTER DATABASE
# JUPYTER / PYTHON VERSION
#
# FINAL OUTPUT COLUMNS:
#   Lead Number
#   Type
#   Vendor / CP Name
#   Customer Name
#   NBFC
#   Customer Region
#   Vendor Region
#   Lead Created At
#   Disbursed At
#   Stage
#   Project Value
#   Dynamic Pricing
#   Disbursed Value
#   Lead Sub Stage
#
# OUTPUT:
# C:\Users\venka\OneDrive\Desktop\CP_Vendor_Lead_Master.xlsx
# ============================================================

import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from getpass import getpass

import numpy as np
import pandas as pd
import requests


# ============================================================
# 1. CONFIG
# ============================================================

BASE_URL = "https://apiportal.rengy.in/api"

LOGIN_URL = f"{BASE_URL}/auth/login"
USERS_URL = f"{BASE_URL}/users"
LEAD_URL = f"{BASE_URL}/lead/"
LOAN_URL = f"{BASE_URL}/loan"
PAYMENT_URL = f"{BASE_URL}/paymentHistory"

OUTPUT_FILE = (
    r"C:\Users\venka\OneDrive\Desktop"
    r"\CP_Vendor_Lead_Master.xlsx"
)

REQUEST_TIMEOUT = 90
API_PAGE_LIMIT = 2000
MAX_API_PAGES = 500

RENGY_USER_TYPE = "rengyStaff"

# Secure runtime login.
RENGY_IDENTIFIER = ""

RENGY_PASSWORD = ""

_AUTH_LOCK = threading.RLock()

_AUTH_STATE = {
    "access_token": "",
    "refresh_token": "",
}


# ============================================================
# 2. BASIC HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value).strip()


def to_number(value):
    if value is None:
        return 0.0

    if isinstance(
        value,
        (int, float, np.integer, np.floating)
    ):
        try:
            if pd.isna(value):
                return 0.0
        except Exception:
            pass

        return float(value)

    text = str(value).strip()

    if not text:
        return 0.0

    text = re.sub(
        r"[₹,\s]",
        "",
        text
    )

    try:
        return float(text)
    except Exception:
        return 0.0


def get_nested_value(
    obj,
    path
):
    current = obj

    for part in path.split("."):

        if not isinstance(
            current,
            dict
        ):
            return None

        if part not in current:
            return None

        current = current.get(part)

    return current


def get_value(
    obj,
    *paths,
    default=None
):
    for path in paths:

        value = get_nested_value(
            obj,
            path
        )

        if value is None:
            continue

        if isinstance(
            value,
            str
        ):
            if value.strip():
                return value
            continue

        return value

    return default


def normalize_region(value):
    raw = clean_text(value)

    return (
        raw
        if raw
        else "Unknown"
    )


def normalize_key(value):
    return re.sub(
        r"[^a-z0-9]+",
        "",
        clean_text(value).casefold()
    )


def normalize_datetime(value):
    dt = pd.to_datetime(
        value,
        errors="coerce",
        utc=True
    )

    if pd.isna(dt):
        return pd.NaT

    try:
        return (
            dt
            .tz_convert("Asia/Kolkata")
            .tz_localize(None)
        )
    except Exception:
        try:
            return dt.tz_localize(None)
        except Exception:
            return dt


def _normalize_token(value):
    token = (
        str(value or "")
        .strip()
        .strip('"')
        .strip("'")
    )

    if token.lower().startswith(
        "bearer "
    ):
        token = token[7:].strip()

    return token


# ============================================================
# 3. AUTHENTICATION
# ============================================================

def rengy_login(force=False):

    with _AUTH_LOCK:

        if (
            _AUTH_STATE["access_token"]
            and not force
        ):
            return _AUTH_STATE[
                "access_token"
            ]

        response = requests.post(
            LOGIN_URL,
            json={
                "identifier":
                    RENGY_IDENTIFIER,
                "password":
                    RENGY_PASSWORD,
                "userType":
                    RENGY_USER_TYPE,
            },
            headers={
                "Accept":
                    "application/json, "
                    "text/plain, */*",
                "Content-Type":
                    "application/json",
                "Origin":
                    "https://portal.rengy.in",
                "Referer":
                    "https://portal.rengy.in/",
            },
            timeout=REQUEST_TIMEOUT,
        )

        try:
            payload = response.json()
        except Exception:
            payload = {}

        if not response.ok:

            message = ""

            if isinstance(
                payload,
                dict
            ):
                message = str(
                    payload.get("message")
                    or payload.get("error")
                    or payload.get("detail")
                    or ""
                )[:300]

            raise PermissionError(
                "Rengy automatic login failed. "
                f"HTTP {response.status_code}"
                + (
                    f" — {message}"
                    if message
                    else ""
                )
            )

        data = (
            payload.get("data", {})
            if isinstance(
                payload,
                dict
            )
            else {}
        )

        if not isinstance(
            data,
            dict
        ):
            data = {}

        access_token = _normalize_token(
            data.get("accessToken")
        )

        refresh_token = _normalize_token(
            data.get("refreshToken")
        )

        if not access_token:
            raise PermissionError(
                "Rengy login returned HTTP 200 "
                "but data.accessToken was missing."
            )

        _AUTH_STATE[
            "access_token"
        ] = access_token

        _AUTH_STATE[
            "refresh_token"
        ] = refresh_token

        return access_token


def current_auth_headers(
    force_login=False
):
    access_token = rengy_login(
        force=force_login
    )

    headers = {
        "Authorization":
            f"Bearer {access_token}",
        "Accept":
            "application/json, "
            "text/plain, */*",
        "Origin":
            "https://portal.rengy.in",
        "Referer":
            "https://portal.rengy.in/",
    }

    refresh_token = _AUTH_STATE.get(
        "refresh_token",
        ""
    )

    if refresh_token:
        headers[
            "x-refresh-token"
        ] = refresh_token

    return headers


def authenticated_get(
    url,
    params=None,
    timeout=REQUEST_TIMEOUT
):
    response = requests.get(
        url,
        params=params,
        headers=current_auth_headers(
            force_login=False
        ),
        timeout=timeout,
    )

    if response.status_code in (
        401,
        403
    ):
        response = requests.get(
            url,
            params=params,
            headers=current_auth_headers(
                force_login=True
            ),
            timeout=timeout,
        )

    return response


# ============================================================
# 4. API PAGINATION HELPERS
# ============================================================

def find_record_list(
    payload,
    expected_keys
):
    if isinstance(
        payload,
        list
    ):
        return payload

    if not isinstance(
        payload,
        dict
    ):
        return []

    for key in expected_keys:

        value = payload.get(key)

        if isinstance(
            value,
            list
        ):
            return value

        if isinstance(
            value,
            dict
        ):
            for nested_key in (
                "data",
                "results",
                "records",
                "items",
                "users",
                "leads",
                "loans",
                "payments",
            ):
                nested_value = value.get(
                    nested_key
                )

                if isinstance(
                    nested_value,
                    list
                ):
                    return nested_value

    data = payload.get("data")

    if isinstance(
        data,
        list
    ):
        return data

    if isinstance(
        data,
        dict
    ):
        for key in expected_keys:
            value = data.get(key)

            if isinstance(
                value,
                list
            ):
                return value

    return []


def find_total_count(payload):

    if not isinstance(
        payload,
        dict
    ):
        return None

    candidates = [
        payload.get("total"),
        payload.get("count"),
        payload.get("totalCount"),
    ]

    data = payload.get("data")

    if isinstance(
        data,
        dict
    ):
        candidates.extend(
            [
                data.get("total"),
                data.get("count"),
                data.get(
                    "totalCount"
                ),
            ]
        )

    for value in candidates:

        try:
            if value is not None:
                return int(value)
        except Exception:
            pass

    return None


def api_record_key(
    record,
    description
):
    if not isinstance(
        record,
        dict
    ):
        return None

    if "lead" in description.casefold():

        key = get_value(
            record,
            "leadCode",
            "leadId",
            "_id",
            "id"
        )

    else:
        key = get_value(
            record,
            "_id",
            "id",
            "userCode",
            "leadCode",
            "paymentHistoryRefNo"
        )

    if key is None:
        return None

    return str(key).strip()


def fetch_all_paginated(
    url,
    description,
    expected_keys,
    base_params=None,
    page_limit=API_PAGE_LIMIT,
):

    all_records = []
    seen_keys = set()

    skip = 0
    page_number = 1
    api_total = None

    while (
        page_number
        <= MAX_API_PAGES
    ):

        params = dict(
            base_params or {}
        )

        params["limit"] = page_limit
        params["skip"] = skip

        print(
            f"\r{description}: "
            f"page {page_number} | "
            f"loaded "
            f"{len(all_records):,}",
            end="",
        )

        response = authenticated_get(
            url,
            params=params,
            timeout=REQUEST_TIMEOUT,
        )

        if not response.ok:
            raise RuntimeError(
                f"{description} API failed. "
                f"HTTP "
                f"{response.status_code}: "
                f"{response.text[:500]}"
            )

        payload = response.json()

        records = find_record_list(
            payload,
            expected_keys
        )

        current_total = (
            find_total_count(
                payload
            )
        )

        if current_total is not None:
            api_total = current_total

        if not records:
            break

        new_records = 0

        for record in records:

            key = api_record_key(
                record,
                description
            )

            if key is not None:

                key = (
                    description
                    + "::"
                    + key
                )

                if key in seen_keys:
                    continue

                seen_keys.add(key)

            all_records.append(
                record
            )

            new_records += 1

        if (
            api_total is not None
            and len(all_records)
            >= api_total
        ):
            break

        if new_records == 0:
            break

        skip += len(records)
        page_number += 1

    print(
        f"\r{description}: "
        f"{len(all_records):,} "
        f"records loaded"
        + (
            f" / API total "
            f"{api_total:,}"
            if api_total
            else ""
        )
        + " " * 20
    )

    return all_records


# ============================================================
# 5. FETCH DATA
# ============================================================

def fetch_channel_partners():

    return fetch_all_paginated(
        url=USERS_URL,
        description=(
            "Channel Partners"
        ),
        expected_keys=[
            "users",
            "channelPartners",
            "data",
            "results",
            "records",
            "items",
        ],
        base_params={
            "userType":
                "channelPartner"
        },
    )


def fetch_vendors():

    return fetch_all_paginated(
        url=USERS_URL,
        description="Vendors",
        expected_keys=[
            "users",
            "vendors",
            "data",
            "results",
            "records",
            "items",
        ],
        base_params={
            "userType":
                "vendor"
        },
    )


def fetch_leads():

    # includeFirstTransactionAt=true:
    # source of truth for Disbursed At.
    return fetch_all_paginated(
        url=LEAD_URL,
        description="CRM Leads",
        expected_keys=[
            "data",
            "leads",
            "lead",
            "results",
            "records",
            "items",
        ],
        base_params={
            "userType":
                "rengyStaff",
            "isArchived":
                "false",
            "includeFirstTransactionAt":
                "true",
        },
        # IMPORTANT: app.py requests the lead master with limit=20000.
        # The Rengy lead endpoint currently does not reliably return the
        # remaining rows when Sales uses limit=2000 + skip=2000.
        # Using the same 20k request as app.py keeps both dashboards on
        # the exact same CRM lead population.
        page_limit=20000,
    )


def fetch_loans():

    return fetch_all_paginated(
        url=LOAN_URL,
        description="Loans",
        expected_keys=[
            "data",
            "loans",
            "results",
            "records",
            "items",
        ],
        base_params={
            "filterStatus":
                "loanStatus"
        },
        page_limit=20000,
    )


def fetch_payments():

    return fetch_all_paginated(
        url=PAYMENT_URL,
        description="Payments",
        expected_keys=[
            "data",
            "payments",
            "paymentHistory",
            "results",
            "records",
            "items",
        ],
        base_params={},
        page_limit=20000,
    )


# ============================================================
# 6. BUILD PARTNER MASTERS
# ============================================================

def normalize_onboarded_at(value):
    """Convert /users onboardedAt to timezone-naive Asia/Kolkata datetime."""
    if value is None or str(value).strip() in {"", "None", "nan", "NaT"}:
        return pd.NaT
    ts = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(ts):
        return pd.NaT
    try:
        return ts.tz_convert("Asia/Kolkata").tz_localize(None)
    except Exception:
        return pd.to_datetime(ts, errors="coerce")


def build_partner_master(
    raw_users,
    partner_type
):
    rows = []

    for user in raw_users:

        if not isinstance(
            user,
            dict
        ):
            continue

        internal_id = get_value(
            user,
            "_id",
            "id"
        )

        code = get_value(
            user,
            "userCode",
            "vendorCode",
            "vendorId",
            "code"
        )

        name = get_value(
            user,
            "fullName",
            "vendorName",
            "partnerName",
            "businessName",
            "name"
        )

        region = get_value(
            user,
            "region",
            "regionName",
            "address.region"
        )

        owner_name = get_value(
            user,
            "rengyStaffOwnerId.fullName",
            "ownerName",
            "owner.fullName",
            "owner.name",
            "assignedOwner.fullName",
            "assignedOwner.name",
        )

        onboarded_at = get_value(
            user,
            "onboardedAt"
        )

        updated_at = get_value(
            user,
            "updatedAt",
            "createdAt"
        )

        rows.append(
            {
                "Type":
                    partner_type,
                "_internalId":
                    clean_text(
                        internal_id
                    ),
                "Code":
                    clean_text(code),
                "Name":
                    clean_text(name),
                "Vendor Region":
                    normalize_region(
                        region
                    ),
                "Owner Name":
                    clean_text(owner_name)
                    or "Unassigned",
                "Onboarded At":
                    normalize_onboarded_at(
                        onboarded_at
                    ),
                "_updatedAt":
                    normalize_datetime(
                        updated_at
                    ),
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(
            columns=[
                "Type",
                "_internalId",
                "Code",
                "Name",
                "Vendor Region",
                "Owner Name",
                "Onboarded At",
                "_updatedAt",
            ]
        )

    # Keep latest master row.
    has_code = (
        df["Code"]
        .fillna("")
        .astype(str)
        .str.strip()
        .ne("")
    )

    with_code = (
        df.loc[has_code]
        .sort_values(
            "_updatedAt",
            ascending=False,
            na_position="last"
        )
        .drop_duplicates(
            "Code",
            keep="first"
        )
    )

    without_code = (
        df.loc[~has_code]
        .sort_values(
            "_updatedAt",
            ascending=False,
            na_position="last"
        )
        .drop_duplicates(
            "_internalId",
            keep="first"
        )
    )

    return pd.concat(
        [
            with_code,
            without_code
        ],
        ignore_index=True
    )


# ============================================================
# 7. PARTNER LOOKUPS
# ============================================================

def make_lookup(master):

    code_lookup = {}
    id_lookup = {}

    for row in master.to_dict(
        "records"
    ):

        code = clean_text(
            row.get(
                "Code",
                ""
            )
        )

        internal_id = clean_text(
            row.get(
                "_internalId",
                ""
            )
        )

        if code:
            code_lookup[
                code.casefold()
            ] = row

        if internal_id:
            id_lookup[
                internal_id.casefold()
            ] = row

    return (
        code_lookup,
        id_lookup
    )


def match_partner(
    lead,
    cp_code_lookup,
    cp_id_lookup,
    vendor_code_lookup,
    vendor_id_lookup,
):

    source_type = normalize_key(
        get_value(
            lead,
            "sourceType",
            default=""
        )
    )

    partner_code = get_value(
        lead,
        "partnerInfo.userCode",
        "partnerInfo.vendorCode",
        "partnerInfo.vendorId",
        "channelPartner.userCode",
        "channelPartnerCode",
        "vendorInfo.userCode"
    )

    partner_internal_id = get_value(
        lead,
        "partnerInfo._id",
        "partnerInfo.id",
        "channelPartner._id",
        "channelPartner.id",
        "vendorInfo._id"
    )

    assigned_vendor_id = get_value(
        lead,
        "assignedVendorId._id",
        "assignedVendorId",
        "vendorId._id",
        "vendorId"
    )

    code_key = (
        clean_text(
            partner_code
        ).casefold()
    )

    partner_id_key = (
        clean_text(
            partner_internal_id
        ).casefold()
    )

    assigned_key = (
        clean_text(
            assigned_vendor_id
        ).casefold()
    )

    # ----------------------------------------
    # Explicit sourceType first.
    # ----------------------------------------

    if source_type == "vendor":

        if code_key:
            match = (
                vendor_code_lookup
                .get(code_key)
            )

            if match is not None:
                return match

        if partner_id_key:
            match = (
                vendor_id_lookup
                .get(partner_id_key)
            )

            if match is not None:
                return match

        if assigned_key:
            match = (
                vendor_id_lookup
                .get(assigned_key)
            )

            if match is not None:
                return match

        return None

    if source_type in {
        "channelpartner",
        "partner",
        "cp",
    }:

        if code_key:
            match = (
                cp_code_lookup
                .get(code_key)
            )

            if match is not None:
                return match

        if partner_id_key:
            match = (
                cp_id_lookup
                .get(partner_id_key)
            )

            if match is not None:
                return match

        if assigned_key:
            match = (
                cp_id_lookup
                .get(assigned_key)
            )

            if match is not None:
                return match

        return None

    # ----------------------------------------
    # Fallback for older CRM records where
    # sourceType may be blank.
    # CP first because partnerInfo historically
    # represented channel partners.
    # ----------------------------------------

    if code_key:

        match = cp_code_lookup.get(
            code_key
        )

        if match is not None:
            return match

        match = vendor_code_lookup.get(
            code_key
        )

        if match is not None:
            return match

    if partner_id_key:

        match = cp_id_lookup.get(
            partner_id_key
        )

        if match is not None:
            return match

        match = vendor_id_lookup.get(
            partner_id_key
        )

        if match is not None:
            return match

    if assigned_key:

        match = vendor_id_lookup.get(
            assigned_key
        )

        if match is not None:
            return match

        match = cp_id_lookup.get(
            assigned_key
        )

        if match is not None:
            return match

    return None


# ============================================================
# 8. BUILD LOAN SUMMARY
# ============================================================

def build_loan_summary(loans_raw, leads_raw):
    """
    Build NBFC + loan disbursed lookup.

    The /loan API may identify a lead by:
      - leadDetails.leadCode (human-readable lead number), or
      - leadId (internal CRM/Mongo lead _id).

    This function supports both so NBFC names are not lost.
    """

    internal_to_code = {}

    for lead in leads_raw:
        if not isinstance(lead, dict):
            continue

        internal_id = clean_text(get_value(lead, "_id", "id"))
        lead_code = clean_text(
            get_value(lead, "leadCode", "leadId", "leadID", "code")
        )

        if internal_id and lead_code:
            internal_to_code[internal_id.casefold()] = lead_code

    rows = []

    for loan in loans_raw:
        if not isinstance(loan, dict):
            continue

        # Preferred: human-readable lead number already populated.
        lead_code = clean_text(
            get_value(
                loan,
                "leadDetails.leadCode",
                "lead.leadCode",
                "leadInfo.leadCode",
                "leadCode",
            )
        )

        # Fallback: resolve loan.leadId (internal id) back to leadCode.
        internal_lead_id = clean_text(
            get_value(
                loan,
                "leadId._id",
                "leadId",
                "leadDetails._id",
                "lead._id",
            )
        )

        if not lead_code and internal_lead_id:
            lead_code = internal_to_code.get(
                internal_lead_id.casefold(),
                ""
            )

        if not lead_code:
            continue

        provider = get_value(
            loan,
            "loanProviderName",
            "providerName",
            "bankName",
            "nbfcName",
            "npsBankInfo.bankName",
            "npsBankInfo.name",
        )

        total_disbursed = get_value(
            loan,
            "totalDisbursedAmount",
            "disbursedAmount",
            "loanDisbursedAmount",
        )

        updated_at = get_value(
            loan,
            "updatedAt",
            "createdAt",
        )

        rows.append(
            {
                "Lead Number": lead_code,
                "NBFC": clean_text(provider),
                "_Loan Disbursed": to_number(total_disbursed),
                "_loanUpdatedAt": normalize_datetime(updated_at),
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(
            columns=[
                "Lead Number",
                "NBFC",
                "_Loan Disbursed",
            ]
        )

    # Latest loan row per actual Lead Number.
    df = (
        df
        .sort_values(
            "_loanUpdatedAt",
            ascending=True,
            na_position="first",
        )
        .drop_duplicates(
            "Lead Number",
            keep="last",
        )
        .reset_index(drop=True)
    )

    df["NBFC"] = (
        df["NBFC"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    return df[
        [
            "Lead Number",
            "NBFC",
            "_Loan Disbursed",
        ]
    ]

# ============================================================
# 9. BUILD PAYMENT SUMMARY
# ============================================================

def build_payment_summary(
    payments_raw
):
    rows = []

    for payment in payments_raw:

        if not isinstance(
            payment,
            dict
        ):
            continue

        finance_status = normalize_key(
            get_value(
                payment,
                "financeApprovalStatus",
                default=""
            )
        )

        if finance_status == "rejected":
            continue

        lead_code = get_value(
            payment,
            "leadDetails.leadCode",
            "leadCode",
            "lead.leadCode",
            "leadInfo.leadCode"
        )

        if not lead_code:
            continue

        transaction_date = get_value(
            payment,
            "transactionDate",
            "createdAt",
            "paymentDate"
        )

        paid_amount = get_value(
            payment,
            "paidAmount",
            "amount",
            "paymentAmount"
        )

        project_value = get_value(
            payment,
            "siteProjectValue",
            "projectValue"
        )

        dynamic_price = get_value(
            payment,
            "dynamicPriceAmount",
            "dynamicPriceValue",
            "dynamicPrice"
        )

        rows.append(
            {
                "Lead Number":
                    clean_text(
                        lead_code
                    ),
                "_transactionDate":
                    normalize_datetime(
                        transaction_date
                    ),
                "_paidAmount":
                    to_number(
                        paid_amount
                    ),
                "_paymentProjectValue":
                    to_number(
                        project_value
                    ),
                "_paymentDynamicPrice":
                    to_number(
                        dynamic_price
                    ),
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(
            columns=[
                "Lead Number",
                "_Payment Total",
                "_Payment Project Value",
                "_Payment Dynamic Price",
                "_Operations Project",
                "_Project Start Date",
            ]
        )

    # --------------------------------------------------------
    # OPERATIONS PROJECT LOGIC — EXACT MATCH
    # A project exists only when Payment History has:
    #   1) a non-rejected finance approval (handled above),
    #   2) a valid lead code, and
    #   3) a valid transaction date.
    # --------------------------------------------------------
    valid_project_payments = df[
        (df["Lead Number"].fillna("").astype(str).str.strip() != "")
        & df["_transactionDate"].notna()
    ].copy()

    if valid_project_payments.empty:
        return pd.DataFrame(
            columns=[
                "Lead Number",
                "_Payment Total",
                "_Payment Project Value",
                "_Payment Dynamic Price",
                "_Operations Project",
                "_Project Start Date",
            ]
        )

    # From here onward use ONLY the same valid payment rows that Operations uses.
    df = valid_project_payments.sort_values(
        ["Lead Number", "_transactionDate"]
    ).copy()

    project_flags = (
        df.groupby("Lead Number", as_index=False)["_transactionDate"]
        .first()
        .rename(columns={"_transactionDate": "_Project Start Date"})
    )
    project_flags["_Operations Project"] = True

    # Total received / disbursed from payment history.
    total = (
        df.groupby(
            "Lead Number",
            as_index=False
        )["_paidAmount"]
        .sum()
        .rename(
            columns={
                "_paidAmount":
                    "_Payment Total"
            }
        )
    )

    # Latest non-zero pricing values.
    df = df.sort_values(
        "_transactionDate",
        ascending=True,
        na_position="first"
    )

    pricing_rows = []

    for lead_code, group in df.groupby(
        "Lead Number",
        sort=False
    ):

        project_values = group.loc[
            group[
                "_paymentProjectValue"
            ] > 0,
            "_paymentProjectValue"
        ]

        dynamic_values = group.loc[
            group[
                "_paymentDynamicPrice"
            ] > 0,
            "_paymentDynamicPrice"
        ]

        pricing_rows.append(
            {
                "Lead Number":
                    lead_code,
                "_Payment Project Value":
                    (
                        float(
                            project_values.iloc[-1]
                        )
                        if len(
                            project_values
                        )
                        else 0.0
                    ),
                "_Payment Dynamic Price":
                    (
                        float(
                            dynamic_values.iloc[-1]
                        )
                        if len(
                            dynamic_values
                        )
                        else 0.0
                    ),
            }
        )

    pricing = pd.DataFrame(
        pricing_rows
    )

    return (
        total.merge(
            pricing,
            on="Lead Number",
            how="outer"
        )
        .merge(
            project_flags,
            on="Lead Number",
            how="left"
        )
    )


# ============================================================
# 10. BUILD CP + VENDOR LEAD MASTER
# ============================================================

def build_lead_master(
    leads_raw,
    cp_master,
    vendor_master,
    loan_summary,
    payment_summary,
):

    (
        cp_code_lookup,
        cp_id_lookup
    ) = make_lookup(
        cp_master
    )

    (
        vendor_code_lookup,
        vendor_id_lookup
    ) = make_lookup(
        vendor_master
    )

    rows = []

    unmatched = 0

    for lead in leads_raw:

        if not isinstance(
            lead,
            dict
        ):
            continue

        partner = match_partner(
            lead,
            cp_code_lookup,
            cp_id_lookup,
            vendor_code_lookup,
            vendor_id_lookup,
        )

        # IMPORTANT — keep the SAME lead population as app.py.
        # app.py does not discard a CRM lead just because CP/Vendor mapping is missing.
        # Sales-specific partner fields are attached when available; otherwise the
        # lead remains in the funnel as Unmapped / Direct.
        if partner is None:
            unmatched += 1
            partner = {
                "Type": "Unmapped / Direct",
                "Name": "Unmapped / Direct",
                "Vendor Region": "Unknown",
                "Owner Name": "Unassigned",
            }

        lead_number = get_value(
            lead,
            "leadCode",
            "leadId",
            "leadID",
            "code",
            "_id"
        )

        if not lead_number:
            continue

        customer_name = get_value(
            lead,
            "fullName",
            "leadName",
            "customerName",
            "name",
            "leadDetails.fullName",
            "leadDetails.customerName",
            "leadDetails.name"
        )

        customer_region = get_value(
            lead,
            "region",
            "Region",
            "regionName",
            "leadDetails.region",
            "address.region"
        )

        created_at = get_value(
            lead,
            "createdAt",
            "created_at",
            "leadCreatedAt"
        )

        disbursed_at = get_value(
            lead,
            "firstTransactionAt",
            "disbursedAt",
            "projectStartDate"
        )

        stage = get_value(
            lead,
            "leadStage",
            "milestone",
            "stage",
            "currentStage"
        )

        lead_sub_stage = get_value(
            lead,
            "leadSubStage",
            "subStage",
            "currentSubStage"
        )

        project_value = get_value(
            lead,
            "projectValue",
            "siteProjectValue",
            "dynamicPrice.projectValue",
            "pricingQuotation.projectValue",
            "quotation.projectValue",
            "leadDetails.projectValue"
        )

        dynamic_pricing = get_value(
            lead,
            "dynamicPriceAmount",
            "dynamicPriceValue",
            "dynamicPrice.amount",
            "dynamicPrice.value",
            "dynamicPrice.total",
            "pricingQuotation.dynamicPriceAmount",
            "pricingQuotation.dynamicPriceValue",
            "quotation.dynamicPriceAmount",
            "leadDetails.dynamicPriceAmount"
        )

        rows.append(
            {
                "Lead Number":
                    clean_text(
                        lead_number
                    ),
                "Type":
                    partner.get(
                        "Type",
                        ""
                    ),
                "Vendor / CP Name":
                    partner.get(
                        "Name",
                        ""
                    ),
                "Customer Name":
                    clean_text(
                        customer_name
                    ),
                "Customer Region":
                    normalize_region(
                        customer_region
                    ),
                "Vendor Region":
                    partner.get(
                        "Vendor Region",
                        "Unknown"
                    ),
                "Owner Name":
                    clean_text(
                        partner.get(
                            "Owner Name",
                            "Unassigned"
                        )
                    )
                    or "Unassigned",
                "Lead Created At":
                    normalize_datetime(
                        created_at
                    ),
                "Disbursed At":
                    normalize_datetime(
                        disbursed_at
                    ),
                "Stage":
                    clean_text(
                        stage
                    ),
                "Project Value":
                    to_number(
                        project_value
                    ),
                "Dynamic Pricing":
                    to_number(
                        dynamic_pricing
                    ),
                "Lead Sub Stage":
                    clean_text(
                        lead_sub_stage
                    ),
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return pd.DataFrame(
            columns=[
                "Lead Number",
                "Type",
                "Vendor / CP Name",
                "Customer Name",
                "NBFC",
                "Customer Region",
                "Vendor Region",
                "Owner Name",
                "Lead Created At",
                "Disbursed At",
                "Stage",
                "Project Value",
                "Dynamic Pricing",
                "Disbursed Value",
                "Lead Sub Stage",
            ]
        )

    # Keep the latest CRM row for each lead.
    df = (
        df
        .sort_values(
            "Lead Created At",
            ascending=True,
            na_position="first"
        )
        .drop_duplicates(
            "Lead Number",
            keep="last"
        )
        .reset_index(drop=True)
    )

    # ----------------------------------------
    # Merge loan details
    # ----------------------------------------

    df = df.merge(
        loan_summary,
        on="Lead Number",
        how="left"
    )

    # ----------------------------------------
    # Merge payment/pricing details — same leadCode key as app.py
    # ----------------------------------------
    # app.py constructs its project master by grouping valid Payment History
    # rows on leadCode. Normalize both sides before joining so whitespace,
    # capitalization and numeric Excel-style suffixes cannot silently
    # produce zero converted projects.
    def _ops_lead_key(values):
        return (values.fillna("").astype(str).str.strip()
                .str.replace(r"\.0$", "", regex=True).str.casefold())

    df["_ops_join_key"] = _ops_lead_key(df["Lead Number"])
    payment_summary = payment_summary.copy()
    payment_summary["_ops_join_key"] = _ops_lead_key(payment_summary["Lead Number"])
    payment_summary = (payment_summary.drop(columns=["Lead Number"])
                       .drop_duplicates("_ops_join_key", keep="first"))
    df = df.merge(payment_summary, on="_ops_join_key", how="left", validate="many_to_one")
    df = df.drop(columns=["_ops_join_key"])
    if not df["_Operations Project"].fillna(False).any():
        st.warning("No Sales leads matched the app.py Payment History project base. "
                   "Check Payment History leadCode values and current Sales filters; "
                   "no stage-based fallback has been applied.")

    for col in [
        "_Loan Disbursed",
        "_Payment Total",
        "_Payment Project Value",
        "_Payment Dynamic Price",
    ]:
        if col not in df.columns:
            df[col] = 0.0

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    # NBFC:
    # If no loan exists, treat as Cash.
    if "NBFC" not in df.columns:
        df["NBFC"] = "Cash"

    df["NBFC"] = (
        df["NBFC"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace(
            "",
            "Cash"
        )
    )

    # ----------------------------------------
    # Project Value reconciliation
    # ----------------------------------------
    # Keep both source values for auditability.
    #
    # Commercial/final Project Value priority:
    #   1. Latest non-zero Payment History project value
    #      (siteProjectValue / projectValue)
    #   2. Lead API project value
    #
    # This prevents an older lead-level projectValue from overriding
    # the latest transaction-linked commercial project value.

    df["_Lead API Project Value"] = pd.to_numeric(
        df["Project Value"],
        errors="coerce",
    ).fillna(0.0)

    df["_Payment History Project Value"] = pd.to_numeric(
        df["_Payment Project Value"],
        errors="coerce",
    ).fillna(0.0)

    df["Project Value"] = np.where(
        df["_Payment History Project Value"] > 0,
        df["_Payment History Project Value"],
        df["_Lead API Project Value"],
    )

    df["_Project Value Source"] = np.select(
        [
            df["_Payment History Project Value"] > 0,
            df["_Lead API Project Value"] > 0,
        ],
        [
            "Payment History",
            "Lead API",
        ],
        default="Missing",
    )

    # ----------------------------------------
    # Dynamic Pricing fallback:
    # Lead API first, Payment History second.
    # ----------------------------------------

    df["Dynamic Pricing"] = np.where(
        pd.to_numeric(
            df["Dynamic Pricing"],
            errors="coerce"
        ).fillna(0) > 0,
        pd.to_numeric(
            df["Dynamic Pricing"],
            errors="coerce"
        ).fillna(0),
        df[
            "_Payment Dynamic Price"
        ],
    )

    # ----------------------------------------
    # Disbursed Value:
    # Prefer loan.totalDisbursedAmount.
    # If no loan disbursement is available,
    # use total payment-history received.
    # ----------------------------------------

    df["Disbursed Value"] = np.where(
        df["_Loan Disbursed"] > 0,
        df["_Loan Disbursed"],
        df["_Payment Total"],
    )

    # Operations-compatible project flag. Missing payment match = not a project.
    if "_Operations Project" not in df.columns:
        df["_Operations Project"] = False
    df["_Operations Project"] = df["_Operations Project"].fillna(False).astype(bool)

    if "_Project Start Date" not in df.columns:
        df["_Project Start Date"] = pd.NaT
    df["_Project Start Date"] = pd.to_datetime(
        df["_Project Start Date"], errors="coerce"
    )

    # ----------------------------------------
    # Final column order
    # ----------------------------------------

    final_columns = [
        "Lead Number",
        "Type",
        "Vendor / CP Name",
        "Customer Name",
        "NBFC",
        "Customer Region",
        "Vendor Region",
        "Owner Name",
        "Lead Created At",
        "Disbursed At",
        "Stage",
        "Project Value",
        "Dynamic Pricing",
        "Disbursed Value",
        "Lead Sub Stage",
        "_Lead API Project Value",
        "_Payment History Project Value",
        "_Project Value Source",
        "_Operations Project",
        "_Project Start Date",
    ]

    # Lead Number is intentionally locked as Column A.
    df = df[
        final_columns
    ].copy()

    if df.columns[0] != "Lead Number":
        raise RuntimeError(
            "Lead Number must be the first output column."
        )

    # Latest leads first.
    df = df.sort_values(
        [
            "Lead Created At",
            "Lead Number"
        ],
        ascending=[
            False,
            True
        ],
        na_position="last"
    ).reset_index(
        drop=True
    )

    print(
        f"\nMatched CP/Vendor leads: "
        f"{len(df):,}"
    )

    print(
        f"Other/unmatched CRM records "
        f"ignored: {unmatched:,}"
    )

    return df


# ============================================================
# 11. EXPORT TO EXCEL
# ============================================================

def export_excel(
    df,
    output_file
):
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font,
        PatternFill,
        Border,
        Side,
        Alignment
    )
    from openpyxl.utils import (
        get_column_letter
    )
    from openpyxl.worksheet.table import (
        Table,
        TableStyleInfo
    )

    output_path = Path(
        output_file
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "CP Vendor Leads"

    # ----------------------------------------
    # Styles
    # ----------------------------------------

    header_fill = PatternFill(
        "solid",
        fgColor="1E3A8A"
    )

    white_font = Font(
        color="FFFFFF",
        bold=True
    )

    thin = Side(
        style="thin",
        color="D1D5DB"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # ----------------------------------------
    # Header
    # ----------------------------------------

    for col_idx, column in enumerate(
        df.columns,
        start=1
    ):
        cell = ws.cell(
            row=1,
            column=col_idx,
            value=column
        )

        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = border

    # ----------------------------------------
    # Data
    # ----------------------------------------

    for row_idx, record in enumerate(
        df.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):

        for col_idx, value in enumerate(
            record,
            start=1
        ):

            if pd.isna(value):
                value = None

            cell = ws.cell(
                row=row_idx,
                column=col_idx,
                value=value
            )

            cell.border = border
            cell.alignment = Alignment(
                vertical="center"
            )

    # ----------------------------------------
    # Formats
    # ----------------------------------------

    column_map = {
        name: idx + 1
        for idx, name
        in enumerate(df.columns)
    }

    for date_col in [
        "Lead Created At",
        "Disbursed At",
    ]:

        col_idx = column_map[
            date_col
        ]

        for row in range(
            2,
            ws.max_row + 1
        ):
            ws.cell(
                row,
                col_idx
            ).number_format = (
                "dd-mmm-yyyy hh:mm"
            )

    for money_col in [
        "Project Value",
        "Dynamic Pricing",
        "Disbursed Value",
    ]:

        col_idx = column_map[
            money_col
        ]

        for row in range(
            2,
            ws.max_row + 1
        ):
            ws.cell(
                row,
                col_idx
            ).number_format = (
                '₹#,##0.00'
            )

    # ----------------------------------------
    # Excel table / filters
    # ----------------------------------------

    if ws.max_row >= 2:

        table_ref = (
            f"A1:"
            f"{get_column_letter(ws.max_column)}"
            f"{ws.max_row}"
        )

        tab = Table(
            displayName=(
                "CPVendorLeadMaster"
            ),
            ref=table_ref
        )

        style = TableStyleInfo(
            name=(
                "TableStyleMedium2"
            ),
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        tab.tableStyleInfo = style
        ws.add_table(tab)

    # ----------------------------------------
    # Widths
    # ----------------------------------------

    widths = {
        "Lead Number": 18,
        "Type": 12,
        "Vendor / CP Name": 32,
        "Customer Name": 28,
        "NBFC": 22,
        "Customer Region": 20,
        "Vendor Region": 20,
        "Lead Created At": 21,
        "Disbursed At": 21,
        "Stage": 20,
        "Project Value": 18,
        "Dynamic Pricing": 18,
        "Disbursed Value": 18,
        "Lead Sub Stage": 24,
    }

    for column, width in (
        widths.items()
    ):

        idx = column_map[
            column
        ]

        ws.column_dimensions[
            get_column_letter(idx)
        ].width = width

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32

    wb.save(
        output_path
    )

    return output_path


# ============================================================
# RENGY SALES INTELLIGENCE — V2 STEP 1
# STANDALONE STREAMLIT APP
# ============================================================

import streamlit as st
import plotly.graph_objects as go

st.set_page_config(
    page_title="Rengy Sales Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PROFESSIONAL COLOR SYSTEM
# Navy / Slate / Emerald / Burnished Amber
# ============================================================
st.markdown(
    """
    <style>
    :root{
        --ink:#0F172A;
        --navy:#102A43;
        --slate:#475569;
        --muted:#64748B;
        --line:#E2E8F0;
        --surface:#FFFFFF;
        --surface-soft:#F8FAFC;
        --emerald:#0F766E;
        --emerald-dark:#0B5F58;
        --amber:#D97706;
        --amber-dark:#B45309;
        --amber-soft:#FFFBEB;
    }

    .stApp{
        background:
            linear-gradient(180deg,#F8FAFC 0%,#FFFFFF 22%,#F8FAFC 100%);
        color:var(--ink);
    }

    div[data-testid="stMetric"]{
        background:#FFFFFF !important;
        border:1px solid #E2E8F0 !important;
        border-top:3px solid #0F766E !important;
        box-shadow:0 5px 18px rgba(15,23,42,.045) !important;
    }

    div[data-testid="stMetricValue"]{
        color:#102A43 !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div{
        background:#FFFFFF !important;
        border-color:#CBD5E1 !important;
    }

    div[data-testid="stDataFrame"]{
        border:1px solid #E2E8F0 !important;
        box-shadow:0 4px 14px rgba(15,23,42,.035) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# UI STYLING
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at 82% 4%, rgba(14,165,233,.09), transparent 24%),
            radial-gradient(circle at 25% 0%, rgba(16,185,129,.055), transparent 20%),
            #f6f9fc;
    }

    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}

    header[data-testid="stHeader"] {
        background: rgba(246,249,252,.78);
        backdrop-filter: blur(12px);
    }

    .block-container {
        max-width: 1780px;
        padding-top: 1rem;
        padding-bottom: 3rem;
        padding-left: 1.45rem;
        padding-right: 1.45rem;
    }

    section[data-testid="stSidebar"] {
        background:
            radial-gradient(circle at 30% 4%, rgba(14,165,233,.22), transparent 23%),
            linear-gradient(180deg, #071c36 0%, #082a4b 48%, #073c55 100%);
        border-right: 1px solid rgba(255,255,255,.08);
    }

    section[data-testid="stSidebar"] * {
        color: #f8fafc;
    }

    section[data-testid="stSidebar"] label {
        color: #dbeafe !important;
        font-weight: 650 !important;
        font-size: .78rem !important;
    }

    .side-brand {
        padding: 1.05rem .1rem .65rem .1rem;
    }

    .side-kicker {
        color: #7dd3fc;
        font-size: .67rem;
        font-weight: 800;
        letter-spacing: .16em;
        text-transform: uppercase;
        margin-bottom: .35rem;
    }

    .side-title {
        color: #fff;
        font-size: 1.30rem;
        font-weight: 800;
        letter-spacing: -.035em;
        margin-bottom: .25rem;
    }

    .side-sub {
        color: #a9c7dc;
        font-size: .77rem;
        line-height: 1.45;
    }

    .live-filter {
        display: inline-flex;
        align-items: center;
        margin: .35rem 0 .9rem 0;
        padding: .35rem .62rem;
        border-radius: 999px;
        background: rgba(14,165,233,.13);
        border: 1px solid rgba(125,211,252,.18);
        color: #bae6fd;
        font-size: .68rem;
        font-weight: 750;
        letter-spacing: .04em;
    }

    .hero {
        position: relative;
        overflow: hidden;
        min-height: 158px;
        border-radius: 25px;
        padding: 1.5rem 1.7rem;
        margin-bottom: .9rem;
        color: white;
        background: linear-gradient(112deg, #0B1F33 0%, #123A56 58%, #0D5A68 100%);
        box-shadow: 0 14px 36px rgba(11,31,51,.14);
        border: 1px solid rgba(255,255,255,.08);
    }

    .hero:after {
        content:"";
        position:absolute;
        right:-65px;
        top:-125px;
        width:430px;
        height:430px;
        border-radius:50%;
        border:1px solid rgba(255,255,255,.10);
        box-shadow:
            0 0 0 45px rgba(255,255,255,.025),
            0 0 0 90px rgba(255,255,255,.018);
    }

    .hero-badge {
        display:inline-flex;
        padding:.38rem .65rem;
        border-radius:999px;
        background:rgba(255,255,255,.09);
        border:1px solid rgba(255,255,255,.12);
        color:#dff7ff;
        font-size:.69rem;
        font-weight:750;
        letter-spacing:.07em;
        margin-bottom:.8rem;
    }

    .hero h1 {
        position:relative;
        z-index:2;
        font-size:2.15rem;
        line-height:1.05;
        letter-spacing:-.055em;
        margin:0 0 .42rem 0;
        font-weight:700;white-space:nowrap;font-size:.55rem;
    }

    .hero p {
        position:relative;
        z-index:2;
        margin:0;
        color:#cde8f4;
        font-size:.90rem;
        max-width:780px;
    }

    .hero-meta {
        position:absolute;
        right:1.6rem;
        bottom:1.35rem;
        z-index:3;
        display:flex;
        gap:.5rem;
    }

    .meta-pill {
        padding:.46rem .7rem;
        border-radius:11px;
        background:rgba(255,255,255,.09);
        border:1px solid rgba(255,255,255,.11);
        color:#eaf8ff;
        font-size:.55rem;
        font-weight:650;
    }


    .kpi-badge {
        display:inline-flex;
        align-items:center;
        padding:.24rem .43rem;
        border-radius:999px;
        background:var(--soft);
        color:var(--accent);
        font-size:.57rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        white-space:nowrap;
        line-height:1;
    }

    .kpi-head-left {
        min-width:0;
    }

    .kpi-head-right {
        display:flex;
        align-items:center;
        gap:.25rem;
        flex-shrink:0;
    }

    .section-row {
        display:flex;
        align-items:end;
        justify-content:space-between;
        margin:.85rem 0 .48rem 0;
    }

    .section-kicker {
        color:#0284c7;
        font-size:.65rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        letter-spacing:.14em;
        text-transform:uppercase;
        margin-bottom:.17rem;
    }

    .section-title {
        color:#0f2740;
        font-size:1.17rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        letter-spacing:-.025em;
    }

    .section-note {
        color:#7890a5;
        font-size:.70rem;
    }

    .kpi-card {
        min-height:138px;
        position:relative;
        overflow:hidden;
        background:rgba(255,255,255,.97);
        border:1px solid #e4edf5;
        border-radius:18px;
        padding:.82rem .88rem .72rem .88rem;
        box-shadow:0 6px 20px rgba(30,64,175,.05);
        transition:transform .18s ease, box-shadow .18s ease;
    }

    .kpi-card:before {
        content:"";
        position:absolute;
        top:0;
        left:0;
        width:100%;
        height:4px;
        background:var(--accent);
    }

    .kpi-top {
        display:flex;
        justify-content:space-between;
        align-items:flex-start;
        margin-bottom:.38rem;
    }

    .kpi-icon {
        min-width:32px;
        height:32px;
        padding:0 .38rem;
        border-radius:10px;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:.78rem;
        background:var(--soft);
        color:var(--accent);
        font-weight:700;white-space:nowrap;font-size:.55rem;
        margin-left:.35rem;
    }


    .kpi-label {
        color:#6c8194;
        font-size:.55rem;
        font-weight:750;
        letter-spacing:.025em;
        text-transform:uppercase;
        margin-bottom:.12rem;
    }

    .kpi-value {
        color:#0b2239;
        font-size:1.52rem;
        line-height:1;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        letter-spacing:-.045em;
        margin-bottom:.45rem;
    }

    .kpi-bottom {
        border-top:1px solid #edf2f7;
        padding-top:.58rem;
    }

    .kpi-mini-label {
        color:#94a4b3;
        font-size:.53rem;
        font-weight:700;
        margin-bottom:.07rem;
        text-transform:uppercase;
        letter-spacing:.02em;
    }

    .kpi-mini-value {
        color:#28445e;
        font-size:.72rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        white-space:nowrap;
    }

    .kpi-bottom-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:.36rem;
        border-top:1px solid #edf2f7;
        padding-top:.48rem;
    }

    .kpi-mini-box + .kpi-mini-box {
        border-left:1px solid #edf2f7;
        padding-left:.48rem;
    }









    .next-space {
        margin-top:1rem;
        border:1.5px dashed #cad9e7;
        border-radius:20px;
        padding:1.05rem;
        text-align:center;
        color:#8297aa;
        background:rgba(255,255,255,.45);
        font-size:.74rem;
    }

    div[data-testid="stButton"] button {
        border-radius:12px;
        font-weight:750;
    }


    .chart-shell {
        background:rgba(255,255,255,.97);
        border:1px solid #e4edf5;
        border-radius:20px;
        padding:.95rem 1rem .35rem 1rem;
        margin-top:.78rem;
        box-shadow:0 7px 24px rgba(30,64,175,.045);
    }

    .chart-title-row {
        display:flex;
        justify-content:space-between;
        align-items:flex-end;
        margin:.95rem 0 .25rem 0;
    }

    .chart-kicker {
        color:#0284c7;
        font-size:.62rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        letter-spacing:.13em;
        text-transform:uppercase;
        margin-bottom:.12rem;
    }

    .chart-title {
        color:#0f2740;
        font-size:1.08rem;
        font-weight:700;white-space:nowrap;font-size:.55rem;
        letter-spacing:-.025em;
    }

    .chart-note {
        color:#64748B;
        font-size:.67rem;
    }

    @media (max-width:1000px) {
        .hero-meta {display:none;}
        
    }
    </style>
    """,
    unsafe_allow_html=True,
)



st.markdown(
    """
    <style>
    /* FINAL EXECUTIVE PALETTE
       Ink #0F172A | Navy #102A43 | Emerald #0F766E |
       Amber #D97706 | Slate #64748B | Surface #F8FAFC */
    .stApp {
        background:#F8FAFC !important;
    }

    .chart-kicker, .side-kicker {
        color:#0F766E !important;
    }

    .chart-title {
        color:#102A43 !important;
    }

    .chart-note {
        color:#64748B !important;
    }

    [class*="st-key-top_basic_filters"] {
        background:#FFFFFF !important;
        border:1px solid #E2E8F0 !important;
        box-shadow:0 4px 14px rgba(15,23,42,.04) !important;
    }

    div[data-testid="stMetric"] {
        background:#FFFFFF !important;
        border:1px solid #E2E8F0 !important;
        border-top:3px solid #0F766E !important;
        box-shadow:0 4px 14px rgba(15,23,42,.04) !important;
    }

    div[data-testid="stMetricValue"] {
        color:#102A43 !important;
    }

    div[data-testid="stDataFrame"] {
        background:#FFFFFF !important;
        border:1px solid #E2E8F0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# STREAMLIT LOGIN
# ============================================================
# Same automatic-login pattern as the proven reference dashboard.
# Streamlit Secrets:
#
#   RENGY_IDENTIFIER = "your_login_identifier"
#   RENGY_PASSWORD   = "your_login_password"
#
# No credential input boxes are rendered in the dashboard.
# ============================================================

def _runtime_secret(name, default=""):
    """Read deployment credentials without rendering login fields."""
    try:
        value = st.secrets.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    # Optional local/server fallback.
    value = os.getenv(name, default)
    return str(value or "").strip()


RENGY_IDENTIFIER = _runtime_secret("RENGY_IDENTIFIER")
RENGY_PASSWORD = _runtime_secret("RENGY_PASSWORD")

RENGY_LOGIN_URL = "https://apiportal.rengy.in/api/auth/login"
RENGY_USER_TYPE = "rengyStaff"

if not RENGY_IDENTIFIER or not RENGY_PASSWORD:
    st.error("CRM configuration is unavailable for this deployment.")
    st.caption(
        "Configure RENGY_IDENTIFIER and RENGY_PASSWORD once in the app's "
        "Streamlit Secrets, then reboot the app."
    )
    st.stop()


with st.sidebar:
    st.markdown(
        """
        <div class="side-brand">
            <div class="side-kicker">Rengy Intelligence</div>
            <div class="side-title">Sales Command Center</div>
            <div class="side-sub">Live CP, Vendor and sales portfolio intelligence.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


_AUTH_LOCK = threading.RLock()
_AUTH_STATE = {
    "access_token": "",
    "refresh_token": "",
}


def _normalize_token(value):
    token = str(value or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    return token


def rengy_login(force=False):
    """
    Authenticate with the confirmed Rengy login API and keep the returned
    access/refresh tokens only in app memory.

    The password is read only from Streamlit Secrets.
    Tokens are never written to GitHub or a local file.
    """
    with _AUTH_LOCK:
        if _AUTH_STATE["access_token"] and not force:
            return _AUTH_STATE["access_token"]

        response = requests.post(
            RENGY_LOGIN_URL,
            json={
                "identifier": RENGY_IDENTIFIER,
                "password": RENGY_PASSWORD,
                "userType": RENGY_USER_TYPE,
            },
            headers={
                "Accept": "application/json, text/plain, */*",
                "Content-Type": "application/json",
                "Origin": "https://portal.rengy.in",
                "Referer": "https://portal.rengy.in/",
            },
            timeout=REQUEST_TIMEOUT,
        )

        try:
            payload = response.json()
        except Exception:
            payload = {}

        if not response.ok:
            message = ""
            if isinstance(payload, dict):
                message = str(
                    payload.get("message")
                    or payload.get("error")
                    or payload.get("detail")
                    or ""
                )[:300]

            raise PermissionError(
                "Rengy automatic login failed. "
                f"HTTP {response.status_code}"
                + (f" — {message}" if message else "")
            )

        data = payload.get("data", {}) if isinstance(payload, dict) else {}
        if not isinstance(data, dict):
            data = {}

        access_token = _normalize_token(data.get("accessToken"))
        refresh_token = _normalize_token(data.get("refreshToken"))

        if not access_token:
            raise PermissionError(
                "Rengy login returned HTTP 200 but data.accessToken was missing."
            )

        _AUTH_STATE["access_token"] = access_token
        _AUTH_STATE["refresh_token"] = refresh_token
        return access_token


def current_auth_headers(force_login=False):
    token = rengy_login(force=force_login)

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://portal.rengy.in",
        "Referer": "https://portal.rengy.in/",
    }

    refresh_token = _AUTH_STATE.get("refresh_token", "")
    if refresh_token:
        headers["x-refresh-token"] = refresh_token

    return headers


# Compatibility headers. api_get() rebuilds auth headers per request.
HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://portal.rengy.in",
    "Referer": "https://portal.rengy.in/",
}


# ============================================================
# LIVE DATA LOAD
# ============================================================

@st.cache_data(ttl=3600, show_spinner=False)
def load_dashboard_data(identifier, password):
    global RENGY_IDENTIFIER
    global RENGY_PASSWORD
    global _AUTH_STATE

    RENGY_IDENTIFIER = identifier
    RENGY_PASSWORD = password

    _AUTH_STATE["access_token"] = ""
    _AUTH_STATE["refresh_token"] = ""

    # One login, then fetch independent CRM resources concurrently.
    # This is materially faster than waiting for CP -> Vendor -> Leads
    # -> Loans -> Payments one after another.
    rengy_login(force=True)

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_cp = executor.submit(fetch_channel_partners)
        future_vendor = executor.submit(fetch_vendors)
        future_leads = executor.submit(fetch_leads)
        future_loans = executor.submit(fetch_loans)
        future_payments = executor.submit(fetch_payments)

        cp_raw = future_cp.result()
        vendor_raw = future_vendor.result()
        leads_raw = future_leads.result()
        loans_raw = future_loans.result()
        payments_raw = future_payments.result()

    cp_master = build_partner_master(
        cp_raw,
        "CP"
    )

    vendor_master = build_partner_master(
        vendor_raw,
        "Vendor"
    )

    loan_summary = build_loan_summary(
        loans_raw,
        leads_raw
    )

    payment_summary = build_payment_summary(
        payments_raw
    )

    final = build_lead_master(
        leads_raw=leads_raw,
        cp_master=cp_master,
        vendor_master=vendor_master,
        loan_summary=loan_summary,
        payment_summary=payment_summary,
    )

    return (
        final,
        cp_master,
        vendor_master,
    )


with st.spinner("Syncing live CRM data..."):
    try:
        (
            df,
            df_cp_master,
            df_vendor_master,
        ) = load_dashboard_data(
            app_identifier,
            app_password,
        )
    except Exception as exc:
        st.error(
            f"CRM sync failed: {exc}"
        )
        st.stop()


# ============================================================
# PREPARE DATA
# ============================================================

def normalized_stage(series):
    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .str.replace(
            " ",
            "",
            regex=False
        )
        .str.replace(
            "_",
            "",
            regex=False
        )
        .str.replace(
            "-",
            "",
            regex=False
        )
    )


# EXACT funnel stage logic from app.py
PENDING_FUNNEL_STAGES = {
    "payment",
    "pricingquotation",
    "leaddetails",
    "sitesurvey",
}

DISBURSED_FUNNEL_STAGES = {
    "dispatch",
    "installation",
    "netmetering",
    "handover",
    "dpr",
}


df = df.copy()

df["Lead Created At"] = pd.to_datetime(
    df["Lead Created At"],
    errors="coerce",
)

df["Disbursed At"] = pd.to_datetime(
    df["Disbursed At"],
    errors="coerce",
)

# Precompute normalized date/period columns ONCE.
# All dashboard filters reuse these vectors instead of repeatedly calling
# .dt.date / .dt.to_period on every Streamlit rerun.
df["_lead_date"] = df["Lead Created At"].dt.normalize()
df["_disbursed_date"] = df["Disbursed At"].dt.normalize()
df["_lead_month"] = df["Lead Created At"].dt.to_period("M")
df["_disbursed_month"] = df["Disbursed At"].dt.to_period("M")

df["Project Value"] = pd.to_numeric(
    df["Project Value"],
    errors="coerce",
).fillna(0)

df["Disbursed Value"] = pd.to_numeric(
    df["Disbursed Value"],
    errors="coerce",
).fillna(0)

df["_stage_key"] = normalized_stage(
    df["Stage"]
)

# EXACT app.py funnel classification.
# A lead is counted as Disbursed/Converted when its CURRENT CRM stage is one of
# app.py's DISBURSED_FUNNEL_STAGES. Payment History is NOT used for this KPI.
df["_is_project"] = df["_stage_key"].isin(DISBURSED_FUNNEL_STAGES)
df["_is_pending"] = df["_stage_key"].isin(PENDING_FUNNEL_STAGES)


# ============================================================
# FILTERS
# ============================================================

with st.sidebar:

    st.markdown(
        '<div class="live-filter">● LIVE CRM FILTERS</div>',
        unsafe_allow_html=True,
    )

    valid_dates = (
        df["Lead Created At"]
        .dropna()
    )

    if not valid_dates.empty:
        min_date = (
            valid_dates.min().date()
        )
        max_date = (
            valid_dates.max().date()
        )

        date_range = st.date_input(
            "Lead Created Date",
            value=(
                min_date,
                max_date
            ),
            min_value=min_date,
            max_value=max_date,
        )
    else:
        date_range = None

    type_options = sorted(
        [
            x
            for x in
            df["Type"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_types = st.multiselect(
        "Partner Type",
        type_options,
        default=type_options,
    )

    partner_options = sorted(
        [
            x
            for x in
            df["Vendor / CP Name"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_partners = st.multiselect(
        "CP / Vendor",
        partner_options,
        placeholder="All CPs & Vendors",
    )

    customer_regions = sorted(
        [
            x
            for x in
            df["Customer Region"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_customer_regions = st.multiselect(
        "Customer Region",
        customer_regions,
        placeholder="All regions",
    )

    vendor_regions = sorted(
        [
            x
            for x in
            df["Vendor Region"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_vendor_regions = st.multiselect(
        "Vendor Region",
        vendor_regions,
        placeholder="All vendor regions",
    )

    consultant_options = sorted(
        [
            x
            for x in
            df["Owner Name"]
            .fillna("Unassigned")
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_consultants = st.multiselect(
        "Consultant / Owner",
        consultant_options,
        placeholder="All consultants",
    )

    nbfc_options = sorted(
        [
            x
            for x in
            df["NBFC"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_nbfcs = st.multiselect(
        "NBFC",
        nbfc_options,
        placeholder="All NBFCs",
    )

    stage_options = sorted(
        [
            x
            for x in
            df["Stage"]
            .dropna()
            .astype(str)
            .unique()
            if x.strip()
        ]
    )

    selected_stages = st.multiselect(
        "Stage",
        stage_options,
        placeholder="All stages",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button(
        "↻ Refresh Live Data",
        use_container_width=True,
    ):
        st.cache_data.clear()
        st.rerun()


filtered = df.copy()

# Period-role flags keep the business logic correct:
#   Leads    = Lead Created At falls in the selected period
#   Projects = Disbursed At falls in the selected period
# The working frame is the UNION of both, so an older lead disbursed this
# month is retained as a project without incorrectly becoming a new lead.
filtered["_period_lead"] = True
filtered["_period_project"] = True

if (
    date_range
    and isinstance(date_range, (tuple, list))
    and len(date_range) == 2
):
    start_date, end_date = date_range
    full_start_date = min_date if not valid_dates.empty else None
    full_end_date = max_date if not valid_dates.empty else None

    if (
        full_start_date is None
        or full_end_date is None
        or start_date != full_start_date
        or end_date != full_end_date
    ):
        _start_ts = pd.Timestamp(start_date)
        _end_ts = pd.Timestamp(end_date)
        _lead_mask = filtered["_lead_date"].between(_start_ts, _end_ts)
        _project_mask = filtered["_disbursed_date"].between(_start_ts, _end_ts)
        filtered["_period_lead"] = _lead_mask
        filtered["_period_project"] = _project_mask
        filtered = filtered.loc[_lead_mask | _project_mask]

if selected_types:
    filtered = filtered.loc[
        filtered["Type"]
        .isin(selected_types)
    ]

if selected_partners:
    filtered = filtered.loc[
        filtered["Vendor / CP Name"]
        .isin(selected_partners)
    ]

if selected_customer_regions:
    filtered = filtered.loc[
        filtered["Customer Region"]
        .isin(
            selected_customer_regions
        )
    ]

if selected_vendor_regions:
    filtered = filtered.loc[
        filtered["Vendor Region"]
        .isin(
            selected_vendor_regions
        )
    ]

if selected_consultants:
    filtered = filtered.loc[
        filtered["Owner Name"]
        .isin(selected_consultants)
    ]

if selected_nbfcs:
    filtered = filtered.loc[
        filtered["NBFC"]
        .isin(selected_nbfcs)
    ]

if selected_stages:
    filtered = filtered.loc[
        filtered["Stage"]
        .isin(selected_stages)
    ]



# ============================================================
# TOP BASIC FILTER BAR
# ============================================================
# Fast executive filters above the KPI cards.
# These further refine the existing sidebar-filtered dataset.
# Month takes precedence over the custom date range when selected.
# ============================================================

# Header timestamp must be defined BEFORE the hero renders.
latest_sync = (
    pd.Timestamp.now()
    .strftime(
        "%d %b %Y • %I:%M %p"
    )
)

st.markdown(
    f"""
    <div class="hero">
        <div class="hero-badge">
            ⚡ LIVE • SALES INTELLIGENCE
        </div>
        <h1>Rengy Sales Command Center</h1>
        <p>
            From lead generation to project conversion —
            one focused view of CP, Vendor and portfolio performance.
        </p>
        <div class="hero-meta">
            <div class="meta-pill">
                ● CRM Connected
            </div>
            <div class="meta-pill">
                Updated {latest_sync}
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


st.markdown(
    """
    <style>
    /* ==========================================================
       RENGY PREMIUM VISUAL SYSTEM — FAST CSS-ONLY GRADING
       ========================================================== */
    :root {
        --rengy-navy:#0b2538;
        --rengy-blue:#2f6fed;
        --rengy-amber:#B45309;
        --rengy-teal:#0F766E;
        --rengy-emerald:#0F766E;
        --rengy-ink:#0F172A;
        --rengy-muted:#7b91a2;
        --rengy-line:rgba(23,56,79,.08);
    }

    div[data-testid="stMetric"] {
        background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);
        border:1px solid var(--rengy-line);
        border-radius:15px;
        padding:.72rem .82rem;
        box-shadow:0 8px 24px rgba(11,37,56,.045);
    }

    div[data-testid="stMetric"] label {
        color:#71899a !important;
        font-weight:750 !important;
    }

    div[data-testid="stMetricValue"] {
        color:#102f46 !important;
        font-weight:900 !important;
        letter-spacing:-.025em;
    }

    div[data-testid="stDataFrame"] {
        border:1px solid rgba(23,56,79,.07);
        border-radius:13px;
        overflow:hidden;
        box-shadow:0 5px 18px rgba(11,37,56,.035);
    }

    div[data-testid="stPlotlyChart"] {
        border-radius:14px;
        overflow:hidden;
    }

    button[kind="secondary"] {
        border-radius:10px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    [class*="st-key-top_basic_filters"] {
        border:1px solid rgba(15,39,64,.075);
        border-radius:16px;
        padding:.48rem .62rem .30rem .62rem;
        margin:.58rem 0 .58rem 0;
        background:linear-gradient(180deg,#FFFFFF 0%,#FFFCF5 100%);
        border-left:4px solid #0F766E;
        box-shadow:0 7px 20px rgba(15,23,42,.05);
    }
    [class*="st-key-top_basic_filters"] div[data-testid="stVerticalBlock"] {
        gap:.18rem;
    }
    [class*="st-key-top_basic_filters"] label {
        font-size:.62rem !important;
        font-weight:800 !important;
        color:#6f8494 !important;
        letter-spacing:.025em;
    }
    [class*="st-key-top_basic_filters"] div[data-baseweb="select"] > div,
    [class*="st-key-top_basic_filters"] div[data-baseweb="input"] > div {
        min-height:2.05rem !important;
        border-radius:9px !important;
        font-size:.73rem !important;
    }
    .top-filter-kicker {
        color:#0F766E;
        font-size:.58rem;
        font-weight:900;
        letter-spacing:.10em;
        text-transform:uppercase;
        margin-bottom:.08rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.container(key="top_basic_filters"):
    st.markdown(
        '<div class="top-filter-kicker">Quick Filters</div>',
        unsafe_allow_html=True,
    )

    top_month_options = ["All Months"]
    top_month_map = {}

    # Month selector covers both acquisition and conversion activity.
    month_periods = pd.Index(
        pd.concat(
            [filtered["_lead_month"], filtered["_disbursed_month"]],
            ignore_index=True,
        ).dropna().unique()
    ).sort_values(ascending=False)

    for period in month_periods:
        label = period.strftime("%b %Y")
        top_month_options.append(label)
        top_month_map[label] = period

    top_regions = sorted(
        [
            x for x in filtered["Customer Region"]
            .dropna().astype(str).unique()
            if x.strip()
        ]
    )

    top_types = sorted(
        [
            x for x in filtered["Type"]
            .dropna().astype(str).unique()
            if x.strip()
        ]
    )

    top_consultants = sorted(
        [
            x for x in filtered["Owner Name"]
            .fillna("Unassigned").astype(str).unique()
            if x.strip()
        ]
    )


    q1, q2, q3, q4, q5 = st.columns(
        [0.14, 0.24, 0.18, 0.20, 0.24],
        gap="small",
    )

    with q1:
        top_month = st.selectbox(
            "Month",
            top_month_options,
            key="top_filter_month",
        )

    with q2:
        _all_activity_dates = pd.concat(
            [filtered["_lead_date"], filtered["_disbursed_date"]],
            ignore_index=True,
        ).dropna()
        if not _all_activity_dates.empty:
            top_min_date = _all_activity_dates.min().date()
            top_max_date = _all_activity_dates.max().date()
            top_date_range = st.date_input(
                "Date Range",
                value=(top_min_date, top_max_date),
                min_value=top_min_date,
                max_value=top_max_date,
                key="top_filter_date_range",
            )
        else:
            top_date_range = None
            top_min_date = None
            top_max_date = None

    with q3:
        top_region = st.selectbox(
            "Region",
            ["All Regions"] + top_regions,
            key="top_filter_region",
        )

    with q4:
        top_type = st.selectbox(
            "Partner Type",
            ["All Types"] + top_types,
            key="top_filter_type",
        )

    with q5:
        top_consultant = st.selectbox(
            "Consultant",
            ["All Consultants"] + top_consultants,
            key="top_filter_consultant",
        )


    # Month/date semantics:
    # - Lead count uses Lead Created At in the selected period.
    # - Project/disbursal count uses Disbursed At in the selected period.
    # - The frame keeps the UNION so previous-month leads converted now remain visible.
    if top_month != "All Months":
        selected_period = top_month_map.get(top_month)
        if selected_period is not None:
            _lead_mask = filtered["_lead_month"].eq(selected_period)
            _project_mask = filtered["_disbursed_month"].eq(selected_period)
            filtered["_period_lead"] = filtered["_period_lead"] & _lead_mask
            filtered["_period_project"] = filtered["_period_project"] & _project_mask
            filtered = filtered.loc[
                filtered["_period_lead"] | filtered["_period_project"]
            ]
    elif (
        top_date_range
        and isinstance(top_date_range, (tuple, list))
        and len(top_date_range) == 2
        and top_min_date is not None
        and top_max_date is not None
    ):
        top_start, top_end = top_date_range
        if top_start != top_min_date or top_end != top_max_date:
            _start_ts = pd.Timestamp(top_start)
            _end_ts = pd.Timestamp(top_end)
            _lead_mask = filtered["_lead_date"].between(_start_ts, _end_ts)
            _project_mask = filtered["_disbursed_date"].between(_start_ts, _end_ts)
            filtered["_period_lead"] = filtered["_period_lead"] & _lead_mask
            filtered["_period_project"] = filtered["_period_project"] & _project_mask
            filtered = filtered.loc[
                filtered["_period_lead"] | filtered["_period_project"]
            ]

    if top_region != "All Regions":
        filtered = filtered.loc[
            filtered["Customer Region"].eq(top_region)
        ].copy()

    if top_type != "All Types":
        filtered = filtered.loc[
            filtered["Type"].eq(top_type)
        ].copy()

    if top_consultant != "All Consultants":
        filtered = filtered.loc[
            filtered["Owner Name"].eq(top_consultant)
        ].copy()


# ============================================================
# KPI CALCULATIONS
# ============================================================

lead_period_df = filtered.loc[filtered["_period_lead"]].copy()

project_df = filtered.loc[
    filtered["_is_project"] & filtered["_period_project"]
].copy()

pending_df = filtered.loc[
    filtered["_is_pending"] & filtered["_period_lead"]
].copy()

total_leads = int(lead_period_df["Lead Number"].nunique())

converted_projects = int(
    project_df[
        "Lead Number"
    ].nunique()
)

pending_leads = int(
    pending_df[
        "Lead Number"
    ].nunique()
)

project_value = float(
    project_df[
        "Project Value"
    ].sum()
)

# Project Value audit — useful when reconciling against the original
# revenue/master data. One row per unique project is already enforced
# upstream by the lead master.
project_value_audit = pd.DataFrame()

if not project_df.empty:
    audit_cols = [
        c for c in [
            "Lead Number",
            "Customer Name",
            "Customer Region",
            "Stage",
            "Project Value",
            "_Lead API Project Value",
            "_Payment History Project Value",
            "_Project Value Source",
        ]
        if c in project_df.columns
    ]

    project_value_audit = (
        project_df[audit_cols]
        .drop_duplicates(
            subset=["Lead Number"],
            keep="first",
        )
        .copy()
    )

    for c in [
        "Project Value",
        "_Lead API Project Value",
        "_Payment History Project Value",
    ]:
        if c in project_value_audit.columns:
            project_value_audit[c] = pd.to_numeric(
                project_value_audit[c],
                errors="coerce",
            ).fillna(0.0)


total_portfolio_value = float(
    filtered[
        "Project Value"
    ].sum()
)

pending_value = float(
    pending_df[
        "Project Value"
    ].sum()
)

disbursed_value = float(
    project_df["Disbursed Value"].sum()
)

# ------------------------------------------------------------
# PROJECT VALUE RECONCILIATION
# ------------------------------------------------------------
# Kept collapsed so it does not clutter the dashboard.
with st.sidebar.expander(
    "Project Value Check",
    expanded=False,
):
    st.caption(
        "Final Project Value now prefers the latest non-zero "
        "Payment History value; Lead API is the fallback."
    )

    # Do not call money_short() here because that helper is defined
    # later in the script. Format locally so this sidebar audit can
    # render safely during top-to-bottom Streamlit execution.
    if abs(project_value) >= 10_000_000:
        project_value_display = f"₹ {project_value / 10_000_000:.2f} Cr"
    elif abs(project_value) >= 100_000:
        project_value_display = f"₹ {project_value / 100_000:.2f} L"
    elif abs(project_value) >= 1_000:
        project_value_display = f"₹ {project_value / 1_000:.1f} K"
    else:
        project_value_display = f"₹ {project_value:,.0f}"

    st.metric(
        "Filtered Project Value",
        project_value_display,
    )

    if not project_value_audit.empty:
        payment_source_count = int(
            project_value_audit["_Project Value Source"]
            .eq("Payment History")
            .sum()
        )
        lead_source_count = int(
            project_value_audit["_Project Value Source"]
            .eq("Lead API")
            .sum()
        )

        a1, a2 = st.columns(2)

        a1.metric(
            "Payment History",
            f"{payment_source_count:,}",
        )
        a2.metric(
            "Lead API fallback",
            f"{lead_source_count:,}",
        )

        audit_export = project_value_audit.copy()

        st.download_button(
            "Download Value Reconciliation",
            data=audit_export.to_csv(index=False).encode("utf-8-sig"),
            file_name="Project_Value_Reconciliation.csv",
            mime="text/csv",
            use_container_width=True,
            key="download_project_value_reconciliation",
        )

total_vendors = int(
    df_vendor_master.shape[0]
)

total_cps = int(
    df_cp_master.shape[0]
)

vendor_leads = int(
    filtered.loc[
        filtered["Type"].eq(
            "Vendor"
        ),
        "Lead Number",
    ].nunique()
)

cp_leads = int(
    filtered.loc[
        filtered["Type"].eq(
            "CP"
        ),
        "Lead Number",
    ].nunique()
)


vendor_project_count = int(
    filtered.loc[
        filtered["Type"].eq("Vendor")
        & filtered["_is_project"],
        "Lead Number",
    ].nunique()
)

cp_project_count = int(
    filtered.loc[
        filtered["Type"].eq("CP")
        & filtered["_is_project"],
        "Lead Number",
    ].nunique()
)

vendor_conversion_rate = (
    vendor_project_count / vendor_leads * 100
    if vendor_leads else 0.0
)

cp_conversion_rate = (
    cp_project_count / cp_leads * 100
    if cp_leads else 0.0
)

pending_rate = (
    pending_leads / total_leads * 100
    if total_leads else 0.0
)

conversion_rate = (
    converted_projects
    / total_leads
    * 100
    if total_leads
    else 0.0
)

avg_project_value = (
    project_value
    / converted_projects
    if converted_projects
    else 0.0
)


avg_pending_value = (
    pending_value
    / pending_leads
    if pending_leads
    else 0.0
)


def money_short(value):
    value = float(
        value or 0
    )

    if abs(value) >= 10_000_000:
        return (
            f"₹ {value / 10_000_000:.2f} Cr"
        )

    if abs(value) >= 100_000:
        return (
            f"₹ {value / 100_000:.2f} L"
        )

    if abs(value) >= 1_000:
        return (
            f"₹ {value / 1_000:.1f} K"
        )

    return f"₹ {value:,.0f}"


def kpi_card(
    icon,
    label,
    value,
    stat1_label,
    stat1_value,
    stat2_label,
    stat2_value,
    accent,
    soft,
    badge,
):
    return (
        f'<div class="kpi-card" style="--accent:{accent};--soft:{soft};">'
        f'<div class="kpi-top">'
        f'<div class="kpi-head-left">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'</div>'
        f'<div class="kpi-head-right">'
        f'<div class="kpi-badge">{badge}</div>'
        f'<div class="kpi-icon">{icon}</div>'
        f'</div>'
        f'</div>'
        f'<div class="kpi-bottom-grid">'
        f'<div class="kpi-mini-box">'
        f'<div class="kpi-mini-label">{stat1_label}</div>'
        f'<div class="kpi-mini-value">{stat1_value}</div>'
        f'</div>'
        f'<div class="kpi-mini-box">'
        f'<div class="kpi-mini-label">{stat2_label}</div>'
        f'<div class="kpi-mini-value">{stat2_value}</div>'
        f'</div>'
        f'</div>'
        f'</div>'
    )



# ============================================================
# HEADER
# ============================================================
# Hero is rendered above the Quick Filters bar.
# latest_sync is initialized before that hero block.



# ============================================================
# KPI CARDS
# ============================================================

st.markdown(
    """
    <div class="section-row">
        <div>
            <div class="section-kicker">
                Portfolio Pulse
            </div>
            <div class="section-title">
                Sales at a glance
            </div>
        </div>
        <div class="section-note">
            All numbers respond instantly to the filters
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3, c4, c5 = (
    st.columns(
        5,
        gap="small"
    )
)

with c1:
    st.markdown(
        kpi_card(
            "◎",
            "TOTAL LEADS",
            f"{total_leads:,}",
            "Project Value",
            money_short(total_portfolio_value),
            "Filtered Portfolio",
            f"{total_leads:,} Leads",
            "#0ea5e9",
            "#e0f2fe",
            "100% BASE",
        ),
        unsafe_allow_html=True,
    )

with c2:
    st.markdown(
        kpi_card(
            "✓",
            "CONVERTED PROJECTS",
            f"{converted_projects:,}",
            "Project Value",
            money_short(project_value),
            "Avg Project",
            money_short(avg_project_value),
            "#10b981",
            "#dcfce7",
            f"{conversion_rate:.1f}%",
        ),
        unsafe_allow_html=True,
    )

with c3:
    st.markdown(
        kpi_card(
            "◷",
            "PENDING LEADS",
            f"{pending_leads:,}",
            "Pending Value",
            money_short(pending_value),
            "Avg Pending",
            money_short(avg_pending_value),
            "#f59e0b",
            "#fef3c7",
            f"{pending_rate:.1f}%",
        ),
        unsafe_allow_html=True,
    )

with c4:
    st.markdown(
        kpi_card(
            "V",
            "TOTAL VENDORS",
            f"{total_vendors:,}",
            "Vendor Leads",
            f"{vendor_leads:,}",
            "Converted",
            f"{vendor_project_count:,}",
            "#D97706",
            "#FFF7E6",
            f"{vendor_conversion_rate:.1f}% CVR",
        ),
        unsafe_allow_html=True,
    )

with c5:
    st.markdown(
        kpi_card(
            "CP",
            "TOTAL CPs",
            f"{total_cps:,}",
            "CP Leads",
            f"{cp_leads:,}",
            "Converted",
            f"{cp_project_count:,}",
            "#0F766E",
            "#ECFDF5",
            f"{cp_conversion_rate:.1f}% CVR",
        ),
        unsafe_allow_html=True,
    )


# ============================================================
# KPI CLICK-TO-AUDIT
# ============================================================
# Native Streamlit buttons are used deliberately:
# - reliable on Streamlit Cloud
# - no URL/query-string navigation
# - no additional API call
# - popups reuse the already loaded filtered dataframe
# ============================================================

st.markdown(
    """
    <style>
    [class*="st-key-kpi_audit_"] {
        margin-top:-.42rem !important;
        padding:0 !important;
    }

    [class*="st-key-kpi_audit_"] div[data-testid="stButton"] {
        margin:0 !important;
        padding:0 !important;
    }

    [class*="st-key-kpi_audit_"] button {
        width:100% !important;
        min-height:1.72rem !important;
        height:1.72rem !important;
        border-radius:0 0 12px 12px !important;
        border:1px solid #dbe5ec !important;
        border-top:0 !important;
        background:#f8fbfd !important;
        color:#31546a !important;
        box-shadow:none !important;
        font-size:.66rem !important;
        font-weight:800 !important;
        letter-spacing:.01em !important;
    }

    [class*="st-key-kpi_audit_"] button:hover {
        background:#eef6fa !important;
        color:#12344a !important;
        border-color:#c6d9e4 !important;
        transform:none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

kpi_click = None

audit_cols = st.columns(5, gap="small")

with audit_cols[0]:
    if st.button(
        "VIEW LEADS  ↗",
        key="kpi_audit_total_leads",
        use_container_width=True,
    ):
        kpi_click = "leads"

with audit_cols[1]:
    if st.button(
        "AUDIT PROJECTS  ↗",
        key="kpi_audit_projects",
        use_container_width=True,
    ):
        kpi_click = "projects"

with audit_cols[2]:
    if st.button(
        "VIEW PENDING  ↗",
        key="kpi_audit_pending",
        use_container_width=True,
    ):
        kpi_click = "pending"

with audit_cols[3]:
    if st.button(
        "VIEW VENDORS  ↗",
        key="kpi_audit_vendors",
        use_container_width=True,
    ):
        kpi_click = "vendors"

with audit_cols[4]:
    if st.button(
        "VIEW CPs  ↗",
        key="kpi_audit_cps",
        use_container_width=True,
    ):
        kpi_click = "cps"


def _kpi_detail_columns(frame):
    """Useful lead-level columns available in the current master."""
    return [
        c for c in [
            "Lead Number",
            "Customer Name",
            "Vendor / CP Name",
            "Type",
            "Customer Region",
            "Vendor Region",
            "Stage",
            "Lead Sub Stage",
            "NBFC",
            "Lead Created At",
            "Disbursed At",
            "Project Value",
            "Dynamic Pricing",
            "Disbursed Value",
            "_Lead API Project Value",
            "_Payment History Project Value",
            "_Project Value Source",
        ]
        if c in frame.columns
    ]


def _excel_filter_table(frame, key_prefix, height=430):
    """Fast Excel-like popup table: search + column filters + download-ready view."""
    if frame is None or frame.empty:
        st.info("No records are available for this selection.")
        return pd.DataFrame()

    work = frame.copy()
    controls = st.columns([0.34, 0.22, 0.22, 0.22], gap="small")

    with controls[0]:
        search_text = st.text_input(
            "Search",
            placeholder="Lead / customer / partner / consultant...",
            key=f"{key_prefix}_search",
        ).strip().casefold()

    filter_specs = [
        ("Customer Region", "Region"),
        ("Stage", "Stage"),
        ("NBFC", "NBFC"),
    ]

    for col_box, (col_name, label) in zip(controls[1:], filter_specs):
        with col_box:
            if col_name in work.columns:
                options = sorted(
                    x for x in work[col_name].fillna("Unknown").astype(str).unique()
                    if x.strip()
                )
                selected = st.selectbox(
                    label,
                    [f"All {label}s"] + options,
                    key=f"{key_prefix}_{col_name}",
                )
                if selected != f"All {label}s":
                    work = work.loc[
                        work[col_name].fillna("Unknown").astype(str).eq(selected)
                    ]
            else:
                st.caption(f"{label}: —")

    if search_text and not work.empty:
        search_cols = [
            c for c in [
                "Lead Number", "Customer Name", "Vendor / CP Name",
                "Owner Name", "Customer Region", "Vendor Region",
                "Stage", "Lead Sub Stage", "NBFC",
            ] if c in work.columns
        ]
        if search_cols:
            search_blob = (
                work[search_cols].fillna("").astype(str)
                .agg(" | ".join, axis=1).str.casefold()
            )
            work = work.loc[search_blob.str.contains(search_text, regex=False)]

    st.caption(f"{len(work):,} matching rows")
    st.dataframe(
        work,
        use_container_width=True,
        hide_index=True,
        height=min(height, 42 + max(1, len(work)) * 35),
        column_config={
            c: st.column_config.DatetimeColumn(c, format="DD MMM YYYY HH:mm")
            for c in ["Lead Created At", "Disbursed At"] if c in work.columns
        } | {
            c: st.column_config.NumberColumn(c, format="₹ %.0f")
            for c in [
                "Project Value", "Dynamic Pricing", "Disbursed Value",
                "_Lead API Project Value", "_Payment History Project Value",
            ] if c in work.columns
        },
    )
    return work


def _kpi_dataframe(frame, height=430):
    """Standard audit table formatting with Excel-like filters."""
    view_cols = _kpi_detail_columns(frame)

    if not view_cols:
        st.info("No detail columns are available.")
        return

    view = frame[view_cols].copy()

    if "Lead Number" in view.columns:
        view = view.drop_duplicates(
            subset=["Lead Number"],
            keep="first",
        )

    sort_col = (
        "Project Value"
        if "Project Value" in view.columns
        else (
            "Lead Created At"
            if "Lead Created At" in view.columns
            else None
        )
    )

    if sort_col:
        view = view.sort_values(
            sort_col,
            ascending=False,
            na_position="last",
        )

    config = {}

    for dt_col in [
        "Lead Created At",
        "Disbursed At",
    ]:
        if dt_col in view.columns:
            config[dt_col] = st.column_config.DatetimeColumn(
                dt_col,
                format="DD MMM YYYY HH:mm",
            )

    for money_col in [
        "Project Value",
        "Dynamic Pricing",
        "Disbursed Value",
        "_Lead API Project Value",
        "_Payment History Project Value",
    ]:
        if money_col in view.columns:
            config[money_col] = st.column_config.NumberColumn(
                money_col,
                format="₹ %.0f",
            )

    _excel_filter_table(
        view,
        key_prefix=f"kpi_{abs(hash(tuple(view.columns))) % 100000}",
        height=height,
    )


if kpi_click == "projects":

    @st.dialog(
        "Converted Projects — Full Audit",
        width="large",
    )
    def show_project_kpi_audit():
        st.caption(
            "Current filters • one row per Lead Number • "
            "project classification + Project Value source shown • "
            "no additional API call"
        )

        project_audit = project_df.copy()

        p1, p2, p3, p4 = st.columns(4)

        p1.metric(
            "Projects",
            f"{converted_projects:,}",
        )
        p2.metric(
            "Project Value",
            money_short(project_value),
        )
        p3.metric(
            "Avg Project",
            money_short(avg_project_value),
        )
        p4.metric(
            "Disbursed",
            money_short(
                pd.to_numeric(
                    project_audit["Disbursed Value"],
                    errors="coerce",
                ).fillna(0).sum()
            ),
        )

        st.markdown("#### Project Stage Check")

        if project_audit.empty:
            st.info("No converted projects under the current filters.")
        else:
            stage_check = (
                project_audit
                .assign(
                    _Stage_Display=(
                        project_audit["Stage"]
                        .fillna("Unknown")
                        .astype(str)
                        .str.strip()
                        .replace("", "Unknown")
                    ),
                    _PV=pd.to_numeric(
                        project_audit["Project Value"],
                        errors="coerce",
                    ).fillna(0.0),
                    _DV=pd.to_numeric(
                        project_audit["Disbursed Value"],
                        errors="coerce",
                    ).fillna(0.0),
                )
                .groupby(
                    "_Stage_Display",
                    observed=True,
                    dropna=False,
                )
                .agg(
                    Projects=("Lead Number", "nunique"),
                    Project_Value=("_PV", "sum"),
                    Disbursed=("_DV", "sum"),
                )
                .reset_index()
                .rename(
                    columns={
                        "_Stage_Display": "Stage",
                        "Project_Value": "Project Value",
                    }
                )
                .sort_values(
                    ["Projects", "Project Value"],
                    ascending=[False, False],
                )
                .reset_index(drop=True)
            )

            stage_check["Share %"] = (
                stage_check["Project Value"]
                .div(
                    project_value
                    if project_value
                    else 1
                )
                .mul(100)
                .round(1)
            )

            st.dataframe(
                stage_check,
                use_container_width=True,
                hide_index=True,
                height=min(
                    330,
                    42 + max(1, len(stage_check)) * 35,
                ),
                column_config={
                    "Project Value":
                        st.column_config.NumberColumn(
                            "Project Value",
                            format="₹ %.0f",
                        ),
                    "Disbursed":
                        st.column_config.NumberColumn(
                            "Disbursed",
                            format="₹ %.0f",
                        ),
                    "Share %":
                        st.column_config.NumberColumn(
                            "Share %",
                            format="%.1f%%",
                        ),
                },
            )

            # Explicit handover audit because this is the stage currently
            # being investigated.
            handover_mask = (
                project_audit["_stage_key"]
                .isin(
                    {
                        "handover",
                        "projecthandover",
                    }
                )
            )

            handover_projects = project_audit.loc[
                handover_mask
            ].copy()

            h_count = int(
                handover_projects["Lead Number"].nunique()
            )

            h_value = float(
                pd.to_numeric(
                    handover_projects["Project Value"],
                    errors="coerce",
                ).fillna(0).sum()
            )

            h1, h2 = st.columns(2)

            h1.metric(
                "Handover / Project Handover",
                f"{h_count:,} projects",
            )
            h2.metric(
                "Handover Project Value",
                money_short(h_value),
            )

            st.markdown(
                "#### Project Handover / Handover Records"
            )

            if handover_projects.empty:
                st.warning(
                    "No handover/projectHandover records are present "
                    "under the current dashboard filters."
                )
            else:
                _kpi_dataframe(
                    handover_projects,
                    height=320,
                )

            st.markdown("#### All Converted Project Records")
            _kpi_dataframe(
                project_audit,
                height=470,
            )

            project_export = project_audit[
                _kpi_detail_columns(project_audit)
            ].drop_duplicates(
                subset=["Lead Number"],
                keep="first",
            )

            st.download_button(
                "Download Converted Project Audit",
                data=project_export.to_csv(
                    index=False
                ).encode("utf-8-sig"),
                file_name="Converted_Project_Audit.csv",
                mime="text/csv",
                use_container_width=True,
                key="download_converted_project_audit",
            )

    show_project_kpi_audit()


elif kpi_click == "leads":

    @st.dialog(
        "Total Leads — Detail Audit",
        width="large",
    )
    def show_lead_kpi_audit():
        st.caption(
            "Current dashboard filters • unique Lead Number • "
            "no additional API call"
        )

        l1, l2, l3 = st.columns(3)
        l1.metric("Total Leads", f"{total_leads:,}")
        l2.metric(
            "Converted",
            f"{converted_projects:,}",
        )
        l3.metric(
            "Pending",
            f"{pending_leads:,}",
        )

        _kpi_dataframe(
            filtered,
            height=500,
        )

    show_lead_kpi_audit()


elif kpi_click == "pending":

    @st.dialog(
        "Pending Leads — Detail Audit",
        width="large",
    )
    def show_pending_kpi_audit():
        st.caption(
            "Leads whose normalized current Stage is not in "
            "app.py funnel stages • current filters applied"
        )

        q1, q2, q3 = st.columns(3)
        q1.metric(
            "Pending Leads",
            f"{pending_leads:,}",
        )
        q2.metric(
            "Pending Value",
            money_short(pending_value),
        )
        q3.metric(
            "Avg Pending Value",
            money_short(avg_pending_value),
        )

        if not pending_df.empty:
            pending_stage = (
                pending_df
                .assign(
                    _Stage_Display=(
                        pending_df["Stage"]
                        .fillna("Unknown")
                        .astype(str)
                        .str.strip()
                        .replace("", "Unknown")
                    )
                )
                .groupby(
                    "_Stage_Display",
                    observed=True,
                    dropna=False,
                )
                .agg(
                    Leads=("Lead Number", "nunique"),
                    Project_Value=(
                        "Project Value",
                        "sum",
                    ),
                )
                .reset_index()
                .rename(
                    columns={
                        "_Stage_Display": "Stage",
                        "Project_Value": "Lead Value",
                    }
                )
                .sort_values(
                    "Leads",
                    ascending=False,
                )
            )

            st.markdown("#### Pending Stage Breakdown")

            st.dataframe(
                pending_stage,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Lead Value":
                        st.column_config.NumberColumn(
                            "Lead Value",
                            format="₹ %.0f",
                        ),
                },
            )

        st.markdown("#### Individual Pending Leads")
        _kpi_dataframe(
            pending_df,
            height=480,
        )

    show_pending_kpi_audit()


elif kpi_click == "vendors":

    @st.dialog(
        "Vendor KPI — Detail Audit",
        width="large",
    )
    def show_vendor_kpi_audit():
        st.caption(
            "Registered vendor master + currently filtered vendor leads"
        )

        v1, v2, v3 = st.columns(3)
        v1.metric(
            "Registered Vendors",
            f"{total_vendors:,}",
        )
        v2.metric(
            "Vendor Leads",
            f"{vendor_leads:,}",
        )
        v3.metric(
            "Converted",
            f"{vendor_project_count:,}",
        )

        vendor_filtered = filtered.loc[
            filtered["Type"].eq("Vendor")
        ].copy()

        _kpi_dataframe(
            vendor_filtered,
            height=480,
        )

    show_vendor_kpi_audit()


elif kpi_click == "cps":

    @st.dialog(
        "CP KPI — Detail Audit",
        width="large",
    )
    def show_cp_kpi_audit():
        st.caption(
            "Registered CP master + currently filtered CP leads"
        )

        cp1, cp2, cp3 = st.columns(3)
        cp1.metric(
            "Registered CPs",
            f"{total_cps:,}",
        )
        cp2.metric(
            "CP Leads",
            f"{cp_leads:,}",
        )
        cp3.metric(
            "Converted",
            f"{cp_project_count:,}",
        )

        cp_filtered = filtered.loc[
            filtered["Type"].eq("CP")
        ].copy()

        _kpi_dataframe(
            cp_filtered,
            height=480,
        )

    show_cp_kpi_audit()



# ============================================================
# CP VS VENDOR — REAL PARTNER ONBOARDING
# SOURCE: /users -> onboardedAt
# Default = All Months (month-wise)
# Specific month = day-wise
# ============================================================

partner_onboarding = pd.concat(
    [df_cp_master.copy(), df_vendor_master.copy()],
    ignore_index=True,
)

if "Onboarded At" not in partner_onboarding.columns:
    partner_onboarding["Onboarded At"] = pd.NaT

partner_onboarding["Onboarded At"] = pd.to_datetime(
    partner_onboarding["Onboarded At"],
    errors="coerce",
)

partner_onboarding = partner_onboarding.loc[
    partner_onboarding["Onboarded At"].notna()
].copy()

# Only partner-level filters apply to onboarding.
# Lead Created Date / NBFC / Stage / Customer Region do NOT change onboarding counts.
if selected_types:
    partner_onboarding = partner_onboarding.loc[
        partner_onboarding["Type"].isin(selected_types)
    ]

if selected_partners:
    partner_onboarding = partner_onboarding.loc[
        partner_onboarding["Name"].isin(selected_partners)
    ]

if selected_vendor_regions:
    partner_onboarding = partner_onboarding.loc[
        partner_onboarding["Vendor Region"].isin(selected_vendor_regions)
    ]

partner_onboarding["_partner_key"] = (
    partner_onboarding["Type"].fillna("").astype(str)
    + "::"
    + partner_onboarding["_internalId"].fillna("").astype(str)
    + "::"
    + partner_onboarding["Code"].fillna("").astype(str)
    + "::"
    + partner_onboarding["Name"].fillna("").astype(str)
)

partner_onboarding = (
    partner_onboarding
    .sort_values("Onboarded At")
    .drop_duplicates("_partner_key", keep="first")
    .copy()
)

# Header and compact month selector on SAME ROW.
head_col, filter_col = st.columns([5.4, 1.35], gap="small")

with head_col:
    st.markdown(
        """
        <div style="padding-top:.08rem;">
            <div class="chart-kicker">Partner Network Growth</div>
            <div class="chart-title">CP vs Vendor Onboarding</div>
            <div class="chart-note" style="margin-top:.14rem;">
                /users onboardedAt • Hover for names • Click for performance
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if partner_onboarding.empty:
    with filter_col:
        st.selectbox(
            "Month",
            options=["All Months"],
            index=0,
            key="partner_onboarding_month_empty",
            label_visibility="collapsed",
        )

    st.info(
        "No CP/Vendor /users records currently contain a usable onboardedAt value."
    )

else:
    partner_onboarding["_onboard_month"] = (
        partner_onboarding["Onboarded At"].dt.to_period("M")
    )

    available_periods = sorted(
        partner_onboarding["_onboard_month"].dropna().unique(),
        reverse=True,
    )

    month_map = {
        period.strftime("%b %Y"): period
        for period in available_periods
    }

    month_options = ["All Months"] + list(month_map.keys())

    with filter_col:
        selected_month_label = st.selectbox(
            "Onboarding Month",
            options=month_options,
            index=0,
            key="partner_onboarding_month",
            label_visibility="collapsed",
            help=(
                "All Months shows month-wise onboarding. "
                "Choose a month for day-wise onboarding."
            ),
        )

    show_all_months = selected_month_label == "All Months"

    if show_all_months:
        chart_partners = partner_onboarding.copy()

        chart_partners["_bucket"] = (
            chart_partners["Onboarded At"]
            .dt.to_period("M")
            .dt.to_timestamp()
        )

        all_buckets = sorted(
            chart_partners["_bucket"].dropna().unique()
        )

        bucket_labels = [
            pd.Timestamp(x).strftime("%b %Y")
            for x in all_buckets
        ]

        x_tick_angle = 0
        x_tick_size = 10
        popup_date_mode = "month"

    else:
        selected_period = month_map[selected_month_label]

        chart_partners = partner_onboarding.loc[
            partner_onboarding["_onboard_month"] == selected_period
        ].copy()

        chart_partners["_bucket"] = (
            chart_partners["Onboarded At"].dt.normalize()
        )

        all_buckets = pd.date_range(
            selected_period.start_time,
            selected_period.end_time.normalize(),
            freq="D",
        )

        bucket_labels = [
            pd.Timestamp(x).strftime("%d %b")
            for x in all_buckets
        ]

        x_tick_angle = -45
        x_tick_size = 9
        popup_date_mode = "day"

    # FAST onboarding aggregation:
    # one groupby replaces repeated DataFrame scans for every bucket/type.
    compact_onboarding = chart_partners[
        ["_bucket", "Type", "Name"]
    ].copy()

    grouped_counts = (
        compact_onboarding
        .groupby(["_bucket", "Type"], observed=True, sort=False)
        .size()
        .rename("Count")
    )

    grouped_names = (
        compact_onboarding
        .dropna(subset=["Name"])
        .assign(Name=lambda d: d["Name"].astype(str).str.strip())
        .loc[lambda d: d["Name"].ne("")]
        .groupby(["_bucket", "Type"], observed=True, sort=False)["Name"]
        .agg(lambda s: sorted(set(s)))
    )

    summary_index = pd.MultiIndex.from_product(
        [pd.DatetimeIndex(all_buckets), ["CP", "Vendor"]],
        names=["_bucket", "Type"],
    )

    onboarding_summary = (
        grouped_counts
        .reindex(summary_index, fill_value=0)
        .reset_index()
    )

    onboarding_summary["NamesList"] = [
        grouped_names.get((bucket, ptype), [])
        for bucket, ptype in zip(
            onboarding_summary["_bucket"],
            onboarding_summary["Type"],
        )
    ]

    def _compact_hover(names):
        if not names:
            return "No onboarding"
        if len(names) <= 20:
            return "<br>".join(names)
        return "<br>".join(names[:20]) + f"<br>+ {len(names) - 20} more"

    onboarding_summary["Names"] = onboarding_summary["NamesList"].map(
        _compact_hover
    )

    if show_all_months:
        onboarding_summary["Label"] = (
            onboarding_summary["_bucket"].dt.strftime("%b %Y")
        )
        onboarding_summary["BucketID"] = (
            onboarding_summary["_bucket"].dt.strftime("%Y-%m")
        )
    else:
        onboarding_summary["Label"] = (
            onboarding_summary["_bucket"].dt.strftime("%d %b")
        )
        onboarding_summary["BucketID"] = (
            onboarding_summary["_bucket"].dt.strftime("%Y-%m-%d")
        )

    onboarding_summary["Count"] = onboarding_summary["Count"].astype("int32")


    fig = go.Figure()

    for ptype, color in [
        ("CP", "#0F766E"),
        ("Vendor", "#D97706"),
    ]:
        part = onboarding_summary.loc[
            onboarding_summary["Type"] == ptype
        ].copy()

        customdata = list(
            zip(
                part["Names"].tolist(),
                part["BucketID"].tolist(),
                [ptype] * len(part),
            )
        )

        fig.add_trace(
            go.Bar(
                name=ptype,
                x=part["Label"],
                y=part["Count"],
                marker=dict(
                    color=color,
                    line=dict(width=0),
                ),
                text=part["Count"].where(
                    part["Count"] > 0,
                    "",
                ),
                textposition="outside",
                textfont=dict(size=10),
                customdata=customdata,
                hovertemplate=(
                    "<b>%{x} • " + ptype + "</b>"
                    "<br>Onboarded: <b>%{y:,}</b>"
                    "<br><br><b>Names</b>"
                    "<br>%{customdata[0]}"
                    "<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        barmode="group",
        bargap=0.24,
        bargroupgap=0.07,
        height=365,
        margin=dict(
            l=12,
            r=12,
            t=18,
            b=24,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            family="Inter, Arial, sans-serif",
            color="#334155",
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            title=None,
            font=dict(size=11),
        ),
        hoverlabel=dict(
            bgcolor="#0f172a",
            font_size=12,
        ),
        xaxis=dict(
            title=None,
            categoryorder="array",
            categoryarray=bucket_labels,
            showgrid=False,
            showline=False,
            tickangle=x_tick_angle,
            tickfont=dict(
                size=x_tick_size,
                color="#64748b",
            ),
        ),
        yaxis=dict(
            title=dict(
                text="Partners Onboarded",
                font=dict(
                    size=10,
                    color="#94a3b8",
                ),
            ),
            rangemode="tozero",
            dtick=1,
            gridcolor="rgba(148,163,184,.16)",
            zeroline=False,
            showline=False,
            tickfont=dict(
                size=10,
                color="#94a3b8",
            ),
        ),
    )

    try:
        onboarding_event = st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "responsive": True,
                "staticPlot": False,
            },
            on_select="rerun",
            selection_mode="points",
            key="partner_onboarding_chart",
        )
    except TypeError:
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "responsive": True,
                "staticPlot": False,
            },
            key="partner_onboarding_chart_fallback",
        )
        onboarding_event = None

    selected_point = None

    if onboarding_event is not None:
        try:
            points = onboarding_event.selection.points
        except Exception:
            try:
                points = (
                    onboarding_event
                    .get("selection", {})
                    .get("points", [])
                )
            except Exception:
                points = []

        if points:
            selected_point = points[0]

    if selected_point:
        curve_number = selected_point.get(
            "curve_number",
            selected_point.get("curveNumber", None),
        )

        clicked_type = None

        if curve_number is not None:
            try:
                clicked_type = fig.data[int(curve_number)].name
            except Exception:
                clicked_type = None

        custom = selected_point.get("customdata")
        clicked_bucket_id = None

        if custom and len(custom) >= 3:
            clicked_bucket_id = custom[1]
            clicked_type = clicked_type or custom[2]

        if clicked_bucket_id and clicked_type in {"CP", "Vendor"}:
            if popup_date_mode == "month":
                clicked_period = pd.Period(
                    clicked_bucket_id,
                    freq="M",
                )

                clicked_partners = partner_onboarding.loc[
                    (partner_onboarding["Type"] == clicked_type)
                    & (
                        partner_onboarding["Onboarded At"]
                        .dt.to_period("M")
                        == clicked_period
                    )
                ].copy()

                clicked_label = clicked_period.strftime("%b %Y")

            else:
                clicked_day = pd.Timestamp(clicked_bucket_id)

                clicked_partners = partner_onboarding.loc[
                    (partner_onboarding["Type"] == clicked_type)
                    & (
                        partner_onboarding["Onboarded At"]
                        .dt.normalize()
                        == clicked_day.normalize()
                    )
                ].copy()

                clicked_label = clicked_day.strftime("%d %b %Y")

            clicked_names = set(
                clicked_partners["Name"]
                .dropna()
                .astype(str)
                .str.strip()
                .loc[lambda s: s.ne("")]
                .tolist()
            )

            # Performance lookup only.
            # The onboarding count itself comes ONLY from /users onboardedAt.
            perf = df.loc[
                (df["Type"] == clicked_type)
                & (
                    df["Vendor / CP Name"]
                    .astype(str)
                    .str.strip()
                    .isin(clicked_names)
                )
            ].copy()

            perf["_pv"] = pd.to_numeric(
                perf["Project Value"],
                errors="coerce",
            ).fillna(0)

            perf["_dv"] = pd.to_numeric(
                perf["Disbursed Value"],
                errors="coerce",
            ).fillna(0)

            name_col = (
                "Vendor Name"
                if clicked_type == "Vendor"
                else "CP Name"
            )

            review_rows = []

            for _, partner in clicked_partners.sort_values("Name").iterrows():
                pname = clean_text(partner.get("Name"))

                p = perf.loc[
                    perf["Vendor / CP Name"]
                    .astype(str)
                    .str.strip()
                    == pname
                ]

                leads_n = int(
                    p["Lead Number"].nunique()
                )

                converted_n = int(
                    p.loc[
                        p["_is_project"],
                        "Lead Number",
                    ].nunique()
                )

                cvr = (
                    converted_n / leads_n * 100
                    if leads_n
                    else 0.0
                )

                review_rows.append(
                    {
                        name_col: pname or "Unknown",
                        "Onboarded At":
                            partner.get("Onboarded At"),
                        "Leads":
                            leads_n,
                        "Converted":
                            converted_n,
                        "Conversion %":
                            round(cvr, 1),
                        "Lead Project Value":
                            float(p["_pv"].sum()),
                        "Converted Value":
                            float(
                                p.loc[
                                    p["_is_project"],
                                    "_pv",
                                ].sum()
                            ),
                        "Disbursed Value":
                            float(p["_dv"].sum()),
                    }
                )

            review_df = pd.DataFrame(review_rows)

            onboarded_n = int(
                len(clicked_partners)
            )

            leads_total = int(
                perf["Lead Number"].nunique()
            )

            converted_total = int(
                perf.loc[
                    perf["_is_project"],
                    "Lead Number",
                ].nunique()
            )

            lead_value_total = float(
                perf["_pv"].sum()
            )

            disbursed_total = float(
                perf["_dv"].sum()
            )

            dialog_title = (
                "Vendor Onboarding & Performance Review"
                if clicked_type == "Vendor"
                else "CP Onboarding & Performance Review"
            )

            @st.dialog(
                dialog_title,
                width="large",
            )
            def show_onboarding_popup():
                st.caption(
                    f"{clicked_label} • "
                    f"{clicked_type} • /users onboardedAt"
                )

                a, b, c, d, e = st.columns(5)

                a.metric(
                    "Onboarded",
                    f"{onboarded_n:,}",
                )

                b.metric(
                    "Leads",
                    f"{leads_total:,}",
                )

                c.metric(
                    "Converted",
                    f"{converted_total:,}",
                )

                d.metric(
                    "Lead Project Value",
                    money_short(lead_value_total),
                )

                e.metric(
                    "Disbursed Value",
                    money_short(disbursed_total),
                )

                st.dataframe(
                    review_df,
                    use_container_width=True,
                    hide_index=True,
                    height=min(
                        430,
                        40 + max(1, len(review_df)) * 35,
                    ),
                    column_config={
                        "Onboarded At":
                            st.column_config.DatetimeColumn(
                                "Onboarded At",
                                format="DD MMM YYYY HH:mm",
                            ),
                        "Leads":
                            st.column_config.NumberColumn(
                                "Leads",
                                format="%d",
                            ),
                        "Converted":
                            st.column_config.NumberColumn(
                                "Converted",
                                format="%d",
                            ),
                        "Conversion %":
                            st.column_config.NumberColumn(
                                "Conversion %",
                                format="%.1f%%",
                            ),
                        "Lead Project Value":
                            st.column_config.NumberColumn(
                                "Lead Project Value",
                                format="₹ %.0f",
                            ),
                        "Converted Value":
                            st.column_config.NumberColumn(
                                "Converted Value",
                                format="₹ %.0f",
                            ),
                        "Disbursed Value":
                            st.column_config.NumberColumn(
                                "Disbursed Value",
                                format="₹ %.0f",
                            ),
                    },
                )

                # ----------------------------------------------------
                # INTERACTIVE INDIVIDUAL LEAD DETAILS
                # Uses the already-loaded lead master only.
                # No API request is made when this section is opened.
                # ----------------------------------------------------
                with st.expander(
                    "Interactive Details — Individual Lead View",
                    expanded=False,
                ):
                    if review_df.empty:
                        st.info(
                            "No partners are available in this onboarding cohort."
                        )
                    else:
                        detail_partner_options = (
                            review_df[name_col]
                            .dropna()
                            .astype(str)
                            .str.strip()
                        )
                        detail_partner_options = [
                            x
                            for x in detail_partner_options.tolist()
                            if x
                        ]

                        if not detail_partner_options:
                            st.info(
                                "No partner names are available for individual lead details."
                            )
                        else:
                            selected_detail_partner = st.selectbox(
                                (
                                    "Select Vendor"
                                    if clicked_type == "Vendor"
                                    else "Select CP"
                                ),
                                options=detail_partner_options,
                                key=(
                                    "onboarding_individual_partner_"
                                    + clicked_type
                                    + "_"
                                    + str(clicked_bucket_id)
                                ),
                            )

                            individual_leads = perf.loc[
                                perf["Vendor / CP Name"]
                                .astype(str)
                                .str.strip()
                                == selected_detail_partner
                            ].copy()

                            individual_lead_count = int(
                                individual_leads["Lead Number"].nunique()
                            )
                            individual_converted = int(
                                individual_leads.loc[
                                    individual_leads["_is_project"],
                                    "Lead Number",
                                ].nunique()
                            )
                            individual_cvr = (
                                individual_converted
                                / individual_lead_count
                                * 100
                                if individual_lead_count
                                else 0.0
                            )
                            individual_project_value = float(
                                individual_leads["_pv"].sum()
                            )
                            individual_disbursed_value = float(
                                individual_leads["_dv"].sum()
                            )

                            st.markdown(
                                f"**{selected_detail_partner}**"
                            )

                            i1, i2, i3, i4 = st.columns(4)
                            i1.metric(
                                "Leads",
                                f"{individual_lead_count:,}",
                            )
                            i2.metric(
                                "Converted",
                                f"{individual_converted:,}",
                                f"{individual_cvr:.1f}%",
                            )
                            i3.metric(
                                "Project Value",
                                money_short(
                                    individual_project_value
                                ),
                            )
                            i4.metric(
                                "Disbursed Value",
                                money_short(
                                    individual_disbursed_value
                                ),
                            )

                            lead_detail_columns = [
                                c
                                for c in [
                                    "Lead Number",
                                    "Customer Name",
                                    "Stage",
                                    "Lead Sub Stage",
                                    "Customer Region",
                                    "Vendor Region",
                                    "NBFC",
                                    "Lead Created At",
                                    "Disbursed At",
                                    "Project Value",
                                    "Dynamic Pricing",
                                    "Disbursed Value",
                                ]
                                if c in individual_leads.columns
                            ]

                            individual_lead_view = (
                                individual_leads[lead_detail_columns]
                                .drop_duplicates(
                                    subset=["Lead Number"],
                                    keep="first",
                                )
                                .sort_values(
                                    "Lead Created At"
                                    if "Lead Created At"
                                    in lead_detail_columns
                                    else "Lead Number",
                                    ascending=False,
                                )
                                .reset_index(drop=True)
                            )

                            if individual_lead_view.empty:
                                st.info(
                                    "This partner currently has no lead records in the loaded lead master."
                                )
                            else:
                                st.dataframe(
                                    individual_lead_view,
                                    use_container_width=True,
                                    hide_index=True,
                                    height=min(
                                        390,
                                        40
                                        + max(
                                            1,
                                            len(individual_lead_view),
                                        )
                                        * 35,
                                    ),
                                    column_config={
                                        "Lead Created At":
                                            st.column_config.DatetimeColumn(
                                                "Lead Created At",
                                                format="DD MMM YYYY HH:mm",
                                            ),
                                        "Disbursed At":
                                            st.column_config.DatetimeColumn(
                                                "Disbursed At",
                                                format="DD MMM YYYY HH:mm",
                                            ),
                                        "Project Value":
                                            st.column_config.NumberColumn(
                                                "Project Value",
                                                format="₹ %.0f",
                                            ),
                                        "Dynamic Pricing":
                                            st.column_config.NumberColumn(
                                                "Dynamic Pricing",
                                                format="₹ %.0f",
                                            ),
                                        "Disbursed Value":
                                            st.column_config.NumberColumn(
                                                "Disbursed Value",
                                                format="₹ %.0f",
                                            ),
                                    },
                                )

            show_onboarding_popup()

    elif onboarding_event is not None:
        st.caption(
            "Click a CP or Vendor onboarding bar to open its performance review."
        )
    else:
        st.caption(
            "Upgrade Streamlit to enable click-to-popup chart drill-down."
        )




st.markdown(
    """
    <style>
    /* ==========================================================
       RENGY ROW 2 — FINAL VISUAL SYSTEM
       ========================================================== */

    /* Kill Streamlit's internal vertical spacing ONLY inside Row 2 boards. */
    [class*="st-key-r2_sales_board"] div[data-testid="stVerticalBlock"],
    [class*="st-key-r2_pipeline_board"] div[data-testid="stVerticalBlock"] {
        gap: 0 !important;
    }

    [class*="st-key-r2_sales_board"],
    [class*="st-key-r2_pipeline_board"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    /* ---------- Sales legend ---------- */
    [class*="st-key-r2_sales_legend_"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    [class*="st-key-r2_sales_legend_"] div[data-testid="stButton"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    [class*="st-key-r2_sales_legend_"] button {
        width: 100% !important;
        min-height: 1.52rem !important;
        height: 1.52rem !important;
        margin: 0 !important;
        padding: 0 .38rem !important;
        border-radius: 7px !important;
        border: 1px solid #dbe5ec !important;
        background: rgba(255,255,255,.82) !important;
        color: #17364a !important;
        box-shadow: none !important;
        justify-content: flex-start !important;
        text-align: left !important;
        font-size: .61rem !important;
        font-weight: 720 !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
    }

    [class*="st-key-r2_sales_legend_"] button p {
        margin: 0 !important;
        line-height: 1 !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        white-space: nowrap !important;
    }

    [class*="st-key-r2_sales_legend_"] button:hover {
        background: #ffffff !important;
        border-color: #b9cfdd !important;
        transform: none !important;
        box-shadow: 0 2px 7px rgba(15,39,64,.05) !important;
    }

    /* ---------- Pipeline cells ---------- */
    [class*="st-key-r2_pipe_"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    [class*="st-key-r2_pipe_"] div[data-testid="stButton"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    [class*="st-key-r2_pipe_"] button {
        width: 100% !important;
        min-height: 2.00rem !important;
        height: 2.00rem !important;
        margin: 0 !important;
        padding: 0 .06rem !important;
        border-radius: 0 !important;
        border-width: 1px !important;
        border-style: solid !important;
        box-shadow: none !important;
        font-size: .63rem !important;
        font-weight: 850 !important;
        line-height: 1 !important;
        letter-spacing: -.01em !important;
        white-space: nowrap !important;
    }

    [class*="st-key-r2_pipe_"] button p {
        margin: 0 !important;
        line-height: 1 !important;
        white-space: nowrap !important;
    }

    [class*="st-key-r2_pipe_"] button:hover:not(:disabled) {
        filter: brightness(.965) !important;
        transform: none !important;
        position: relative !important;
        z-index: 2 !important;
    }

    [class*="st-key-r2_pipe_"] button:disabled {
        opacity: 1 !important;
        background: #f8fafb !important;
        color: #b5c0c9 !important;
        border-color: #dce5eb !important;
    }

    div[data-testid="stDialog"] {
        transition: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# ROW 2 — RELIABLE CLICKABLE REGION SALES + PENDING PIPELINE
# IMPORTANT:
# Plotly Pie/Heatmap selection is not equally reliable across Streamlit
# versions. Visuals stay premium, while native Streamlit click targets
# guarantee that the drill-down popup works on Streamlit Cloud.
# No API call is made on click.
# ============================================================


row2_left, row2_right = st.columns([0.78, 1.72], gap="medium")

# ------------------------------------------------------------
# ROW 2 / LEFT — REGION x SALES
# Premium live donut + reliable native region click controls.
# ------------------------------------------------------------
with row2_left:
    st.markdown(
        """
        <div style="margin-bottom:.12rem;">
            <div class="chart-kicker" style="margin-bottom:.02rem;">SALES DISTRIBUTION</div>
            <div class="chart-title" style="font-size:.92rem;line-height:1.02;">Region × Sales</div>
            <div class="chart-note" style="margin:.04rem 0 0 0;font-size:.64rem;">
                Converted project value • Click a region for details
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    region_sales_base = project_df.loc[
        project_df["Customer Region"].notna()
        & project_df["Customer Region"].astype(str).str.strip().ne("")
    ].copy()

    if region_sales_base.empty:
        st.info("No converted project value is available for the selected filters.")
    else:
        region_sales_base["_region"] = (
            region_sales_base["Customer Region"]
            .astype(str)
            .str.strip()
        )

        region_sales_base["_pv"] = pd.to_numeric(
            region_sales_base["Project Value"],
            errors="coerce",
        ).fillna(0)

        region_sales_base["_dv"] = pd.to_numeric(
            region_sales_base["Disbursed Value"],
            errors="coerce",
        ).fillna(0)

        region_sales = (
            region_sales_base
            .groupby("_region", observed=True)
            .agg(
                Project_Value=("_pv", "sum"),
                Projects=("Lead Number", "nunique"),
                Disbursed_Value=("_dv", "sum"),
            )
            .reset_index()
            .sort_values("Project_Value", ascending=False)
            .reset_index(drop=True)
        )

        total_region_sales = float(
            region_sales["Project_Value"].sum()
        )

        total_region_projects = int(
            region_sales_base["Lead Number"].nunique()
        )

        region_sales["Share"] = (
            region_sales["Project_Value"]
            .div(total_region_sales if total_region_sales else 1)
            .mul(100)
        )

        # Stable, preformatted hover content.
        # This avoids Plotly customdata values becoming NaN on Streamlit Cloud.
        region_sales["Projects"] = pd.to_numeric(
            region_sales["Projects"], errors="coerce"
        ).fillna(0).astype(int)
        region_sales["Share"] = pd.to_numeric(
            region_sales["Share"], errors="coerce"
        ).fillna(0.0)
        region_sales["Disbursed_Value"] = pd.to_numeric(
            region_sales["Disbursed_Value"], errors="coerce"
        ).fillna(0.0)
        region_sales["Project_Value"] = pd.to_numeric(
            region_sales["Project_Value"], errors="coerce"
        ).fillna(0.0)

        pie_hover_text = [
            (
                f"<b>{region}</b>"
                f"<br>Sales&nbsp;&nbsp;<b>₹{sales:,.0f}</b>"
                f"<br>Projects&nbsp;&nbsp;<b>{projects:,}</b>"
                f"<br>Share&nbsp;&nbsp;<b>{share:.1f}%</b>"
                f"<br>Disbursed&nbsp;&nbsp;<b>₹{disbursed:,.0f}</b>"
            )
            for region, sales, projects, share, disbursed in zip(
                region_sales["_region"].astype(str),
                region_sales["Project_Value"],
                region_sales["Projects"],
                region_sales["Share"],
                region_sales["Disbursed_Value"],
            )
        ]

        # Executive Rengy palette: blue -> cyan -> teal -> green,
        # with warm accents only for smaller trailing regions.
        region_palette = [
            "#0F766E",  # emerald
            "#B45309",  # burnt amber
            "#7C3AED",  # royal violet
            "#BE123C",  # wine
            "#0369A1",  # steel blue
            "#4D7C0F",  # olive
            "#9A3412",  # rust
            "#4338CA",  # indigo
            "#A16207",  # ochre
            "#475569",  # slate
        ]

        donut_colors = [
            region_palette[i % len(region_palette)]
            for i in range(len(region_sales))
        ]

        donut_pull = [0.0 for _ in range(len(region_sales))]

        region_fig = go.Figure(
            go.Pie(
                labels=region_sales["_region"],
                domain=dict(x=[0.06, 0.94], y=[0.03, 0.97]),
                values=region_sales["Project_Value"],
                hole=0.66,
                sort=False,
                direction="clockwise",
                rotation=225,
                pull=donut_pull,
                hovertext=pie_hover_text,
                texttemplate="<b>%{percent:.1%}</b>",
                textposition="inside",
                textfont=dict(size=9),
                insidetextorientation="horizontal",
                marker=dict(
                    colors=donut_colors,
                    line=dict(
                        color="rgba(255,255,255,.98)",
                        width=1.6,
                    ),
                ),
                hovertemplate="%{hovertext}<extra></extra>",
            )
        )

        # Match the right-hand pipeline panel visually.
        # The chart itself gets a taller canvas and the region controls
        # sit inside the same overall vertical rhythm.
        region_fig.update_layout(
            height=248,
            margin=dict(l=2, r=2, t=0, b=0),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(
                family="Inter, Arial, sans-serif",
                color="#334155",
            ),
            hoverlabel=dict(
                bgcolor="#102A43",
                bordercolor="#274C6B",
                font=dict(
                    color="#F8FAFC",
                    size=10,
                    family="Inter, Arial, sans-serif",
                ),
                align="left",
            ),
            # Custom native legend below opens Streamlit dialogs directly.
            showlegend=False,
            annotations=[
                dict(
                    text=(
                        f"<span style='font-size:16px'><b>{money_short(total_region_sales)}</b></span>"
                        "<br><span style='font-size:9px;color:#71869a;letter-spacing:1px'>SALES</span>"
                        f"<br><span style='font-size:9px;color:#94a3b8'>{total_region_projects:,} projects</span>"
                    ),
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    align="center",
                    font=dict(
                        size=13,
                        color="#0f2740",
                    ),
                )
            ],
            uirevision="region-sales-live",
        )

        st.plotly_chart(
            region_fig,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "responsive": True,
                "scrollZoom": False,
            },
            key="row2_region_sales_live_donut",
        )

        # Native clickable legend — no URL navigation / no page refresh.
        region_options = region_sales["_region"].tolist()
        selected_region_click = None

        region_color_map = {
            region: donut_colors[idx]
            for idx, region in enumerate(region_options)
        }

        # Per-region color rail on each native Streamlit button.
        legend_css_rules = []
        for legend_index, region_name in enumerate(region_options):
            dot_color = region_color_map.get(region_name, "#96A8B4")
            legend_css_rules.append(
                f"""
                [class*="st-key-r2_sales_legend_{legend_index}_"] button {{
                    border-left: 5px solid {dot_color} !important;
                }}
                """
            )

        st.markdown(
            "<style>" + "".join(legend_css_rules) + "</style>",
            unsafe_allow_html=True,
        )

        with st.container(key="r2_sales_board"):
            legend_rows = [
                region_options[i:i + 3]
                for i in range(0, len(region_options), 3)
            ]

            for legend_row_index, legend_row in enumerate(legend_rows):
                legend_cols = st.columns(3, gap="small")

                for legend_col_index in range(3):
                    with legend_cols[legend_col_index]:
                        if legend_col_index < len(legend_row):
                            region_name = legend_row[legend_col_index]
                            legend_index = region_options.index(region_name)

                            if st.button(
                                region_name,
                                key=(
                                    f"r2_sales_legend_"
                                    f"{legend_index}_"
                                    f"{region_name}"
                                ),
                                use_container_width=True,
                            ):
                                selected_region_click = region_name
                        else:
                            st.markdown(
                                "<div style='height:1.52rem;'></div>",
                                unsafe_allow_html=True,
                            )

        if selected_region_click:
            region_detail = region_sales_base.loc[
                region_sales_base["_region"]
                == str(selected_region_click)
            ].copy()

            r_projects = int(
                region_detail["Lead Number"].nunique()
            )

            r_value = float(
                region_detail["_pv"].sum()
            )

            r_disbursed = float(
                region_detail["_dv"].sum()
            )

            r_share = (
                r_value / total_region_sales * 100
                if total_region_sales
                else 0.0
            )

            r_avg_ticket = (
                r_value / r_projects
                if r_projects
                else 0.0
            )

            @st.dialog(
                f"{selected_region_click} — Region Sales Review",
                width="large",
            )
            def show_region_sales_popup():
                st.caption(
                    "Converted Project Review • Current dashboard filters applied • No additional API call"
                )

                x1, x2, x3, x4, x5 = st.columns(5)

                x1.metric(
                    "Projects",
                    f"{r_projects:,}",
                )

                x2.metric(
                    "Sales Value",
                    money_short(r_value),
                )

                x3.metric(
                    "Sales Share",
                    f"{r_share:.1f}%",
                )

                x4.metric(
                    "Disbursed",
                    money_short(r_disbursed),
                )

                x5.metric(
                    "Avg Project",
                    money_short(r_avg_ticket),
                )

                st.markdown("#### Individual Project Details")

                region_cols = [
                    c for c in [
                        "Lead Number",
                        "Customer Name",
                        "Vendor / CP Name",
                        "Type",
                        "Customer Region",
                        "Vendor Region",
                        "Stage",
                        "Lead Sub Stage",
                        "NBFC",
                        "Lead Created At",
                        "Disbursed At",
                        "Project Value",
                        "Dynamic Pricing",
                        "Disbursed Value",
                    ]
                    if c in region_detail.columns
                ]

                region_view = (
                    region_detail[region_cols]
                    .drop_duplicates(
                        subset=["Lead Number"],
                        keep="first",
                    )
                    .sort_values(
                        "Project Value",
                        ascending=False,
                    )
                    .reset_index(drop=True)
                )

                st.dataframe(
                    region_view,
                    use_container_width=True,
                    hide_index=True,
                    height=min(
                        500,
                        42 + max(1, len(region_view)) * 35,
                    ),
                    column_config={
                        "Lead Created At":
                            st.column_config.DatetimeColumn(
                                "Lead Created At",
                                format="DD MMM YYYY HH:mm",
                            ),
                        "Disbursed At":
                            st.column_config.DatetimeColumn(
                                "Disbursed At",
                                format="DD MMM YYYY HH:mm",
                            ),
                        "Project Value":
                            st.column_config.NumberColumn(
                                "Project Value",
                                format="₹ %.0f",
                            ),
                        "Dynamic Pricing":
                            st.column_config.NumberColumn(
                                "Dynamic Pricing",
                                format="₹ %.0f",
                            ),
                        "Disbursed Value":
                            st.column_config.NumberColumn(
                                "Disbursed Value",
                                format="₹ %.0f",
                            ),
                    },
                )

            show_region_sales_popup()

        else:
            pass

        # No footer/caption here: keeps the left panel on the same baseline
        # as the right-side pending table.
        # Intentionally no footer: left and right panels end on the same rhythm.


# ------------------------------------------------------------
# ROW 2 / RIGHT — CLICKABLE REGION-WISE PENDING PIPELINE
# Native Streamlit buttons preserve the reliable st.dialog behavior.
# ------------------------------------------------------------
with row2_right:
    st.markdown(
        """
        <div style="margin-bottom:.12rem;">
            <div class="chart-kicker" style="margin-bottom:.02rem;">PENDING FUNNEL</div>
            <div class="chart-title" style="font-size:.92rem;line-height:1.02;">
                Overall Region-wise Pending Pipeline
            </div>
            <div class="chart-note" style="margin:.04rem 0 .16rem 0;font-size:.62rem;">
                Count | avg age • darker coral = higher pending • click any non-zero cell
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pipeline = pending_df.loc[
        pending_df["Lead Number"].notna()
    ].copy()

    if pipeline.empty:
        st.info(
            "No pending pipeline is available for the selected filters."
        )
    else:
        today_ist = pd.Timestamp.now(
            tz="Asia/Kolkata"
        ).tz_localize(None).normalize()

        pipeline["_age_days"] = (
            today_ist
            - pipeline["Lead Created At"].dt.normalize()
        ).dt.days.clip(lower=0)

        pipeline["_region"] = (
            pipeline["Customer Region"]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
            .replace("", "Unknown")
        )

        pipeline["_stage"] = (
            pipeline["Stage"]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
            .replace("", "Unknown")
        )

        pipeline = (
            pipeline
            .sort_values(
                ["Lead Number", "Lead Created At"],
                ascending=[True, False],
            )
            .drop_duplicates(
                subset=["Lead Number"],
                keep="first",
            )
            .copy()
        )

        preferred_regions = [
            "Secunderabad",
            "Warangal",
            "Nizamabad",
            "Rayalaseema",
            "Vijayawada",
            "Nellore",
            "Vizag",
            "Kadapa",
            "Berhampur",
            "Cuttack",
            "Unknown",
        ]

        present_regions = (
            pipeline["_region"]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[lambda s: s.ne("")]
            .unique()
            .tolist()
        )

        region_lookup = {
            str(x).strip().lower(): str(x).strip()
            for x in present_regions
        }

        pipeline_regions = []

        for region in preferred_regions:
            actual = region_lookup.get(region.lower())
            if actual and actual not in pipeline_regions:
                pipeline_regions.append(actual)

        for region in sorted(present_regions):
            if region not in pipeline_regions:
                pipeline_regions.append(region)

        stage_totals = (
            pipeline
            .groupby("_stage", observed=True)["Lead Number"]
            .nunique()
            .sort_values(ascending=False)
        )

        pipeline_stages = stage_totals.index.tolist()

        # Reference-style table proportions.
        matrix_widths = [1.28] + [1.0] * len(pipeline_regions) + [0.72]

        # Pre-calculate cell counts so heat intensity is relative to live data.
        cell_counts = {}
        max_cell_count = 1

        for stage in pipeline_stages:
            for region in pipeline_regions:
                count = int(
                    pipeline.loc[
                        (pipeline["_stage"] == stage)
                        & (pipeline["_region"] == region),
                        "Lead Number",
                    ].nunique()
                )
                cell_counts[(stage, region)] = count
                max_cell_count = max(max_cell_count, count)

        def _heat_colors(count):
            if count <= 0:
                return "#F8FAFB", "#DCE5EB", "#B2BEC8"

            ratio = min(1.0, count / max_cell_count)

            # Soft coral -> stronger coral. Executive, not aggressive.
            light = (255, 240, 239)
            dark = (232, 128, 132)

            rgb = tuple(
                round(light[i] + (dark[i] - light[i]) * ratio)
                for i in range(3)
            )

            bg = f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"
            border = "#F3D6A4" if ratio < .45 else "#D6A34A"
            text = "#12344A"
            return bg, border, text

        # Generate CSS per live cell so native buttons become a true heat table.
        heat_css = []

        for stage_index, stage in enumerate(pipeline_stages):
            for region_index, region in enumerate(pipeline_regions):
                count = cell_counts[(stage, region)]
                bg, border, text_color = _heat_colors(count)

                heat_css.append(
                    f"""
                    [class*="st-key-r2_pipe_{stage_index}_{region_index}_"] button {{
                        background:{bg} !important;
                        border-color:{border} !important;
                        color:{text_color} !important;
                    }}
                    """
                )

        st.markdown(
            "<style>" + "".join(heat_css) + "</style>",
            unsafe_allow_html=True,
        )

        clicked_pipeline_cell = None

        with st.container(key="r2_pipeline_board"):
            # Header
            header_cols = st.columns(matrix_widths, gap=None)

            with header_cols[0]:
                st.markdown(
                    """
                    <div style="
                        height:2rem;
                        display:flex;
                        align-items:center;
                        padding:0 .48rem;
                        border:1px solid #C9D7DF;
                        background:#ECFDF5;
                        color:#102F45;
                        font-size:.60rem;
                        font-weight:900;
                    ">STAGE</div>
                    """,
                    unsafe_allow_html=True,
                )

            for idx, region in enumerate(pipeline_regions, start=1):
                with header_cols[idx]:
                    st.markdown(
                        f"""
                        <div title="{region}" style="
                            height:2rem;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            padding:0 .08rem;
                            border:1px solid #C9D7DF;
                            border-left:0;
                            background:#F0FDF4;
                            color:#102F45;
                            font-size:.52rem;
                            font-weight:850;
                            line-height:1;
                            white-space:nowrap;
                            overflow:hidden;
                            text-overflow:ellipsis;
                        ">{region}</div>
                        """,
                        unsafe_allow_html=True,
                    )

            with header_cols[-1]:
                st.markdown(
                    """
                    <div style="
                        height:2rem;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        border:1px solid #C9D7DF;
                        border-left:0;
                        background:#DCE8F0;
                        color:#102F45;
                        font-size:.57rem;
                        font-weight:900;
                    ">TOTAL</div>
                    """,
                    unsafe_allow_html=True,
                )

            # Body rows
            for stage_index, stage in enumerate(pipeline_stages):
                row_cols = st.columns(matrix_widths, gap=None)

                with row_cols[0]:
                    stage_label = (
                        str(stage)
                        .replace("_", " ")
                        .replace("pricingQuotation", "Pricing Quotation")
                        .replace("leadDetails", "Lead Details")
                        .replace("siteSurvey", "Site Survey")
                        .upper()
                    )

                    st.markdown(
                        f"""
                        <div title="{stage_label}" style="
                            height:2rem;
                            display:flex;
                            align-items:center;
                            padding:0 .48rem;
                            border:1px solid #D3DEE6;
                            border-top:0;
                            background:#F3F7F9;
                            color:#102F45;
                            font-size:.55rem;
                            font-weight:900;
                            line-height:1.05;
                            overflow:hidden;
                        ">{stage_label}</div>
                        """,
                        unsafe_allow_html=True,
                    )

                for region_index, region in enumerate(pipeline_regions):
                    cell = pipeline.loc[
                        (pipeline["_stage"] == stage)
                        & (pipeline["_region"] == region)
                    ].copy()

                    count = cell_counts[(stage, region)]

                    avg_age = (
                        float(cell["_age_days"].mean())
                        if count
                        else 0.0
                    )

                    with row_cols[region_index + 1]:
                        clicked = st.button(
                            f"{count} | {avg_age:.0f}d",
                            key=(
                                f"r2_pipe_"
                                f"{stage_index}_"
                                f"{region_index}_"
                                f"{stage}_"
                                f"{region}"
                            ),
                            use_container_width=True,
                            disabled=(count == 0),
                        )

                    if clicked:
                        clicked_pipeline_cell = (stage, region)

                stage_total_count = int(
                    pipeline.loc[
                        pipeline["_stage"] == stage,
                        "Lead Number",
                    ].nunique()
                )

                with row_cols[-1]:
                    st.markdown(
                        f"""
                        <div style="
                            height:2rem;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            border:1px solid #C9D7DF;
                            border-top:0;
                            border-left:0;
                            background:#E6EEF4;
                            color:#102F45;
                            font-size:.62rem;
                            font-weight:950;
                        ">{stage_total_count:,}</div>
                        """,
                        unsafe_allow_html=True,
                    )

            # Bottom totals
            total_cols = st.columns(matrix_widths, gap=None)

            with total_cols[0]:
                st.markdown(
                    """
                    <div style="
                        height:2rem;
                        display:flex;
                        align-items:center;
                        padding:0 .48rem;
                        border:1px solid #C9D7DF;
                        border-top:0;
                        background:#DCE8F0;
                        color:#102F45;
                        font-size:.58rem;
                        font-weight:950;
                    ">TOTAL</div>
                    """,
                    unsafe_allow_html=True,
                )

            for region_index, region in enumerate(pipeline_regions, start=1):
                region_total_count = int(
                    pipeline.loc[
                        pipeline["_region"] == region,
                        "Lead Number",
                    ].nunique()
                )

                with total_cols[region_index]:
                    st.markdown(
                        f"""
                        <div style="
                            height:2rem;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            border:1px solid #C9D7DF;
                            border-top:0;
                            border-left:0;
                            background:#E6EEF4;
                            color:#102F45;
                            font-size:.58rem;
                            font-weight:900;
                        ">{region_total_count:,}</div>
                        """,
                        unsafe_allow_html=True,
                    )

            grand_total_count = int(
                pipeline["Lead Number"].nunique()
            )

            with total_cols[-1]:
                st.markdown(
                    f"""
                    <div style="
                        height:2rem;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        border:1px solid #BFCFD9;
                        border-top:0;
                        border-left:0;
                        background:#D7E5EE;
                        color:#102F45;
                        font-size:.61rem;
                        font-weight:950;
                    ">{grand_total_count:,}</div>
                    """,
                    unsafe_allow_html=True,
                )

        if clicked_pipeline_cell:
            clicked_stage, clicked_pipeline_region = (
                clicked_pipeline_cell
            )

            cell_detail = pipeline.loc[
                (pipeline["_stage"] == clicked_stage)
                & (
                    pipeline["_region"]
                    == clicked_pipeline_region
                )
            ].copy()

            cell_count = int(
                cell_detail["Lead Number"].nunique()
            )

            cell_avg_age = (
                float(
                    cell_detail["_age_days"].mean()
                )
                if cell_count
                else 0.0
            )

            cell_oldest = (
                int(
                    cell_detail["_age_days"].max()
                )
                if cell_count
                else 0
            )

            cell_value = float(
                pd.to_numeric(
                    cell_detail["Project Value"],
                    errors="coerce",
                ).fillna(0).sum()
            )

            cell_disbursed = float(
                pd.to_numeric(
                    cell_detail["Disbursed Value"],
                    errors="coerce",
                ).fillna(0).sum()
            )

            @st.dialog(
                f"{clicked_stage} × {clicked_pipeline_region}",
                width="large",
            )
            def show_pipeline_popup():
                st.caption(
                    "Pending Pipeline Review • Age = Lead Created At → today • No additional API call"
                )

                p1, p2, p3, p4, p5 = st.columns(5)

                p1.metric(
                    "Pending Leads",
                    f"{cell_count:,}",
                )

                p2.metric(
                    "Avg Age",
                    f"{cell_avg_age:.1f}d",
                )

                p3.metric(
                    "Oldest",
                    f"{cell_oldest:,}d",
                )

                p4.metric(
                    "Lead Value",
                    money_short(cell_value),
                )

                p5.metric(
                    "Disbursed",
                    money_short(cell_disbursed),
                )

                st.markdown(
                    "#### Individual Pending Lead Details"
                )

                detail_cols = [
                    c for c in [
                        "Lead Number",
                        "Customer Name",
                        "Vendor / CP Name",
                        "Type",
                        "Customer Region",
                        "Vendor Region",
                        "Stage",
                        "Lead Sub Stage",
                        "NBFC",
                        "Lead Created At",
                        "Project Value",
                        "Dynamic Pricing",
                        "Disbursed Value",
                    ]
                    if c in cell_detail.columns
                ]

                pipeline_view = (
                    cell_detail[
                        detail_cols
                        + ["_age_days"]
                    ]
                    .rename(
                        columns={
                            "_age_days":
                                "Pending Age (Days)"
                        }
                    )
                    .sort_values(
                        "Pending Age (Days)",
                        ascending=False,
                    )
                    .reset_index(drop=True)
                )

                st.dataframe(
                    pipeline_view,
                    use_container_width=True,
                    hide_index=True,
                    height=min(
                        510,
                        42
                        + max(
                            1,
                            len(pipeline_view),
                        )
                        * 35,
                    ),
                    column_config={
                        "Lead Created At":
                            st.column_config.DatetimeColumn(
                                "Lead Created At",
                                format="DD MMM YYYY HH:mm",
                            ),
                        "Pending Age (Days)":
                            st.column_config.NumberColumn(
                                "Pending Age (Days)",
                                format="%d d",
                            ),
                        "Project Value":
                            st.column_config.NumberColumn(
                                "Project Value",
                                format="₹ %.0f",
                            ),
                        "Dynamic Pricing":
                            st.column_config.NumberColumn(
                                "Dynamic Pricing",
                                format="₹ %.0f",
                            ),
                        "Disbursed Value":
                            st.column_config.NumberColumn(
                                "Disbursed Value",
                                format="₹ %.0f",
                            ),
                    },
                )

            show_pipeline_popup()



# ============================================================
# INTENTIONALLY STOP HERE FOR STEP 1
# ============================================================




# ============================================================
# ROW 3 — CONSULTANT LEADERBOARD
# ============================================================
# Clean full-width leaderboard.
# - All Regions -> Top 5 by highest Project count
# - Specific Region -> all consultants in that region
# - Region selector sits at top-right
# - Rank symbols live directly above the PROJECT bars
# - No extra leaderboard cards / button strip
# - Click the chart bar to open the consultant dossier
# ============================================================

def consultant_key(value):
    """Stable local key helper for consultant widgets."""
    text = str(value or "").strip().casefold()
    cleaned = "".join(
        ch if ch.isalnum() else "_"
        for ch in text
    )
    cleaned = "_".join(
        part
        for part in cleaned.split("_")
        if part
    )
    return cleaned[:70] or "unassigned"


st.markdown(
    """
    <style>
    [class*="st-key-consultant_vibe_board"] {
        border:1px solid rgba(15,39,64,.075);
        border-radius:20px;
        padding:.70rem .95rem .50rem .95rem;
        background:linear-gradient(180deg,#FFFFFF 0%,#F8FAFC 100%);
        border-top:3px solid #0F766E;
        box-shadow:0 12px 32px rgba(15,39,64,.055);
        overflow:hidden;
    }

    [class*="st-key-consultant_vibe_board"]
    div[data-testid="stVerticalBlock"] {
        gap:.18rem;
    }

    [class*="st-key-consultant_region_compact"]
    div[data-baseweb="select"] > div {
        min-height:2.05rem !important;
        border-radius:10px !important;
        border-color:rgba(15,39,64,.09) !important;
        background:#F8FAFC !important;
        font-size:.75rem !important;
    }

    [class*="st-key-consultant_region_compact"]
    label {
        display:none !important;
    }

    .consultant-vibe-kicker {
        color:#0F766E;
        font-size:.60rem;
        font-weight:900;
        letter-spacing:.12em;
        text-transform:uppercase;
        margin-bottom:.06rem;
    }

    .consultant-vibe-title {
        color:#102f46;
        font-size:1.05rem;
        font-weight:900;
        letter-spacing:-.025em;
        line-height:1.12;
    }

    .consultant-vibe-note {
        color:#64748B;
        font-size:.65rem;
        margin-top:.10rem;
    }

    .consultant-view-pill {
        display:inline-flex;
        align-items:center;
        border:1px solid rgba(21,155,121,.13);
        background:#F0FDF4;
        color:#0F766E;
        border-radius:999px;
        padding:.18rem .48rem;
        font-size:.56rem;
        font-weight:900;
        letter-spacing:.055em;
        white-space:nowrap;
    }

    .consultant-click-note {
        color:#64748B;
        font-size:.60rem;
        text-align:right;
        margin-top:-.18rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.container(key="consultant_vibe_board"):

    consultant_source = filtered.copy()

    consultant_source["Owner Name"] = (
        consultant_source["Owner Name"]
        .fillna("Unassigned")
        .astype(str)
        .str.strip()
        .replace("", "Unassigned")
    )

    consultant_source = (
        consultant_source
        .sort_values(
            ["Lead Number", "Lead Created At"],
            ascending=[True, False],
            na_position="last",
        )
        .drop_duplicates(
            subset=["Lead Number"],
            keep="first",
        )
        .copy()
    )

    consultant_region_values = [
        str(x).strip()
        for x in consultant_source["Customer Region"]
        .dropna()
        .astype(str)
        .unique()
        if str(x).strip()
    ]

    preferred_consultant_region_order = [
        "Secunderabad",
        "Warangal",
        "Nizamabad",
        "Rayalaseema",
        "Vijayawada",
        "Nellore",
        "Vizag",
        "Kadapa",
        "Berhampur",
        "Cuttack",
        "Unknown",
    ]

    consultant_region_options = [
        region
        for region in preferred_consultant_region_order
        if region in consultant_region_values
    ] + sorted(
        [
            region
            for region in consultant_region_values
            if region not in preferred_consultant_region_order
        ]
    )

    # --------------------------------------------------------
    # HEADER LEFT + COMPACT REGION FILTER TOP-RIGHT
    # --------------------------------------------------------
    consultant_head_left, consultant_head_right = st.columns(
        [0.78, 0.22],
        gap="medium",
        vertical_alignment="center",
    )

    with consultant_head_left:
        st.markdown(
            """
            <div class="consultant-vibe-kicker">Sales Leaderboard</div>
            <div class="consultant-vibe-title">Consultant Champions</div>
            <div class="consultant-vibe-note">
                Leads vs converted projects • ranking is driven by project wins
            </div>
            """,
            unsafe_allow_html=True,
        )

    with consultant_head_right:
        with st.container(
            key="consultant_region_compact"
        ):
            selected_consultant_region = st.selectbox(
                "Region",
                ["All Regions"] + consultant_region_options,
                key="consultant_leaderboard_region_v2",
                label_visibility="collapsed",
            )

    region_is_selected = (
        selected_consultant_region != "All Regions"
    )

    if region_is_selected:
        consultant_source = consultant_source.loc[
            consultant_source["Customer Region"].eq(
                selected_consultant_region
            )
        ].copy()

    if consultant_source.empty:
        st.info(
            "No consultant records are available for the selected region and dashboard filters."
        )

    else:
        consultant_source["_period_project_win"] = (
            consultant_source["_is_project"] & consultant_source["_period_project"]
        ).astype("int8")
        consultant_source["_period_lead"] = consultant_source["_period_lead"].astype("int8")
        consultant_source["_period_project_value"] = np.where(
            consultant_source["_period_project_win"].eq(1),
            consultant_source["Project Value"],
            0.0,
        )
        consultant_source["_period_disbursed_value"] = np.where(
            consultant_source["_period_project_win"].eq(1),
            consultant_source["Disbursed Value"],
            0.0,
        )

        consultant_summary = (
            consultant_source
            .groupby(
                "Owner Name",
                observed=True,
                dropna=False,
            )
            .agg(
                Leads=("_period_lead", "sum"),
                Projects=("_period_project_win", "sum"),
                Project_Value=("_period_project_value", "sum"),
                Disbursed_Value=("_period_disbursed_value", "sum"),
            )
            .reset_index()
        )

        consultant_summary["Leads"] = pd.to_numeric(
            consultant_summary["Leads"],
            errors="coerce",
        ).fillna(0).astype(int)

        consultant_summary["Projects"] = pd.to_numeric(
            consultant_summary["Projects"],
            errors="coerce",
        ).fillna(0).astype(int)

        consultant_summary["Project_Value"] = pd.to_numeric(
            consultant_summary["Project_Value"],
            errors="coerce",
        ).fillna(0.0)

        consultant_summary["Disbursed_Value"] = pd.to_numeric(
            consultant_summary["Disbursed_Value"],
            errors="coerce",
        ).fillna(0.0)

        consultant_summary["Pending"] = (
            consultant_summary["Leads"]
            - consultant_summary["Projects"]
        ).clip(lower=0)

        consultant_summary["Conversion %"] = np.where(
            consultant_summary["Leads"] > 0,
            consultant_summary["Projects"]
            / consultant_summary["Leads"]
            * 100,
            0.0,
        )

        consultant_summary = (
            consultant_summary
            .sort_values(
                [
                    "Projects",
                    "Leads",
                    "Project_Value",
                    "Owner Name",
                ],
                ascending=[
                    False,
                    False,
                    False,
                    True,
                ],
            )
            .reset_index(drop=True)
        )

        if region_is_selected:
            display_consultants = (
                consultant_summary.copy()
            )
            view_label = (
                f"{len(display_consultants):,} CONSULTANTS • "
                f"{selected_consultant_region.upper()}"
            )
        else:
            display_consultants = (
                consultant_summary
                .head(5)
                .copy()
            )
            view_label = "TOP 5 • ALL REGIONS"

        display_consultants["Rank"] = np.arange(
            1,
            len(display_consultants) + 1,
        )

        # Only the Top 5 ranking symbols are special.
        # These appear DIRECTLY ON TOP OF PROJECT BARS.
        def rank_symbol(rank):
            rank = int(rank)
            if rank == 1:
                return "👑 #1"
            if rank == 2:
                return "⭐ #2"
            if rank == 3:
                return "🥉 #3"
            return f"⚡ #{rank}"

        display_consultants["Project Bar Label"] = [
            (
                f"{rank_symbol(rank)}"
                f"<br><b>{int(projects):,} P</b>"
            )
            for rank, projects in zip(
                display_consultants["Rank"],
                display_consultants["Projects"],
            )
        ]

        consultant_custom = np.column_stack(
            [
                display_consultants["Owner Name"].astype(str),
                display_consultants["Projects"].astype(int),
                display_consultants["Leads"].astype(int),
                display_consultants["Pending"].astype(int),
                display_consultants["Conversion %"].round(1),
                display_consultants["Project_Value"].round(0),
                display_consultants["Rank"].astype(int),
            ]
        )

        consultant_count = len(display_consultants)

        # Full-width means we can keep the graph compact.
        consultant_chart_height = (
            305
            if consultant_count <= 5
            else min(
                390,
                305
                + max(
                    0,
                    consultant_count - 5,
                )
                * 5,
            )
        )

        # Small status pill consumes almost no vertical space.
        st.markdown(
            f"""
            <div style="
                display:flex;
                justify-content:flex-end;
                margin-top:-.05rem;
                margin-bottom:-.32rem;
                position:relative;
                z-index:2;
            ">
                <span class="consultant-view-pill">
                    {view_label}
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        consultant_fig = go.Figure()

        # LEADS — quiet comparison bar.
        consultant_fig.add_trace(
            go.Bar(
                name="Leads",
                x=display_consultants["Owner Name"],
                y=display_consultants["Leads"],
                width=.24,
                offset=-.13,
                marker=dict(
                    color="#CBD5E1",
                    line=dict(
                        color="rgba(255,255,255,.96)",
                        width=1,
                    ),
                ),
                text=[
                    f"{int(value):,}"
                    for value in display_consultants["Leads"]
                ],
                textposition="outside",
                textfont=dict(
                    size=10,
                    color="#475569",
                ),
                cliponaxis=False,
                customdata=consultant_custom,
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Rank  #%{customdata[6]}<br>"
                    "Leads  %{customdata[2]:,.0f}<br>"
                    "Projects  %{customdata[1]:,.0f}<br>"
                    "Pending  %{customdata[3]:,.0f}<br>"
                    "Conversion  %{customdata[4]:.1f}%<br>"
                    "Project Value  ₹%{customdata[5]:,.0f}"
                    "<extra></extra>"
                ),
            )
        )

        # PROJECTS — hero bar carrying the leaderboard symbol.
        consultant_fig.add_trace(
            go.Bar(
                name="Projects",
                x=display_consultants["Owner Name"],
                y=display_consultants["Projects"],
                width=.24,
                offset=.13,
                marker=dict(
                    color="#0F766E",
                    line=dict(
                        color="rgba(255,255,255,.96)",
                        width=1,
                    ),
                ),
                text=display_consultants[
                    "Project Bar Label"
                ],
                textposition="outside",
                textfont=dict(
                    size=10,
                    color="#0B5F58",
                ),
                cliponaxis=False,
                customdata=consultant_custom,
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Rank  #%{customdata[6]}<br>"
                    "Projects  %{customdata[1]:,.0f}<br>"
                    "Leads  %{customdata[2]:,.0f}<br>"
                    "Conversion  %{customdata[4]:.1f}%<br>"
                    "Project Value  ₹%{customdata[5]:,.0f}"
                    "<extra></extra>"
                ),
            )
        )

        consultant_fig.update_layout(
            height=consultant_chart_height,
            barmode="group",
            bargap=.48,
            bargroupgap=.04,
            margin=dict(
                l=24,
                r=24,
                t=30,
                b=68
                if consultant_count > 7
                else 42,
            ),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(
                family="Inter",
                color="#334155",
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.005,
                xanchor="left",
                x=0,
                font=dict(size=10),
                bgcolor="rgba(255,255,255,0)",
            ),
            xaxis=dict(
                title=None,
                showgrid=False,
                tickfont=dict(
                    size=10,
                    color="#334155",
                ),
                tickangle=(
                    -24
                    if consultant_count > 7
                    else 0
                ),
                automargin=True,
            ),
            yaxis=dict(
                title=None,
                showgrid=True,
                gridcolor="rgba(148,163,184,.12)",
                zeroline=False,
                tickfont=dict(size=9),
                rangemode="tozero",
            ),
            hoverlabel=dict(
                bgcolor="#ffffff",
                bordercolor="#CBD5E1",
                font_size=12,
                font_family="Inter",
                font_color="#0F172A",
            ),
            clickmode="event+select",
        )

        # Primary drill-down: click either bar.
        consultant_event = st.plotly_chart(
            consultant_fig,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "responsive": True,
                "staticPlot": False,
            },
            key="consultant_vibe_click_chart",
            on_select="rerun",
            selection_mode="points",
        )

        selected_consultant_name = None

        if consultant_event is not None:
            try:
                selected_points = (
                    consultant_event.selection.points
                )
            except Exception:
                selected_points = []

            if selected_points:
                custom = selected_points[0].get(
                    "customdata",
                    None,
                )
                if custom is not None and len(custom):
                    selected_consultant_name = str(
                        custom[0]
                    ).strip()

        st.markdown(
            """
            <div class="consultant-click-note">
                Click a Leads or Projects bar to open the consultant dossier ↗
            </div>
            """,
            unsafe_allow_html=True,
        )

        # --------------------------------------------------------
        # CLICK POPUP — CLEAN CONSULTANT DOSSIER
        # --------------------------------------------------------
        if selected_consultant_name:

            @st.dialog(
                "🏆 Consultant Dossier",
                width="large",
            )
            def show_consultant_vibe_popup():

                consultant_rows = (
                    consultant_source.loc[
                        consultant_source[
                            "Owner Name"
                        ].eq(
                            selected_consultant_name
                        )
                    ]
                    .copy()
                )

                consultant_rows = (
                    consultant_rows
                    .sort_values(
                        [
                            "Lead Number",
                            "Lead Created At",
                        ],
                        ascending=[
                            True,
                            False,
                        ],
                        na_position="last",
                    )
                    .drop_duplicates(
                        subset=["Lead Number"],
                        keep="first",
                    )
                    .copy()
                )

                consultant_projects = (
                    consultant_rows.loc[
                        consultant_rows["_is_project"].astype(bool)
                        & consultant_rows["_period_project"].eq(1)
                    ]
                    .copy()
                )

                lead_count = int(
                    consultant_rows.loc[
                        consultant_rows["_period_lead"].eq(1), "Lead Number"
                    ].nunique()
                )

                project_count = int(
                    consultant_projects[
                        "Lead Number"
                    ].nunique()
                )

                pending_count = max(
                    lead_count - project_count,
                    0,
                )

                conversion = (
                    project_count
                    / lead_count
                    * 100
                    if lead_count
                    else 0.0
                )

                project_value = float(
                    pd.to_numeric(
                        consultant_projects[
                            "Project Value"
                        ],
                        errors="coerce",
                    )
                    .fillna(0)
                    .sum()
                )

                disbursed_value = float(
                    pd.to_numeric(
                        consultant_rows[
                            "Disbursed Value"
                        ],
                        errors="coerce",
                    )
                    .fillna(0)
                    .sum()
                )

                rank_rows = (
                    consultant_summary.loc[
                        consultant_summary[
                            "Owner Name"
                        ].eq(
                            selected_consultant_name
                        )
                    ]
                    .index
                    .tolist()
                )

                consultant_rank = (
                    rank_rows[0] + 1
                    if rank_rows
                    else None
                )

                rank_badge = (
                    rank_symbol(
                        consultant_rank
                    )
                    if consultant_rank
                    else "⚡"
                )

                st.markdown(
                    f"""
                    <div style="
                        border:1px solid rgba(15,39,64,.075);
                        border-radius:16px;
                        padding:.72rem .90rem;
                        background:
                            radial-gradient(circle at 95% 5%,rgba(25,155,121,.10),transparent 30%),
                            linear-gradient(135deg,#f9fcff,#ffffff);
                        margin-bottom:.55rem;
                    ">
                        <div style="
                            color:#0F766E;
                            font-size:.64rem;
                            font-weight:900;
                            letter-spacing:.08em;
                        ">
                            {rank_badge} • PROJECT LEADERBOARD
                        </div>
                        <div style="
                            color:#102f46;
                            font-size:1.20rem;
                            font-weight:900;
                            letter-spacing:-.025em;
                            margin-top:.10rem;
                        ">
                            {selected_consultant_name}
                        </div>
                        <div style="
                            color:#64748B;
                            font-size:.66rem;
                            margin-top:.06rem;
                        ">
                            {selected_consultant_region if region_is_selected else "All Regions"}
                            • current dashboard filters applied
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                m1, m2, m3, m4, m5, m6 = st.columns(6)

                m1.metric(
                    "Rank",
                    (
                        f"#{consultant_rank}"
                        if consultant_rank
                        else "—"
                    ),
                )
                m2.metric(
                    "Leads",
                    f"{lead_count:,}",
                )
                m3.metric(
                    "Projects",
                    f"{project_count:,}",
                )
                m4.metric(
                    "Pending",
                    f"{pending_count:,}",
                )
                m5.metric(
                    "Conversion",
                    f"{conversion:.1f}%",
                )
                m6.metric(
                    "Project Value",
                    money_short(project_value),
                )

                st.caption(
                    "Total Disbursed "
                    + money_short(
                        disbursed_value
                    )
                )

                # Two compact tabs make the popup cleaner.
                portfolio_tab, stage_tab = st.tabs(
                    [
                        "Lead Portfolio",
                        "Stage Mix",
                    ]
                )

                with portfolio_tab:
                    popup_cols = [
                        c
                        for c in [
                            "Lead Number",
                            "Customer Name",
                            "Vendor / CP Name",
                            "Type",
                            "Customer Region",
                            "Vendor Region",
                            "Stage",
                            "Lead Sub Stage",
                            "NBFC",
                            "Lead Created At",
                            "Disbursed At",
                            "Project Value",
                            "Dynamic Pricing",
                            "Disbursed Value",
                            "_Project Value Source",
                        ]
                        if c in consultant_rows.columns
                    ]

                    consultant_view = (
                        consultant_rows[
                            popup_cols
                        ]
                        .copy()
                    )

                    consultant_view = _excel_filter_table(
                        consultant_view,
                        key_prefix="consultant_" + consultant_key(selected_consultant_name),
                        height=430,
                    )

                    st.download_button(
                        "Download Consultant Portfolio",
                        data=consultant_view.to_csv(
                            index=False
                        ).encode(
                            "utf-8-sig"
                        ),
                        file_name=(
                            f"Consultant_"
                            f"{consultant_key(selected_consultant_name)}"
                            f"_Portfolio.csv"
                        ),
                        mime="text/csv",
                        use_container_width=True,
                        key=(
                            "download_consultant_vibe_"
                            + consultant_key(
                                selected_consultant_name
                            )
                        ),
                    )

                with stage_tab:
                    stage_breakdown = (
                        consultant_rows
                        .groupby(
                            "Stage",
                            observed=True,
                            dropna=False,
                        )
                        .agg(
                            Leads=(
                                "Lead Number",
                                "nunique",
                            ),
                            Project_Value=(
                                "Project Value",
                                "sum",
                            ),
                        )
                        .reset_index()
                        .sort_values(
                            "Leads",
                            ascending=False,
                        )
                    )

                    if stage_breakdown.empty:
                        st.info(
                            "No stage data is available."
                        )
                    else:
                        stage_fig = go.Figure(
                            go.Bar(
                                x=stage_breakdown[
                                    "Stage"
                                ],
                                y=stage_breakdown[
                                    "Leads"
                                ],
                                text=stage_breakdown[
                                    "Leads"
                                ],
                                textposition="outside",
                                marker=dict(
                                    color="#0F766E",
                                ),
                                hovertemplate=(
                                    "<b>%{x}</b><br>"
                                    "Leads: %{y:,.0f}"
                                    "<extra></extra>"
                                ),
                            )
                        )

                        stage_fig.update_layout(
                            height=245,
                            margin=dict(
                                l=10,
                                r=10,
                                t=20,
                                b=50,
                            ),
                            paper_bgcolor="rgba(0,0,0,0)",
                            plot_bgcolor="rgba(0,0,0,0)",
                            showlegend=False,
                            xaxis=dict(
                                title=None,
                                showgrid=False,
                                tickfont=dict(
                                    size=9
                                ),
                            ),
                            yaxis=dict(
                                title=None,
                                showgrid=True,
                                gridcolor="rgba(148,163,184,.13)",
                                zeroline=False,
                                tickfont=dict(
                                    size=9
                                ),
                            ),
                        )

                        st.plotly_chart(
                            stage_fig,
                            use_container_width=True,
                            config={
                                "displayModeBar": False,
                            },
                            key=(
                                "consultant_stage_vibe_"
                                + consultant_key(
                                    selected_consultant_name
                                )
                            ),
                        )

            show_consultant_vibe_popup()

# ============================================================
# ROW 4 — CP / VENDOR LEADERBOARD
# ============================================================

def partner_lb_key(value):
    text = str(value or "").strip().casefold()
    cleaned = "".join(ch if ch.isalnum() else "_" for ch in text)
    cleaned = "_".join(part for part in cleaned.split("_") if part)
    return cleaned[:70] or "unassigned"


st.markdown(
    """
    <style>
    [class*="st-key-partner_vibe_board"] {
        border:1px solid rgba(15,39,64,.075);
        border-radius:20px;
        padding:.70rem .95rem .50rem .95rem;
        margin-top:.72rem;
        background:linear-gradient(180deg,#ffffff 0%,#f7f9fc 100%);
        border-top:3px solid #B45309;
        box-shadow:0 12px 32px rgba(15,39,64,.055);
    }
    [class*="st-key-partner_vibe_board"] div[data-testid="stVerticalBlock"] {gap:.18rem;}
    [class*="st-key-partner_type_compact"] label,
    [class*="st-key-partner_region_compact"] label {display:none !important;}
    [class*="st-key-partner_type_compact"] div[data-baseweb="select"] > div,
    [class*="st-key-partner_region_compact"] div[data-baseweb="select"] > div {
        min-height:2.05rem !important;
        border-radius:10px !important;
        background:#F8FAFC !important;
        border-color:rgba(15,39,64,.09) !important;
        font-size:.75rem !important;
    }
    .partner-vibe-kicker {
        color:#D97706;font-size:.60rem;font-weight:900;
        letter-spacing:.12em;text-transform:uppercase;
    }
    .partner-vibe-title {
        color:#102f46;font-size:1.05rem;font-weight:900;
        letter-spacing:-.025em;line-height:1.12;
    }
    .partner-vibe-note {color:#64748B;font-size:.65rem;margin-top:.10rem;}
    .partner-view-pill {
        display:inline-flex;border:1px solid rgba(82,104,255,.13);
        background:#FFFBEB;color:#B45309;border-radius:999px;
        padding:.18rem .48rem;font-size:.56rem;font-weight:900;
        letter-spacing:.055em;
    }
    .partner-click-note {
        color:#64748B;font-size:.60rem;text-align:right;margin-top:-.18rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.container(key="partner_vibe_board"):
    partner_source = filtered.copy()

    partner_source["Type"] = (
        partner_source["Type"].fillna("Unknown").astype(str).str.strip()
    )
    partner_source["Vendor / CP Name"] = (
        partner_source["Vendor / CP Name"]
        .fillna("Unmapped / Direct").astype(str).str.strip()
        .replace("", "Unmapped / Direct")
    )

    partner_type_key = (
        partner_source["Type"].astype(str).str.casefold()
        .str.replace(r"[^a-z0-9]+", "", regex=True)
    )

    partner_source["_Partner View"] = np.select(
        [
            partner_type_key.isin(["channelpartner", "channelpartners", "cp", "partner"]),
            partner_type_key.isin(["vendor", "vendors"]),
        ],
        ["Channel Partner", "Vendor"],
        default="Other",
    )

    partner_source = partner_source.loc[
        partner_source["_Partner View"].isin(["Channel Partner", "Vendor"])
    ].copy()

    partner_source = (
        partner_source
        .sort_values(["Lead Number", "Lead Created At"], ascending=[True, False], na_position="last")
        .drop_duplicates(subset=["Lead Number"], keep="first")
        .copy()
    )

    head_left, type_col, region_col = st.columns(
        [0.58, 0.20, 0.22], gap="small", vertical_alignment="center"
    )

    with head_left:
        st.markdown(
            """
            <div class="partner-vibe-kicker">Partner Leaderboard</div>
            <div class="partner-vibe-title">CP & Vendor Champions</div>
            <div class="partner-vibe-note">
                Leads vs converted projects • switch CP / Vendor from the top-right
            </div>
            <div style="
                display:flex;
                align-items:center;
                gap:.85rem;
                margin-top:.34rem;
                font-size:.65rem;
                font-weight:750;
                color:#475569;
            ">
                <span style="display:inline-flex;align-items:center;gap:.28rem;">
                    <span style="width:9px;height:9px;border-radius:2px;background:#94A3B8;display:inline-block;"></span>
                    Leads
                </span>
                <span style="display:inline-flex;align-items:center;gap:.28rem;">
                    <span style="width:9px;height:9px;border-radius:2px;background:#D97706;display:inline-block;"></span>
                    Projects
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with type_col:
        with st.container(key="partner_type_compact"):
            selected_partner_view = st.selectbox(
                "Partner Type",
                ["Channel Partner", "Vendor"],
                key="partner_leaderboard_type",
                label_visibility="collapsed",
            )

    partner_source = partner_source.loc[
        partner_source["_Partner View"].eq(selected_partner_view)
    ].copy()

    region_values = [
        str(x).strip()
        for x in partner_source["Customer Region"].dropna().astype(str).unique()
        if str(x).strip()
    ]

    preferred_regions = [
        "Secunderabad", "Warangal", "Nizamabad", "Rayalaseema",
        "Vijayawada", "Nellore", "Vizag", "Kadapa",
        "Berhampur", "Cuttack", "Unknown",
    ]

    region_options = (
        [r for r in preferred_regions if r in region_values]
        + sorted([r for r in region_values if r not in preferred_regions])
    )

    with region_col:
        with st.container(key="partner_region_compact"):
            selected_partner_region = st.selectbox(
                "Region",
                ["All Regions"] + region_options,
                key="partner_leaderboard_region",
                label_visibility="collapsed",
            )

    region_selected = selected_partner_region != "All Regions"

    if region_selected:
        partner_source = partner_source.loc[
            partner_source["Customer Region"].eq(selected_partner_region)
        ].copy()

    if partner_source.empty:
        st.info(f"No {selected_partner_view} records are available for this selection.")
    else:
        partner_source["_period_project_win"] = (
            partner_source["_is_project"] & partner_source["_period_project"]
        ).astype("int8")
        partner_source["_period_lead"] = partner_source["_period_lead"].astype("int8")
        partner_source["_period_project_value"] = np.where(
            partner_source["_period_project_win"].eq(1),
            partner_source["Project Value"],
            0.0,
        )
        partner_source["_period_disbursed_value"] = np.where(
            partner_source["_period_project_win"].eq(1),
            partner_source["Disbursed Value"],
            0.0,
        )

        partner_summary = (
            partner_source
            .groupby("Vendor / CP Name", observed=True, dropna=False)
            .agg(
                Leads=("_period_lead", "sum"),
                Projects=("_period_project_win", "sum"),
                Project_Value=("_period_project_value", "sum"),
                Disbursed_Value=("_period_disbursed_value", "sum"),
            )
            .reset_index()
        )

        for c in ["Leads", "Projects"]:
            partner_summary[c] = pd.to_numeric(
                partner_summary[c], errors="coerce"
            ).fillna(0).astype(int)

        for c in ["Project_Value", "Disbursed_Value"]:
            partner_summary[c] = pd.to_numeric(
                partner_summary[c], errors="coerce"
            ).fillna(0.0)

        partner_summary["Pending"] = (
            partner_summary["Leads"] - partner_summary["Projects"]
        ).clip(lower=0)

        partner_summary["Conversion %"] = np.where(
            partner_summary["Leads"] > 0,
            partner_summary["Projects"] / partner_summary["Leads"] * 100,
            0.0,
        )

        partner_summary = partner_summary.sort_values(
            ["Projects", "Leads", "Project_Value", "Vendor / CP Name"],
            ascending=[False, False, False, True],
        ).reset_index(drop=True)

        if region_selected:
            display_partners = partner_summary.copy()
            view_label = f"{len(display_partners):,} • {selected_partner_region.upper()}"
        else:
            display_partners = partner_summary.head(5).copy()
            view_label = "TOP 5 • ALL REGIONS"

        display_partners["Rank"] = np.arange(1, len(display_partners) + 1)

        def partner_rank_symbol(rank):
            rank = int(rank)
            if rank == 1: return "👑 #1"
            if rank == 2: return "⭐ #2"
            if rank == 3: return "🥉 #3"
            return f"⚡ #{rank}"

        display_partners["Project Bar Label"] = [
            f"{partner_rank_symbol(r)}<br><b>{int(p):,} P</b>"
            for r, p in zip(display_partners["Rank"], display_partners["Projects"])
        ]

        custom = np.column_stack([
            display_partners["Vendor / CP Name"].astype(str),
            display_partners["Projects"].astype(int),
            display_partners["Leads"].astype(int),
            display_partners["Pending"].astype(int),
            display_partners["Conversion %"].round(1),
            display_partners["Project_Value"].round(0),
            display_partners["Rank"].astype(int),
        ])

        st.markdown(
            f"""
            <div style="display:flex;justify-content:flex-end;
                        margin-top:-.05rem;margin-bottom:-.32rem;
                        position:relative;z-index:2;">
                <span class="partner-view-pill">
                    {selected_partner_view.upper()} • {view_label}
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        count = len(display_partners)
        fig = go.Figure()

        fig.add_trace(go.Bar(
            name="Leads",
            x=display_partners["Vendor / CP Name"],
            y=display_partners["Leads"],
            width=.24, offset=-.13,
            marker=dict(color="#CBD5E1", line=dict(color="white", width=1)),
            text=[f"{int(v):,}" for v in display_partners["Leads"]],
            textposition="outside",
            textfont=dict(size=10, color="#475569"),
            cliponaxis=False,
            customdata=custom,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Rank #%{customdata[6]}<br>"
                "Leads %{customdata[2]:,.0f}<br>"
                "Projects %{customdata[1]:,.0f}<br>"
                "Pending %{customdata[3]:,.0f}<br>"
                "Conversion %{customdata[4]:.1f}%<br>"
                "Project Value ₹%{customdata[5]:,.0f}<extra></extra>"
            ),
        ))

        fig.add_trace(go.Bar(
            name="Projects",
            x=display_partners["Vendor / CP Name"],
            y=display_partners["Projects"],
            width=.24, offset=.13,
            marker=dict(color="#D97706", line=dict(color="white", width=1)),
            text=display_partners["Project Bar Label"],
            textposition="outside",
            textfont=dict(size=10, color="#B45309"),
            cliponaxis=False,
            customdata=custom,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Rank #%{customdata[6]}<br>"
                "Projects %{customdata[1]:,.0f}<br>"
                "Leads %{customdata[2]:,.0f}<br>"
                "Conversion %{customdata[4]:.1f}%<br>"
                "Project Value ₹%{customdata[5]:,.0f}<extra></extra>"
            ),
        ))

        fig.update_layout(
            height=305 if count <= 5 else min(410, 305 + (count - 5) * 5),
            barmode="group", bargap=.48, bargroupgap=.04,
            margin=dict(l=24, r=24, t=14, b=72 if count > 7 else 46),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter", color="#334155"),
            showlegend=False,
            xaxis=dict(
                title=None, showgrid=False,
                tickfont=dict(size=9, color="#334155"),
                tickangle=-24 if count > 7 else 0, automargin=True,
            ),
            yaxis=dict(
                title=None, showgrid=True,
                gridcolor="rgba(148,163,184,.12)",
                zeroline=False, tickfont=dict(size=9), rangemode="tozero",
            ),
            hoverlabel=dict(
                bgcolor="#fff", bordercolor="#CBD5E1",
                font_size=12, font_family="Inter", font_color="#0F172A",
            ),
            clickmode="event+select",
        )

        event = st.plotly_chart(
            fig, use_container_width=True,
            config={"displayModeBar": False, "responsive": True},
            key="partner_vibe_click_chart",
            on_select="rerun", selection_mode="points",
        )

        selected_partner_name = None
        if event is not None:
            try:
                pts = event.selection.points
            except Exception:
                pts = []
            if pts:
                cd = pts[0].get("customdata", None)
                if cd is not None and len(cd):
                    selected_partner_name = str(cd[0]).strip()

        st.markdown(
            '<div class="partner-click-note">Click a bar to open the partner dossier ↗</div>',
            unsafe_allow_html=True,
        )

        if selected_partner_name:
            @st.dialog("🏆 Partner Dossier", width="large")
            def show_partner_vibe_popup():
                rows = partner_source.loc[
                    partner_source["Vendor / CP Name"].eq(selected_partner_name)
                ].copy()

                rows = (
                    rows.sort_values(
                        ["Lead Number", "Lead Created At"],
                        ascending=[True, False], na_position="last"
                    )
                    .drop_duplicates("Lead Number", keep="first")
                    .copy()
                )

                projects = rows.loc[
                    rows["_is_project"] & rows["_period_project"]
                ].copy()
                leads_n = int(
                    rows.loc[rows["_period_lead"].eq(1), "Lead Number"].nunique()
                )
                projects_n = int(projects["Lead Number"].nunique())
                pending_n = max(leads_n - projects_n, 0)
                conversion = projects_n / leads_n * 100 if leads_n else 0.0
                pv = float(pd.to_numeric(
                    projects["Project Value"], errors="coerce"
                ).fillna(0).sum())
                disb = float(pd.to_numeric(
                    rows["Disbursed Value"], errors="coerce"
                ).fillna(0).sum())

                rank_match = partner_summary.index[
                    partner_summary["Vendor / CP Name"].eq(selected_partner_name)
                ].tolist()
                rank = rank_match[0] + 1 if rank_match else None

                st.markdown(
                    f"### {partner_rank_symbol(rank) if rank else '⚡'} "
                    f"{selected_partner_name}"
                )
                st.caption(
                    f"{selected_partner_view} • "
                    f"{selected_partner_region if region_selected else 'All Regions'}"
                )

                m1, m2, m3, m4, m5, m6 = st.columns(6)
                m1.metric("Rank", f"#{rank}" if rank else "—")
                m2.metric("Leads", f"{leads_n:,}")
                m3.metric("Projects", f"{projects_n:,}")
                m4.metric("Pending", f"{pending_n:,}")
                m5.metric("Conversion", f"{conversion:.1f}%")
                m6.metric("Project Value", money_short(pv))
                st.caption("Total Disbursed " + money_short(disb))

                # Stage Mix first — quick visual context.
                st.markdown("**Stage Mix**")

                stages = (
                    rows.groupby("Stage", observed=True, dropna=False)
                    .agg(Leads=("Lead Number", "nunique"))
                    .reset_index()
                    .sort_values("Leads", ascending=False)
                )

                if stages.empty:
                    st.info("No stage data is available.")
                else:
                    sf = go.Figure(
                        go.Bar(
                            x=stages["Stage"],
                            y=stages["Leads"],
                            text=stages["Leads"],
                            textposition="outside",
                            marker=dict(color="#D97706"),
                            hovertemplate=(
                                "<b>%{x}</b><br>"
                                "Leads: %{y:,.0f}"
                                "<extra></extra>"
                            ),
                        )
                    )
                    sf.update_layout(
                        height=220,
                        margin=dict(l=10, r=10, t=12, b=48),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        showlegend=False,
                        xaxis=dict(
                            title=None,
                            showgrid=False,
                            tickfont=dict(size=9),
                        ),
                        yaxis=dict(
                            title=None,
                            showgrid=True,
                            gridcolor="rgba(148,163,184,.13)",
                            zeroline=False,
                            tickfont=dict(size=9),
                        ),
                    )
                    st.plotly_chart(
                        sf,
                        use_container_width=True,
                        config={"displayModeBar": False},
                        key="partner_stage_" + partner_lb_key(selected_partner_name),
                    )

                # Lead details directly below Stage Mix.
                st.markdown("**Lead Details**")

                cols = [
                    c for c in [
                        "Lead Number",
                        "Customer Name",
                        "Owner Name",
                        "Customer Region",
                        "Vendor Region",
                        "Stage",
                        "Lead Sub Stage",
                        "NBFC",
                        "Lead Created At",
                        "Disbursed At",
                        "Project Value",
                        "Dynamic Pricing",
                        "Disbursed Value",
                        "_Project Value Source",
                    ]
                    if c in rows.columns
                ]

                view = rows[cols].copy()

                view = _excel_filter_table(
                    view,
                    key_prefix="partner_" + partner_lb_key(selected_partner_name),
                    height=390,
                )

                st.download_button(
                    "Download Partner Lead Details",
                    data=view.to_csv(index=False).encode("utf-8-sig"),
                    file_name=(
                        f"{partner_lb_key(selected_partner_view)}_"
                        f"{partner_lb_key(selected_partner_name)}_Lead_Details.csv"
                    ),
                    mime="text/csv",
                    use_container_width=True,
                    key="download_partner_" + partner_lb_key(selected_partner_name),
                )

            show_partner_vibe_popup()

