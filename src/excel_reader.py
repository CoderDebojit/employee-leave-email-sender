"""
excel_reader.py
───────────────
Reads Employee Master and Leave Tracker sheets from the shared Excel file.
Returns typed dataclass objects — no raw DataFrames leak outside this module.

Author : Debojit Dhali
License: MIT
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from typing import List, Optional, Tuple

import pandas as pd


# ── Data models ───────────────────────────────────────────────────────────────

@dataclass
class Employee:
    name:          str
    email:         str
    team:          str
    manager_name:  str
    manager_email: str
    cc_emails:     List[str] = field(default_factory=list)


@dataclass
class LeaveRecord:
    employee_name: str
    leave_type:    str
    start_date:    date
    end_date:      date
    reason:        str


# ── Helpers ───────────────────────────────────────────────────────────────────

_EMAIL_RE = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")


def _is_valid_email(val: str) -> bool:
    return bool(_EMAIL_RE.match(val.strip()))


def _clean(val) -> str:
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "") else s


def _parse_date(val) -> Optional[date]:
    try:
        return pd.to_datetime(val, dayfirst=True, errors="raise").date()
    except Exception:
        return None


# ── Public API ────────────────────────────────────────────────────────────────

def read_excel(path: str) -> Tuple[List[Employee], List[LeaveRecord], Optional[str]]:
    """
    Read *path* and return (employees, leave_records, error_string).
    error_string is None on success.
    """
    try:
        xl = pd.ExcelFile(path)
    except FileNotFoundError:
        return [], [], f"File not found: {path}"
    except Exception as e:
        return [], [], f"Cannot open file: {e}"

    missing = [s for s in ("Employee Master", "Leave Tracker") if s not in xl.sheet_names]
    if missing:
        return [], [], f"Missing sheet(s): {', '.join(missing)}"

    # ── Employee Master ───────────────────────────────────────────────────────
    emp_df = pd.read_excel(path, sheet_name="Employee Master", dtype=str)
    emp_df.columns = [c.strip() for c in emp_df.columns]

    required_emp = {"Name", "Email", "Team", "Manager Name", "Manager Email"}
    missing_cols = required_emp - set(emp_df.columns)
    if missing_cols:
        return [], [], f"Employee Master missing columns: {', '.join(missing_cols)}"

    cc_cols = [c for c in emp_df.columns if c.upper().startswith("CC")]

    employees: List[Employee] = []
    for _, row in emp_df.iterrows():
        name = _clean(row.get("Name", ""))
        if not name:
            continue
        cc = [
            _clean(row[c]) for c in cc_cols
            if _is_valid_email(_clean(row.get(c, "")))
        ]
        employees.append(Employee(
            name          = name,
            email         = _clean(row.get("Email", "")),
            team          = _clean(row.get("Team", "")),
            manager_name  = _clean(row.get("Manager Name", "")) or "Manager",
            manager_email = _clean(row.get("Manager Email", "")),
            cc_emails     = cc,
        ))

    # ── Leave Tracker ─────────────────────────────────────────────────────────
    leave_df = pd.read_excel(path, sheet_name="Leave Tracker", dtype=str)
    leave_df.columns = [c.strip() for c in leave_df.columns]

    required_leave = {"Employee Name", "Leave Type", "Start Date", "End Date", "Reason"}
    missing_cols = required_leave - set(leave_df.columns)
    if missing_cols:
        return [], [], f"Leave Tracker missing columns: {', '.join(missing_cols)}"

    leaves: List[LeaveRecord] = []
    for _, row in leave_df.iterrows():
        emp_name = _clean(row.get("Employee Name", ""))
        if not emp_name:
            continue
        start = _parse_date(row.get("Start Date"))
        end   = _parse_date(row.get("End Date"))
        if not start or not end:
            continue
        leaves.append(LeaveRecord(
            employee_name = emp_name,
            leave_type    = _clean(row.get("Leave Type", "")) or "Leave",
            start_date    = start,
            end_date      = end,
            reason        = _clean(row.get("Reason", "")) or "Not specified",
        ))

    return employees, leaves, None
