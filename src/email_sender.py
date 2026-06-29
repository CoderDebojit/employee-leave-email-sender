"""
email_sender.py
───────────────
Two sending backends, same interface:
  • OutlookSender  — uses system's default mail client via mailto: (New Outlook compatible)
  • SmtpSender     — uses any SMTP server (Gmail, Office 365, custom)

Author : Debojit Dhali
License: MIT
"""

from __future__ import annotations

import smtplib
import ssl
import urllib.parse
import webbrowser
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Callable, List, Optional, Tuple

from src.excel_reader import Employee, LeaveRecord


# ── Config dataclass passed from Settings tab ─────────────────────────────────

@dataclass
class SmtpConfig:
    host:     str
    port:     int
    username: str
    password: str
    use_tls:  bool = True


# ── HTML email builder ────────────────────────────────────────────────────────

def build_html(emp: Employee, leave: LeaveRecord, duration: str) -> str:
    badge_color = "#dc2626" if leave.leave_type.lower() == "unplanned" else "#0ea5e9"

    def row(bg: str, label: str, value: str) -> str:
        return f"""
        <tr style="border-top:1px solid #e2e8f0;">
          <td style="padding:11px 20px;font-size:13px;color:#64748b;
                     font-weight:600;background:{bg};">{label}</td>
          <td style="padding:11px 20px;font-size:13px;color:#1e293b;
                     background:{bg};">{value}</td>
        </tr>"""

    rows = (
        row("#ffffff", "👤 Employee",  emp.name)
      + row("#f1f5f9", "🏢 Team",      emp.team)
      + row("#ffffff", "📅 From",      leave.start_date.strftime("%d %b %Y"))
      + row("#f1f5f9", "📅 To",        leave.end_date.strftime("%d %b %Y"))
      + row("#ffffff", "⏱ Duration",   duration)
      + row("#f1f5f9", "📝 Reason",    leave.reason)
    )

    return f"""
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;background:#f1f5f9;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#f1f5f9;padding:30px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0"
           style="background:#fff;border-radius:12px;overflow:hidden;
                  box-shadow:0 4px 24px rgba(0,0,0,.10);">

      <tr>
        <td style="background:linear-gradient(135deg,#0c4a6e 0%,#0ea5e9 100%);
                   padding:32px 40px;text-align:center;">
          <p  style="margin:0;font-size:28px;">📬</p>
          <h1 style="margin:8px 0 4px;color:#fff;font-size:22px;font-weight:700;">
              Leave Notification</h1>
          <p  style="margin:0;color:#bae6fd;font-size:13px;">
              Automated Team Notification System</p>
        </td>
      </tr>

      <tr>
        <td style="padding:32px 40px 0;">
          <p style="margin:0;color:#1e293b;font-size:15px;">
            Dear <strong>{emp.manager_name}</strong>,</p>
          <p style="margin:12px 0 0;color:#475569;font-size:14px;line-height:1.6;">
            This is an automated notification. The following team member has a leave
            recorded. Please plan accordingly and arrange coverage if required.</p>
        </td>
      </tr>

      <tr>
        <td style="padding:20px 40px 0;">
          <span style="display:inline-block;background:{badge_color};color:#fff;
                       font-size:12px;font-weight:700;padding:4px 14px;
                       border-radius:20px;letter-spacing:.8px;text-transform:uppercase;">
            {leave.leave_type} Leave
          </span>
        </td>
      </tr>

      <tr>
        <td style="padding:16px 40px 0;">
          <table width="100%" cellpadding="0" cellspacing="0"
                 style="background:#f8fafc;border:1px solid #e2e8f0;
                        border-radius:10px;overflow:hidden;">
            <tr style="background:#e2e8f0;">
              <td style="padding:10px 20px;font-size:12px;font-weight:700;
                         color:#64748b;text-transform:uppercase;letter-spacing:.6px;"
                  width="35%">Field</td>
              <td style="padding:10px 20px;font-size:12px;font-weight:700;
                         color:#64748b;text-transform:uppercase;letter-spacing:.6px;">
                  Details</td>
            </tr>
            {rows}
          </table>
        </td>
      </tr>

      <tr>
        <td style="padding:24px 40px 0;">
          <p style="margin:0;color:#475569;font-size:13px;">
            Contact employee:
            <a href="mailto:{emp.email}"
               style="color:#0ea5e9;font-weight:600;">{emp.email}</a>
          </p>
        </td>
      </tr>

      <tr>
        <td style="padding:28px 40px 32px;">
          <hr style="border:none;border-top:1px solid #e2e8f0;margin:0 0 16px;">
          <p style="margin:0;color:#94a3b8;font-size:12px;text-align:center;">
            🤖 Automated message · Leave Management System · Do not reply
          </p>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>"""


# ── Abstract base ─────────────────────────────────────────────────────────────

class BaseEmailSender(ABC):
    @abstractmethod
    def send(
        self,
        employees:    List[Employee],
        leaves:       List[LeaveRecord],
        log:          Callable[[str], None],
        display_only: bool = True,
    ) -> Tuple[int, List[str]]:
        ...

    def _filter_today(self, leaves: List[LeaveRecord]) -> List[LeaveRecord]:
        today = date.today()
        return [l for l in leaves if l.start_date <= today <= l.end_date]

    def _duration(self, leave: LeaveRecord) -> str:
        days = (leave.end_date - leave.start_date).days + 1
        return f"{days} day(s)" if days > 1 else "1 day (today)"


# ── Outlook COM backend ───────────────────────────────────────────────────────

class OutlookSender(BaseEmailSender):
    """
    Sends via the system's default mail client (New Outlook, Mail app, browser).
    Opens a draft window for the user to review and send manually.
    """

    def send(self, employees, leaves, log, display_only=True):
        on_leave = self._filter_today(leaves)
        if not on_leave:
            log("✅  No leaves found for today.")
            return 0, []

        log(f"👥  {len(on_leave)} person(s) on leave today.")
        emp_map = {e.name: e for e in employees}
        errors, sent = [], 0

        for leave in on_leave:
            emp = emp_map.get(leave.employee_name)
            if not emp:
                msg = f"⚠️  '{leave.employee_name}' not in Employee Master — skipped."
                log(msg); errors.append(msg); continue

            if not emp.manager_email:
                msg = f"⚠️  No Manager Email for '{emp.name}' — skipped."
                log(msg); errors.append(msg); continue

            duration = self._duration(leave)
            subject  = (f"[Leave Alert] {emp.name} | "
                        f"{leave.leave_type} Leave | "
                        f"{leave.start_date.strftime('%d %b %Y')}")

            # Create a clean plain-text body for the draft
            body = (
                f"Dear {emp.manager_name},\n\n"
                f"This is an automated notification. The following team member has a leave recorded. "
                f"Please plan accordingly and arrange coverage if required.\n\n"
                f"----- LEAVE DETAILS -----\n"
                f"Employee   : {emp.name}\n"
                f"Team       : {emp.team}\n"
                f"Leave Type : {leave.leave_type} Leave\n"
                f"From       : {leave.start_date.strftime('%d %b %Y')}\n"
                f"To         : {leave.end_date.strftime('%d %b %Y')}\n"
                f"Duration   : {duration}\n"
                f"Reason     : {leave.reason}\n"
                f"-------------------------\n\n"
                f"Contact employee: {emp.email}\n"
            )

            # Build the mailto link
            query_params = {'subject': subject, 'body': body}
            if emp.cc_emails:
                query_params['cc'] = "; ".join(emp.cc_emails)

            # Safely encode the URL so the email client reads the spaces and line breaks correctly
            query_string = urllib.parse.urlencode(query_params, quote_via=urllib.parse.quote)
            mailto_url = f"mailto:{emp.manager_email}?{query_string}"

            # Open the draft in the default mail client (New Outlook)
            webbrowser.open(mailto_url)

            sent += 1
            cc_info = f" | CC: {', '.join(emp.cc_emails)}" if emp.cc_emails else ""
            log(f"✅  [{emp.name}] Draft opened for → {emp.manager_email}{cc_info}")

        return sent, errors


# ── SMTP backend ──────────────────────────────────────────────────────────────

class SmtpSender(BaseEmailSender):
    """
    Sends via any SMTP server.
    Works with Gmail (App Password), Office 365, corporate SMTP, etc.
    Cross-platform — no Outlook installation needed.
    """

    def __init__(self, config: SmtpConfig):
        self.config = config

    def send(self, employees, leaves, log, display_only=False):
        on_leave = self._filter_today(leaves)
        if not on_leave:
            log("✅  No leaves found for today.")
            return 0, []

        log(f"👥  {len(on_leave)} person(s) on leave today.")
        log(f"🔌  Connecting to {self.config.host}:{self.config.port} …")

        emp_map = {e.name: e for e in employees}
        errors, sent = [], 0

        try:
            context = ssl.create_default_context()
            if self.config.use_tls:
                server = smtplib.SMTP(self.config.host, self.config.port, timeout=15)
                server.ehlo()
                server.starttls(context=context)
            else:
                server = smtplib.SMTP_SSL(self.config.host, self.config.port,
                                          context=context, timeout=15)
            server.login(self.config.username, self.config.password)
            log("🔐  SMTP login successful.")
        except Exception as e:
            raise RuntimeError(f"SMTP connection failed: {e}")

        try:
            for leave in on_leave:
                emp = emp_map.get(leave.employee_name)
                if not emp:
                    msg = f"⚠️  '{leave.employee_name}' not in Employee Master — skipped."
                    log(msg); errors.append(msg); continue

                if not emp.manager_email:
                    msg = f"⚠️  No Manager Email for '{emp.name}' — skipped."
                    log(msg); errors.append(msg); continue

                duration = self._duration(leave)
                html     = build_html(emp, leave, duration)
                subject  = (f"[Leave Alert] {emp.name} | "
                            f"{leave.leave_type} Leave | "
                            f"{leave.start_date.strftime('%d %b %Y')}")

                msg = MIMEMultipart("alternative")
                msg["Subject"] = subject
                msg["From"]    = self.config.username
                msg["To"]      = emp.manager_email
                if emp.cc_emails:
                    msg["CC"]  = ", ".join(emp.cc_emails)

                msg.attach(MIMEText(html, "html"))

                all_recipients = [emp.manager_email] + emp.cc_emails
                server.sendmail(self.config.username, all_recipients, msg.as_string())

                sent += 1
                cc_info = f" | CC: {', '.join(emp.cc_emails)}" if emp.cc_emails else ""
                log(f"✅  [{emp.name}]  TO → {emp.manager_email}{cc_info}")

        finally:
            try:
                server.quit()
            except Exception:
                pass

        return sent, errors