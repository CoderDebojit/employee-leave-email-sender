"""
leave_mailer.py
───────────────
Entry point. Tkinter GUI with three tabs:
  • Send Emails  — load Excel, pick backend, fire emails
  • Settings     — configure SMTP or use Outlook COM
  • Help         — usage guide

Author  : Debojit Dhali
GitHub  : https://github.com/CoderDebojit
License : MIT © 2025 Debojit Dhali
"""

from __future__ import annotations

import os
import sys
import json
import threading
from datetime import datetime
from tkinter import (
    BooleanVar, Frame, Label, StringVar, Text, Tk,
    filedialog, messagebox, scrolledtext, ttk,
)
import tkinter as tk

from excel_reader import read_excel
from email_sender import OutlookSender, SmtpSender, SmtpConfig
from template_generator import generate


# ── Resolve base directory (works for both .py and PyInstaller .exe) ──────────

def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ── Settings file (persists SMTP config between sessions) ─────────────────────

_SETTINGS_PATH = os.path.join(_base_dir(), "lm_settings.json")

_DEFAULT_SETTINGS = {
    "backend":   "outlook",   # "outlook" | "smtp"
    "smtp_host": "smtp.gmail.com",
    "smtp_port": "587",
    "smtp_user": "",
    "smtp_pass": "",
    "smtp_tls":  True,
    "preview":   True,
    "last_file": "",
}


def _load_settings() -> dict:
    try:
        with open(_SETTINGS_PATH) as f:
            data = json.load(f)
        return {**_DEFAULT_SETTINGS, **data}
    except Exception:
        return dict(_DEFAULT_SETTINGS)


def _save_settings(cfg: dict) -> None:
    try:
        with open(_SETTINGS_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ── Colour tokens ──────────────────────────────────────────────────────────────

C = {
    "bg":       "#0d1117",
    "surface":  "#161b22",
    "border":   "#30363d",
    "accent":   "#2f81f7",
    "success":  "#238636",
    "danger":   "#da3633",
    "text":     "#e6edf3",
    "muted":    "#8b949e",
    "hdr":      "#010409",
    "entry_bg": "#0d1117",
    "log_bg":   "#010409",
    "log_fg":   "#3fb950",
    "btn_dark": "#21262d",
    "btn_dh":   "#30363d",
}


# ── Reusable widget factory ────────────────────────────────────────────────────

class W:
    """Thin factory so we stop repeating style kwargs everywhere."""

    @staticmethod
    def label(parent, text, size=9, bold=False, fg=None, bg=None, **kw):
        return Label(
            parent, text=text,
            font=("Segoe UI", size, "bold" if bold else "normal"),
            fg=fg or C["text"], bg=bg or C["bg"], **kw,
        )

    @staticmethod
    def entry(parent, textvariable=None, show=None, width=40, **kw):
        e = tk.Entry(
            parent, textvariable=textvariable, show=show, width=width,
            font=("Consolas", 10),
            bg=C["entry_bg"], fg=C["text"],
            insertbackground=C["text"],
            highlightbackground=C["border"],
            highlightthickness=1,
            relief="flat", bd=6,
            **kw,
        )
        return e

    @staticmethod
    def button(parent, text, bg, hover, command,
               bold=False, fg=None, **kw):
        btn = tk.Button(
            parent, text=text,
            font=("Segoe UI", 9, "bold" if bold else "normal"),
            bg=bg, fg=fg or C["text"],
            activebackground=hover, activeforeground=fg or C["text"],
            relief="flat", cursor="hand2", bd=0,
            command=command, **kw,
        )
        btn.bind("<Enter>", lambda _: btn.configure(bg=hover))
        btn.bind("<Leave>", lambda _: btn.configure(bg=bg))
        return btn

    @staticmethod
    def separator(parent):
        return Frame(parent, bg=C["border"], height=1)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════

class LeaveMailerApp:

    def __init__(self, root: Tk):
        self.root = root
        self.cfg  = _load_settings()
        self._configure_root()
        self._apply_notebook_style()
        self._build_header()
        self._build_notebook()
        self._build_status_bar()

    # ── Window setup ──────────────────────────────────────────────────────────

    def _configure_root(self):
        self.root.title("Leave Mailer  v2.0")
        self.root.geometry("820x700")
        self.root.minsize(820, 700)
        self.root.configure(bg=C["bg"])
        self.root.resizable(True, True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_notebook_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",
                        background=C["bg"], borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=C["surface"], foreground=C["muted"],
                        font=("Segoe UI", 9), padding=[14, 6],
                        borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", C["bg"])],
                  foreground=[("selected", C["text"])])

    # ── Header ────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = Frame(self.root, bg=C["hdr"])
        hdr.pack(fill="x")

        inner = Frame(hdr, bg=C["hdr"], pady=18, padx=28)
        inner.pack(fill="x")

        # Left: icon + title
        left = Frame(inner, bg=C["hdr"])
        left.pack(side="left")

        W.label(left, "📬", 26, bold=True, fg=C["accent"],
                bg=C["hdr"]).pack(side="left", padx=(0, 12))

        text_block = Frame(left, bg=C["hdr"])
        text_block.pack(side="left")
        W.label(text_block, "Leave Mailer", 18, bold=True, bg=C["hdr"]).pack(anchor="w")
        W.label(text_block, "Automated Team Notification System",
                9, fg=C["muted"], bg=C["hdr"]).pack(anchor="w")

        # Right: version badge
        W.label(inner, " v2.0 ", 8, bold=True,
                fg="#58a6ff", bg="#1f3a5f").pack(side="right", anchor="n", pady=4)

        W.separator(self.root).pack(fill="x")

    # ── Notebook ──────────────────────────────────────────────────────────────

    def _build_notebook(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        self._tab_send(nb)
        self._tab_settings(nb)
        self._tab_help(nb)

    # ═══════════════════════════════════
    #  TAB 1 — Send Emails
    # ═══════════════════════════════════

    def _tab_send(self, nb: ttk.Notebook):
        tab = Frame(nb, bg=C["bg"])
        nb.add(tab, text="  📨  Send Emails  ")

        body = Frame(tab, bg=C["bg"], padx=28, pady=20)
        body.pack(fill="both", expand=True)

        # ── Section ①: File ───────────────────────────────────────────────────
        self._section(body, "① Excel File")

        W.label(body, "Path to your shared Leave Tracker Excel file:",
                fg=C["muted"]).pack(anchor="w", pady=(0, 5))

        row_file = Frame(body, bg=C["bg"])
        row_file.pack(fill="x", pady=(0, 4))

        self._file_var = StringVar(value=self.cfg.get("last_file", ""))
        file_entry = W.entry(row_file, textvariable=self._file_var)
        file_entry.pack(side="left", fill="x", expand=True, ipady=7)

        W.button(row_file, "  📁  Browse", C["btn_dark"], C["btn_dh"],
                 self._browse, padx=10).pack(side="left", padx=(8, 0), ipady=7)

        # ── Info box ──────────────────────────────────────────────────────────
        info = Frame(body, bg="#1c2128",
                     highlightbackground=C["border"], highlightthickness=1)
        info.pack(fill="x", pady=(4, 18))
        info_inner = Frame(info, bg="#1c2128", padx=14, pady=10)
        info_inner.pack(fill="x")
        Label(info_inner,
              text=("ℹ️  Excel must have two sheets:\n"
                    "    • Employee Master  →  Name | Email | Team | "
                    "Manager Name | Manager Email | CC1 | CC2 ...\n"
                    "    • Leave Tracker    →  Employee Name | Leave Type | "
                    "Start Date (DD-MM-YYYY) | End Date | Reason"),
              font=("Consolas", 8), fg="#8b949e", bg="#1c2128",
              justify="left").pack(anchor="w")

        # ── Section ②: Actions ────────────────────────────────────────────────
        self._section(body, "② Actions")

        row_actions = Frame(body, bg=C["bg"])
        row_actions.pack(fill="x", pady=(0, 18))

        W.button(row_actions, "  📄  Generate Template",
                 C["btn_dark"], C["btn_dh"],
                 self._on_generate,
                 padx=10).pack(side="left", ipady=7)

        self._preview_var = BooleanVar(value=self.cfg.get("preview", True))
        preview_cb = tk.Checkbutton(
            row_actions,
            text=" Preview before send",
            variable=self._preview_var,
            font=("Segoe UI", 9),
            fg=C["muted"], bg=C["bg"],
            activeforeground=C["text"], activebackground=C["bg"],
            selectcolor=C["surface"],
        )
        preview_cb.pack(side="left", padx=(16, 0))

        self._send_btn = W.button(
            row_actions, "  📨  Send Emails Today",
            C["accent"], "#1a6ee8",
            self._on_send, bold=True, fg="white", padx=10,
        )
        self._send_btn.pack(side="right", ipady=7)

        # ── Section ③: Log ────────────────────────────────────────────────────
        log_hdr = Frame(body, bg=C["bg"])
        log_hdr.pack(fill="x", pady=(0, 4))
        self._section(log_hdr, "③ Activity Log", pack=False).pack(side="left")
        W.button(log_hdr, "  Clear  ", C["btn_dark"], C["btn_dh"],
                 self._clear_log).pack(side="right", ipady=4)

        self._log_box = scrolledtext.ScrolledText(
            body, state="disabled",
            font=("Consolas", 9),
            bg=C["log_bg"], fg=C["log_fg"],
            insertbackground=C["text"],
            highlightbackground=C["border"], highlightthickness=1,
            relief="flat", bd=0,
        )
        self._log_box.pack(fill="both", expand=True)

    # ═══════════════════════════════════
    #  TAB 2 — Settings
    # ═══════════════════════════════════

    def _tab_settings(self, nb: ttk.Notebook):
        tab = Frame(nb, bg=C["bg"])
        nb.add(tab, text="  ⚙️  Settings  ")

        body = Frame(tab, bg=C["bg"], padx=28, pady=20)
        body.pack(fill="both", expand=True)

        # ── Backend selector ──────────────────────────────────────────────────
        self._section(body, "Email Backend")

        W.label(body,
                "Choose how emails are sent. Outlook uses your desktop app "
                "(Windows only, no credentials).\nSMTP works anywhere — "
                "needs your email + app password.",
                fg=C["muted"]).pack(anchor="w", pady=(0, 12))

        self._backend_var = StringVar(value=self.cfg.get("backend", "outlook"))

        for val, lbl, desc in [
            ("outlook", "📮  Outlook (COM)",
             "Uses locally installed Outlook. No password needed. Windows only."),
            ("smtp",    "📡  SMTP / Normal Email",
             "Gmail, Office 365, or any SMTP server. Works on all platforms."),
        ]:
            row = Frame(body, bg=C["surface"],
                        highlightbackground=C["border"], highlightthickness=1)
            row.pack(fill="x", pady=4)
            inner = Frame(row, bg=C["surface"], padx=14, pady=10)
            inner.pack(fill="x")

            tk.Radiobutton(
                inner, text=lbl, value=val, variable=self._backend_var,
                font=("Segoe UI", 10, "bold"),
                fg=C["text"], bg=C["surface"],
                activeforeground=C["text"], activebackground=C["surface"],
                selectcolor=C["bg"],
                command=self._toggle_smtp_fields,
            ).pack(anchor="w")
            W.label(inner, desc, fg=C["muted"], bg=C["surface"]).pack(anchor="w", padx=22)

        # ── SMTP fields ───────────────────────────────────────────────────────
        self._section(body, "SMTP Configuration")

        self._smtp_frame = Frame(body, bg=C["bg"])
        self._smtp_frame.pack(fill="x", pady=(0, 12))

        self._smtp_host = StringVar(value=self.cfg.get("smtp_host", "smtp.gmail.com"))
        self._smtp_port = StringVar(value=str(self.cfg.get("smtp_port", "587")))
        self._smtp_user = StringVar(value=self.cfg.get("smtp_user", ""))
        self._smtp_pass = StringVar(value=self.cfg.get("smtp_pass", ""))
        self._smtp_tls  = BooleanVar(value=self.cfg.get("smtp_tls", True))

        fields = [
            ("SMTP Host",   self._smtp_host, False),
            ("SMTP Port",   self._smtp_port, False),
            ("Your Email",  self._smtp_user, False),
            ("App Password",self._smtp_pass, True ),
        ]

        for lbl_text, var, is_pass in fields:
            row = Frame(self._smtp_frame, bg=C["bg"])
            row.pack(fill="x", pady=3)
            W.label(row, f"{lbl_text}:", fg=C["muted"],
                    width=14, anchor="w").pack(side="left")
            W.entry(row, textvariable=var,
                    show="•" if is_pass else None,
                    width=36).pack(side="left", ipady=5)

        tls_row = Frame(self._smtp_frame, bg=C["bg"])
        tls_row.pack(fill="x", pady=3)
        W.label(tls_row, "Use STARTTLS:", fg=C["muted"],
                width=14, anchor="w").pack(side="left")
        tk.Checkbutton(
            tls_row, variable=self._smtp_tls,
            fg=C["muted"], bg=C["bg"],
            activebackground=C["bg"], selectcolor=C["surface"],
        ).pack(side="left")

        # ── Gmail tip box ─────────────────────────────────────────────────────
        tip = Frame(body, bg="#1c2128",
                    highlightbackground=C["border"], highlightthickness=1)
        tip.pack(fill="x", pady=(4, 16))
        tip_inner = Frame(tip, bg="#1c2128", padx=14, pady=10)
        tip_inner.pack(fill="x")
        Label(tip_inner,
              text=("💡 Gmail users: enable 2FA → go to Google Account → Security →\n"
                    "   App Passwords → generate password for 'Mail'. "
                    "Use that as App Password above.\n\n"
                    "   Office 365: host = smtp.office365.com | port = 587 | TLS = ON"),
              font=("Segoe UI", 8), fg="#8b949e", bg="#1c2128",
              justify="left").pack(anchor="w")

        # ── Save button ───────────────────────────────────────────────────────
        W.button(body, "  💾  Save Settings", C["success"],
                 "#196127", self._save_settings_ui,
                 bold=True, fg="white", padx=10).pack(anchor="e", ipady=7)

        self._toggle_smtp_fields()

    # ═══════════════════════════════════
    #  TAB 3 — Help
    # ═══════════════════════════════════

    def _tab_help(self, nb: ttk.Notebook):
        tab = Frame(nb, bg=C["bg"])
        nb.add(tab, text="  ❓  Help  ")

        body = Frame(tab, bg=C["bg"], padx=28, pady=20)
        body.pack(fill="both", expand=True)

        help_text = """HOW TO USE
──────────────────────────────────────────────────────────────────

STEP 1 — Generate Template
  Click "Generate Template" on the Send Emails tab.
  A file Leave_Tracker_Template.xlsx is created in the same
  folder as this application.
  Upload it to your shared network drive.

STEP 2 — Fill Employee Master sheet
  Add each team member row:
  • Name        → Full name (must match Leave Tracker exactly)
  • Email       → Their email address
  • Team        → Team name (Dev / Testing / DB etc.)
  • Manager Name / Manager Email → Their direct manager
  • CC1, CC2 … → Scrum Master, PM, any others to always CC

STEP 3 — Fill Leave Tracker sheet (daily)
  When someone goes on leave, add a row:
  • Employee Name → must match Employee Master exactly
  • Leave Type    → Planned / Unplanned
  • Start Date    → DD-MM-YYYY
  • End Date      → DD-MM-YYYY
  • Reason        → Short description

STEP 4 — Send Emails
  Go to Settings → pick Outlook or SMTP → save.
  Come back to Send Emails tab → browse to Excel → click Send.

──────────────────────────────────────────────────────────────────

OUTLOOK MODE
  • Outlook desktop app must be installed and logged in.
  • No password required — uses your active Outlook session.
  • Windows only.
  • "Preview" checkbox = opens email draft for review before send.

SMTP MODE
  • Works on any OS and without Outlook.
  • Gmail: use an App Password (not your main password).
  • Office 365: host smtp.office365.com, port 587, TLS ON.
  • Emails send directly — no preview window.

──────────────────────────────────────────────────────────────────

COMMON ISSUES
  ❌ "Not in Employee Master" → name in Leave Tracker doesn't
     match Employee Master exactly (check spaces / spelling).
  ❌ Outlook not found → Open Outlook first, then retry.
  ❌ SMTP auth error → Wrong password or App Password not set up.
  ❌ Excel locked → Close the file in Excel first, then run.
"""

        txt = Text(body, font=("Consolas", 9),
                   bg=C["log_bg"], fg=C["muted"],
                   relief="flat", bd=0,
                   highlightbackground=C["border"], highlightthickness=1,
                   wrap="word", state="normal")
        txt.insert("1.0", help_text)
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_status_bar(self):
        W.separator(self.root).pack(fill="x")

        bar = Frame(self.root, bg=C["hdr"], pady=5, padx=16)
        bar.pack(fill="x", side="bottom")

        self._status_var = StringVar(value="● Ready")
        Label(bar, textvariable=self._status_var,
              font=("Segoe UI", 8), fg=C["muted"],
              bg=C["hdr"]).pack(side="left")

        # Copyright — your name, always visible
        Label(bar,
              text="© 2025 Debojit Dhali  |  MIT License",
              font=("Segoe UI", 8), fg="#30363d",
              bg=C["hdr"]).pack(side="right")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _section(self, parent, text, pack=True):
        lbl = W.label(parent, text, 10, bold=True)
        if pack:
            lbl.pack(anchor="w", pady=(0, 4))
        return lbl

    def _log(self, msg: str):
        def _do():
            self._log_box.configure(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self._log_box.insert("end", f"[{ts}]  {msg}\n")
            self._log_box.see("end")
            self._log_box.configure(state="disabled")
        self.root.after(0, _do)

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _set_status(self, msg: str):
        self.root.after(0, lambda: self._status_var.set(msg))

    def _toggle_smtp_fields(self):
        state = "normal" if self._backend_var.get() == "smtp" else "disabled"
        for widget in self._smtp_frame.winfo_children():
            for child in widget.winfo_children():
                try:
                    child.configure(state=state)
                except Exception:
                    pass

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select Leave Tracker Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if path:
            self._file_var.set(path)

    # ── Settings persistence ──────────────────────────────────────────────────

    def _save_settings_ui(self):
        self.cfg.update({
            "backend":   self._backend_var.get(),
            "smtp_host": self._smtp_host.get(),
            "smtp_port": self._smtp_port.get(),
            "smtp_user": self._smtp_user.get(),
            "smtp_pass": self._smtp_pass.get(),
            "smtp_tls":  self._smtp_tls.get(),
            "preview":   self._preview_var.get(),
            "last_file": self._file_var.get(),
        })
        _save_settings(self.cfg)
        messagebox.showinfo("Saved", "Settings saved successfully.")

    def _on_close(self):
        # Auto-save last used file path on exit
        self.cfg["last_file"] = self._file_var.get()
        self.cfg["preview"]   = self._preview_var.get()
        _save_settings(self.cfg)
        self.root.destroy()

    # ── Actions ───────────────────────────────────────────────────────────────

    def _on_generate(self):
        self._log("─" * 54)
        self._log("Generating Excel template …")
        self._set_status("Generating template …")

        def _worker():
            ok, path, err = generate(_base_dir())
            if ok:
                self._log(f"✅  Template saved:\n    {path}")
                self._set_status("✅  Template generated.")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Template Created",
                    f"File saved to:\n{path}\n\n"
                    "Upload this to your shared network drive.",
                ))
            else:
                self._log(f"❌  {err}")
                self._set_status("❌  Template generation failed.")

        threading.Thread(target=_worker, daemon=True).start()

    def _on_send(self):
        path = self._file_var.get().strip()
        if not path:
            messagebox.showwarning("No File", "Please select the Excel file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("File Not Found", f"Cannot find:\n{path}")
            return

        self._send_btn.configure(state="disabled", text="  ⏳  Sending …")
        self._set_status("● Sending emails …")
        self._log("─" * 54)
        self._log(f"Run started: {datetime.now().strftime('%d %b %Y  %H:%M:%S')}")

        backend     = self._backend_var.get()
        display_only = self._preview_var.get()

        def _worker():
            try:
                self._log("📂  Reading Excel …")
                employees, leaves, err = read_excel(path)
                if err:
                    self._log(f"❌  {err}")
                    self._set_status("❌  Read error.")
                    return

                self._log(f"📋  {len(employees)} employees | {len(leaves)} leave records.")

                if backend == "outlook":
                    sender = OutlookSender()
                else:
                    smtp_cfg = SmtpConfig(
                        host     = self._smtp_host.get().strip(),
                        port     = int(self._smtp_port.get().strip() or "587"),
                        username = self._smtp_user.get().strip(),
                        password = self._smtp_pass.get().strip(),
                        use_tls  = self._smtp_tls.get(),
                    )
                    sender = SmtpSender(smtp_cfg)

                sent, errors = sender.send(employees, leaves, self._log, display_only)
                action = "opened for review" if (backend == "outlook" and display_only) else "sent"
                self._log(f"\n🎉  Done! {sent} email(s) {action}.")
                if errors:
                    self._log(f"⚠️   {len(errors)} skipped — see warnings above.")
                self._set_status(f"✅  {sent} email(s) {action}.")

            except Exception as e:
                self._log(f"❌  {e}")
                self._set_status("❌  Error — check log.")
            finally:
                self.root.after(0, lambda: self._send_btn.configure(
                    state="normal", text="  📨  Send Emails Today"))

        threading.Thread(target=_worker, daemon=True).start()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = Tk()
    LeaveMailerApp(root)
    root.mainloop()
