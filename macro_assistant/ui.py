import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


from dotenv import load_dotenv

from .config import app_directory, configured_provider, load_config
from .excel_injector import CompileCheckError, TrustAccessError, enable_vbom_for_current_user, inject_macro
from .logging_setup import get_logger
from .providers.factory import available_providers
from .safety import scan_vba
from .vba_generation import generate_macro
from .workbook_context import read_workbook_context


class MacroAssistantApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.logger = get_logger()
        load_dotenv(app_directory() / ".env", override=False)
        self.preview_data = None
        self.file_path = tk.StringVar()
        self.provider = tk.StringVar(value=configured_provider())
        self.status = tk.StringVar(value="Choose an Excel file to get started.")

        root.title("Excel Macro Assistant")
        root.geometry("780x700")
        root.minsize(650, 570)
        root.configure(background="#f4f6f4")
        self._style()

        page = ttk.Frame(root, padding=(28, 24))
        page.pack(fill="both", expand=True)
        ttk.Label(page, text="Excel Macro Assistant", style="Title.TLabel").pack(anchor="w")
        ttk.Label(page, text="Describe a task. Preview the macro. Save a new macro-enabled copy.", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 20))

        ttk.Label(page, text="1  Workbook", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        file_row = ttk.Frame(page)
        file_row.pack(fill="x", pady=(0, 18))
        self.file_entry = ttk.Entry(file_row, textvariable=self.file_path, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True, ipady=6)
        ttk.Button(file_row, text="Browse...", command=self._browse).pack(side="left", padx=(9, 0), ipady=3)

        ttk.Label(page, text="2  What should the macro do?", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        self.request_text = tk.Text(page, height=5, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1, padx=9, pady=8)
        self.request_text.pack(fill="x", pady=(0, 14))

        provider_row = ttk.Frame(page)
        provider_row.pack(fill="x", pady=(0, 17))
        ttk.Label(provider_row, text="AI provider", style="Section.TLabel").pack(side="left")
        self.provider_box = ttk.Combobox(provider_row, textvariable=self.provider, values=available_providers(), state="readonly", width=14)
        self.provider_box.pack(side="left", padx=(12, 0))
        ttk.Label(provider_row, text="Model and API key come from your .env file", style="Hint.TLabel").pack(side="left", padx=(12, 0))

        button_row = ttk.Frame(page)
        button_row.pack(fill="x", pady=(0, 15))
        self.preview_button = ttk.Button(button_row, text="Preview macro", style="Primary.TButton", command=self._preview)
        self.preview_button.pack(side="left", ipadx=12, ipady=5)
        self.apply_button = ttk.Button(button_row, text="Apply to a new copy", command=self._apply, state="disabled")
        self.apply_button.pack(side="left", padx=(10, 0), ipadx=8, ipady=5)
        if sys.platform == "win32":
            ttk.Button(button_row, text="Enable Excel access", command=self._enable_vbom).pack(side="right", ipady=4)


        ttk.Label(page, text="3  Preview", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        self.summary = tk.Text(page, height=6, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1, padx=9, pady=8, state="disabled")
        self.summary.pack(fill="both", expand=True)
        self.warning = ttk.Label(page, text="", style="Warning.TLabel", wraplength=710, justify="left")
        self.warning.pack(fill="x", anchor="w", pady=(8, 0))
        ttk.Separator(page).pack(fill="x", pady=(13, 9))
        ttk.Label(page, textvariable=self.status, style="Status.TLabel", wraplength=710).pack(fill="x", anchor="w")

    def _style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6f4")
        style.configure("TLabel", background="#f4f6f4", foreground="#24342e", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 21), foreground="#183b31")
        style.configure("Subtitle.TLabel", foreground="#53665e")
        style.configure("Section.TLabel", font=("Segoe UI Semibold", 10), foreground="#1f4337")
        style.configure("Hint.TLabel", foreground="#66756f", font=("Segoe UI", 9))
        style.configure("Status.TLabel", foreground="#40554b")
        style.configure("Warning.TLabel", foreground="#8a3a17", font=("Segoe UI Semibold", 9))
        style.configure("TButton", font=("Segoe UI Semibold", 9), padding=(10, 6))
        style.configure("Primary.TButton", background="#176b53", foreground="white")
        style.map("Primary.TButton", background=[("active", "#105440"), ("disabled", "#aab9b2")])
        style.configure("TCombobox", padding=4)

    def _browse(self):
        selected = filedialog.askopenfilename(
            title="Choose an Excel workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if selected:
            self.file_path.set(selected)
            self.preview_data = None
            self.apply_button.configure(state="disabled")
            self.warning.configure(text="")
            self._set_text(self.summary, "")
            self.status.set("Workbook selected. Describe the task and preview the macro.")

    def _preview(self):
        source = Path(self.file_path.get())
        request = self.request_text.get("1.0", "end").strip()
        if not source.is_file() or source.suffix.lower() not in {".xlsx", ".xlsm"}:
            messagebox.showerror("Choose a workbook", "Select an existing .xlsx or .xlsm file first.", parent=self.root)
            return
        if not request:
            messagebox.showerror("Describe the task", "Write what you want the macro to do first.", parent=self.root)
            return
        self.preview_data = None
        self.apply_button.configure(state="disabled")
        self._busy(True, "Reading workbook headings and preparing your preview...")
        self.logger.info("preview_started")
        threading.Thread(target=self._preview_worker, args=(source, request, self.provider.get()), daemon=True).start()

    def _preview_worker(self, source: Path, request: str, provider_id: str):
        try:
            context = read_workbook_context(source)
            config = load_config(provider_override=provider_id)
            result = generate_macro(config.provider, config.model, config.api_key, request, context)
            data = {**result, "request": request, "context": context, "provider": config.provider, "model": config.model}
            self.root.after(0, lambda: self._preview_done(data))
            self.logger.info("preview_completed")
        except Exception as error:
            self.logger.info("preview_failed")
            self.root.after(0, lambda error=error: self._failed(error))

    def _preview_done(self, data: dict):
        self.preview_data = data
        findings = scan_vba(data["vba_code"])
        self._set_text(self.summary, data["summary"])
        if findings:
            self.warning.configure(text="Safety review needed: " + "; ".join(findings) + ". Apply will ask you to confirm before adding this code.")
        else:
            self.warning.configure(text="No listed high-risk operations were detected. This automated check cannot guarantee that all VBA is safe.")
        self.apply_button.configure(state="normal")
        self._busy(False, "Preview ready. Review the summary before applying.")

    def _apply(self):
        if not self.preview_data:
            return
        data = self.preview_data
        if (
            data["request"] != self.request_text.get("1.0", "end").strip()
            or data["provider"] != self.provider.get()
        ):
            messagebox.showinfo("Preview again", "The request or AI provider changed. Preview the macro again before applying it.", parent=self.root)
            self.apply_button.configure(state="disabled")
            return
        findings = scan_vba(data["vba_code"])
        if findings:
            proceed = messagebox.askyesno(
                "Review safety warning",
                "The macro contains operations that may affect files, Windows settings, or internet connections:\n\n"
                + "\n".join(f"• {finding}" for finding in findings)
                + "\n\nOnly continue if you understand and trust these operations. Add this macro anyway?",
                parent=self.root,
            )
            if not proceed:
                return
        self._busy(True, "Adding the macro to a new workbook copy...")
        self.logger.info("apply_started")
        threading.Thread(target=self._apply_worker, args=(Path(self.file_path.get()), data, bool(findings)), daemon=True).start()

    def _apply_worker(self, source: Path, data: dict, risk_confirmed: bool):
        current = data
        for attempt in range(3):
            findings = scan_vba(current["vba_code"])
            if findings and (attempt > 0 or not risk_confirmed):
                error = RuntimeError("An automatic repair produced code requiring a new safety review. Preview the macro again before applying.")
                self.root.after(0, lambda error=error: self._failed(error))
                return
            try:
                result = inject_macro(source, current["vba_code"])
                self.logger.info("apply_completed")
                self.root.after(0, lambda result=result: self._apply_done(result))
                return
            except TrustAccessError as error:
                self.logger.info("apply_blocked_trust_setting")
                self.root.after(0, lambda error=error: self._failed(error, trust_access=True))
                return
            except Exception as error:
                if attempt == 2:
                    self.logger.info("apply_failed")
                    self.root.after(0, lambda error=error: self._failed(error))
                    return
                try:
                    config = load_config(provider_override=data["provider"])
                    current = generate_macro(
                        config.provider,
                        config.model,
                        config.api_key,
                        data["request"],
                        data["context"],
                        repair_error=str(error),
                    )
                except Exception as repair_error:
                    self.logger.info("repair_failed")
                    self.root.after(0, lambda error=repair_error: self._failed(error))
                    return

    def _apply_done(self, result):
        self._busy(False, f"Saved to: {result.output_path}")
        if result.applied_in_excel:
            compile_note = "Excel's VBA editor accepted a compile check." if result.compile_checked else "Excel's automatic compile check was unavailable."
            module_note = f"\n\nVBA module export saved to:\n{result.module_path}" if result.module_path else ""
            messagebox.showinfo(
                "Macro Applied Successfully",
                f"The macro '{result.macro_name}' was injected and executed!\n\n"
                f"New output workbook:\n{result.output_path}"
                f"{module_note}\n\n"
                f"When you open the new workbook, you can view the changes immediately, or press Alt + F8 to run the macro again if needed.\n\n"
                f"{compile_note}",
                parent=self.root,
            )
        else:
            module_info = f"\n\nVBA Module (.bas):\n{result.module_path}" if result.module_path else ""
            messagebox.showinfo(
                "Macro Ready for Excel",
                f"A new workbook and the VBA macro module have been created successfully!\n\n"
                f"Workbook copy:\n{result.output_path}"
                f"{module_info}\n\n"
                f"To run the macro inside Excel:\n"
                f"1. Open the workbook in Microsoft Excel.\n"
                f"2. Press Alt + F11 (or Option + F11 on Mac) to open the VBA Editor.\n"
                f"3. Click File > Import File... and choose the .bas file.\n"
                f"4. Press Alt + F8 (or Option + F8 on Mac) to select and run '{result.macro_name}'.",
                parent=self.root,
            )


    def _failed(self, error: Exception, trust_access: bool = False):
        self._busy(False, "The task could not be completed.")
        if trust_access:
            messagebox.showerror(
                "Excel access is turned off",
                "Excel is blocking the app from adding a macro. Use the 'Enable Excel access' button after closing Excel. "
                "If this setting is managed by your organization, ask IT for help.",
                parent=self.root,
            )
        else:
            messagebox.showerror("Could not complete", str(error), parent=self.root)

    def _enable_vbom(self):
        confirmed = messagebox.askyesno(
            "Enable Excel macro access",
            "This changes an Excel setting for your Windows account so this app can add a VBA module. "
            "It does not enable macros to run. Close all Excel windows first. Continue?",
            parent=self.root,
        )
        if not confirmed:
            return
        try:
            enable_vbom_for_current_user()
            messagebox.showinfo("Excel access enabled", "The setting was enabled for your Windows account. Try Preview or Apply again.", parent=self.root)
        except Exception as error:
            messagebox.showerror("Could not enable Excel access", str(error), parent=self.root)

    def _set_text(self, widget: tk.Text, text: str):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        if widget is self.summary:
            widget.configure(state="disabled")

    def _busy(self, busy: bool, status: str):
        self.status.set(status)
        state = "disabled" if busy else "normal"
        self.preview_button.configure(state=state)
        self.provider_box.configure(state="disabled" if busy else "readonly")
        if busy:
            self.apply_button.configure(state="disabled")
        elif self.preview_data:
            self.apply_button.configure(state="normal")


def run():
    root = tk.Tk()
    MacroAssistantApp(root)
    root.mainloop()
