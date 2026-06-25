"""CustomTkinter GUI for the email attachment downloader."""

import os
import threading
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .email_client import (
    EmailClient, EmailProvider, PROVIDER_CONFIG,
    AuthenticationError, EmailClientError, get_provider_by_name
)
from .downloader import AttachmentDownloader, get_extensions_for_types
from .renamer import RENAME_TEMPLATES, create_renamer_from_template
from .preview_window import PreviewWindow
from .date_picker import DatePickerButton


class EmailDownloaderApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title("Email Attachment Downloader")
        self.geometry("750x900")
        self.minsize(650, 800)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.email_client: EmailClient | None = None
        self.is_connected = False
        self.search_results = []
        self.working_thread = None

        self._create_widgets()

    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main scrollable container
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # === Connection Section ===
        self._create_connection_section()

        # === Filter Section ===
        self._create_filter_section()

        # === File Types Section ===
        self._create_file_types_section()

        # === Search Section ===
        self._create_search_section()

        # === Results Section ===
        self._create_results_section()

        # === Rename Section ===
        self._create_rename_section()

        # === Download Section ===
        self._create_download_section()

        # === Progress Section ===
        self._create_progress_section()

        # === Log Section ===
        self._create_log_section()

    def _create_connection_section(self):
        """Create connection input section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="Email Connection", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        # Provider selection
        provider_frame = ctk.CTkFrame(section, fg_color="transparent")
        provider_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(provider_frame, text="Provider:", width=100, anchor="w").pack(side="left")

        provider_names = [config["name"] for config in PROVIDER_CONFIG.values()]
        self.provider_dropdown = ctk.CTkOptionMenu(
            provider_frame,
            values=provider_names,
            width=200,
            command=self._on_provider_change
        )
        self.provider_dropdown.set(provider_names[0])  # Default to Gmail
        self.provider_dropdown.pack(side="left")

        # Help button
        self.help_btn = ctk.CTkButton(
            provider_frame,
            text="?",
            width=30,
            command=self._show_provider_help
        )
        self.help_btn.pack(side="left", padx=(10, 0))

        # Email input
        email_frame = ctk.CTkFrame(section, fg_color="transparent")
        email_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(email_frame, text="Email:", width=100, anchor="w").pack(side="left")
        self.email_entry = ctk.CTkEntry(email_frame, width=300, placeholder_text="your.email@gmail.com")
        self.email_entry.pack(side="left", padx=(0, 10))

        # Password input
        pass_frame = ctk.CTkFrame(section, fg_color="transparent")
        pass_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(pass_frame, text="Password:", width=100, anchor="w").pack(side="left")
        self.password_entry = ctk.CTkEntry(pass_frame, width=300, show="*", placeholder_text="App Password")
        self.password_entry.pack(side="left", padx=(0, 10))

        # Password hint label
        self.password_hint = ctk.CTkLabel(
            section,
            text="Gmail requires an App Password, not your regular password",
            text_color="gray",
            font=("", 11)
        )
        self.password_hint.pack(anchor="w", padx=110, pady=(0, 5))

        # Connect button
        btn_frame = ctk.CTkFrame(section, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(5, 10))

        self.connect_btn = ctk.CTkButton(
            btn_frame, text="Connect", command=self._toggle_connection, width=120
        )
        self.connect_btn.pack(side="left")

        self.status_label = ctk.CTkLabel(btn_frame, text="Not connected", text_color="gray")
        self.status_label.pack(side="left", padx=15)

    def _create_filter_section(self):
        """Create email filter section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="Email Filters", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        # Sender filter
        sender_frame = ctk.CTkFrame(section, fg_color="transparent")
        sender_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(sender_frame, text="From:", width=100, anchor="w").pack(side="left")
        self.sender_entry = ctk.CTkEntry(
            sender_frame, width=300, placeholder_text="sender@example.com"
        )
        self.sender_entry.pack(side="left")

        # Subject filter
        subject_frame = ctk.CTkFrame(section, fg_color="transparent")
        subject_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(subject_frame, text="Subject:", width=100, anchor="w").pack(side="left")
        self.subject_entry = ctk.CTkEntry(
            subject_frame, width=300, placeholder_text="Contains text..."
        )
        self.subject_entry.pack(side="left")

        # Date range with calendar pickers
        date_frame = ctk.CTkFrame(section, fg_color="transparent")
        date_frame.pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(date_frame, text="Date From:", width=100, anchor="w").pack(side="left")
        self.date_from_picker = DatePickerButton(date_frame, placeholder="Start date")
        self.date_from_picker.pack(side="left", padx=(0, 20))

        ctk.CTkLabel(date_frame, text="To:", anchor="w").pack(side="left")
        self.date_to_picker = DatePickerButton(date_frame, placeholder="End date")
        self.date_to_picker.pack(side="left", padx=(5, 0))

    def _create_file_types_section(self):
        """Create file type filter section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="File Types", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        types_frame = ctk.CTkFrame(section, fg_color="transparent")
        types_frame.pack(fill="x", padx=10, pady=(0, 10))

        # Create checkboxes for each file type
        self.file_type_vars = {}

        row1 = ctk.CTkFrame(types_frame, fg_color="transparent")
        row1.pack(fill="x", pady=2)

        row2 = ctk.CTkFrame(types_frame, fg_color="transparent")
        row2.pack(fill="x", pady=2)

        types_row1 = ["pdf", "images", "documents"]
        types_row2 = ["spreadsheets", "presentations", "archives"]

        for type_name in types_row1:
            var = ctk.BooleanVar(value=True)
            self.file_type_vars[type_name] = var
            ctk.CTkCheckBox(row1, text=type_name.upper(), variable=var, width=120).pack(
                side="left", padx=(0, 10)
            )

        for type_name in types_row2:
            var = ctk.BooleanVar(value=True)
            self.file_type_vars[type_name] = var
            ctk.CTkCheckBox(row2, text=type_name.upper(), variable=var, width=120).pack(
                side="left", padx=(0, 10)
            )

    def _create_search_section(self):
        """Create search button section."""
        section = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        section.pack(fill="x", pady=(0, 15))

        self.search_btn = ctk.CTkButton(
            section,
            text="Search Emails",
            command=self._start_search,
            height=40,
            font=("", 14, "bold"),
            state="disabled"
        )
        self.search_btn.pack()

    def _create_results_section(self):
        """Create results display section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="Results", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        results_frame = ctk.CTkFrame(section, fg_color="transparent")
        results_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.results_label = ctk.CTkLabel(
            results_frame,
            text="No search performed yet",
            text_color="gray"
        )
        self.results_label.pack(side="left")

        self.preview_btn = ctk.CTkButton(
            results_frame,
            text="Preview Results",
            command=self._show_preview,
            width=120,
            state="disabled"
        )
        self.preview_btn.pack(side="right")

    def _create_rename_section(self):
        """Create rename pattern section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="Rename Pattern", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        rename_frame = ctk.CTkFrame(section, fg_color="transparent")
        rename_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(rename_frame, text="Pattern:", width=100, anchor="w").pack(side="left")

        # Dropdown with template names
        template_names = [info["name"] for info in RENAME_TEMPLATES.values()]
        self.rename_dropdown = ctk.CTkOptionMenu(
            rename_frame,
            values=template_names,
            width=250
        )
        self.rename_dropdown.set("Date + Filename")
        self.rename_dropdown.pack(side="left")

        # Options checkboxes
        options_frame = ctk.CTkFrame(section, fg_color="transparent")
        options_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.lowercase_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame, text="Convert to lowercase", variable=self.lowercase_var
        ).pack(side="left", padx=(100, 0))

    def _create_download_section(self):
        """Create download options section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(section, text="Download Location", font=("", 16, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        path_frame = ctk.CTkFrame(section, fg_color="transparent")
        path_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.download_path = ctk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Downloads", "EmailAttachments")
        )
        self.path_entry = ctk.CTkEntry(path_frame, textvariable=self.download_path, width=400)
        self.path_entry.pack(side="left", padx=(0, 10))

        ctk.CTkButton(path_frame, text="Browse", command=self._browse_folder, width=80).pack(
            side="left"
        )

        # Download button
        self.download_btn = ctk.CTkButton(
            section,
            text="Download All Attachments",
            command=self._start_download,
            height=40,
            font=("", 14, "bold"),
            state="disabled"
        )
        self.download_btn.pack(pady=(5, 15))

    def _create_progress_section(self):
        """Create progress tracking section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="x", pady=(0, 15))

        self.progress_label = ctk.CTkLabel(section, text="Ready")
        self.progress_label.pack(anchor="w", padx=10, pady=(10, 5))

        self.progress_bar = ctk.CTkProgressBar(section, width=500)
        self.progress_bar.pack(padx=10, pady=(0, 5))
        self.progress_bar.set(0)

        self.stats_label = ctk.CTkLabel(section, text="", text_color="gray")
        self.stats_label.pack(anchor="w", padx=10, pady=(0, 10))

    def _create_log_section(self):
        """Create log output section."""
        section = ctk.CTkFrame(self.main_frame)
        section.pack(fill="both", expand=True)

        ctk.CTkLabel(section, text="Log", font=("", 14, "bold")).pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        self.log_text = ctk.CTkTextbox(section, height=100)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _log(self, message: str):
        """Add message to log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

    def _on_provider_change(self, choice: str):
        """Handle provider dropdown change."""
        if choice == "Gmail":
            self.password_hint.configure(
                text="Gmail requires an App Password, not your regular password"
            )
            self.email_entry.configure(placeholder_text="your.email@gmail.com")
        else:
            self.password_hint.configure(
                text="Use App Password if 2FA is enabled, otherwise use regular password"
            )
            self.email_entry.configure(placeholder_text="your.email@outlook.com")

    def _show_provider_help(self):
        """Show help dialog for selected provider."""
        provider_name = self.provider_dropdown.get()
        provider = get_provider_by_name(provider_name)
        help_text = PROVIDER_CONFIG[provider]["help_text"]

        messagebox.showinfo(f"{provider_name} Setup", help_text)

    def _get_selected_provider(self) -> EmailProvider:
        """Get currently selected email provider."""
        provider_name = self.provider_dropdown.get()
        return get_provider_by_name(provider_name)

    def _toggle_connection(self):
        """Connect or disconnect from email server."""
        if self.is_connected:
            if self.email_client:
                self.email_client.disconnect()
            self.email_client = None
            self.is_connected = False
            self.connect_btn.configure(text="Connect")
            self.status_label.configure(text="Disconnected", text_color="gray")
            self.search_btn.configure(state="disabled")
            self.download_btn.configure(state="disabled")
            self.provider_dropdown.configure(state="normal")
            self._log("Disconnected")
        else:
            self._connect()

    def _connect(self):
        """Connect to email server via IMAP."""
        email_addr = self.email_entry.get().strip()
        password = self.password_entry.get().strip()

        if not email_addr or not password:
            messagebox.showerror("Error", "Please enter email and password")
            return

        provider = self._get_selected_provider()

        try:
            self.connect_btn.configure(state="disabled")
            self.provider_dropdown.configure(state="disabled")
            self.status_label.configure(text="Connecting...", text_color="yellow")
            self.update()

            # Create client for selected provider
            self.email_client = EmailClient(provider)
            self.email_client.connect(email_addr, password)
            self.is_connected = True

            self.connect_btn.configure(text="Disconnect", state="normal")
            self.status_label.configure(text="Connected", text_color="green")
            self.search_btn.configure(state="normal")
            self._log(f"Connected to {self.email_client.provider_name} as {email_addr}")

        except AuthenticationError as e:
            self.connect_btn.configure(state="normal")
            self.provider_dropdown.configure(state="normal")
            self.status_label.configure(text="Auth failed", text_color="red")
            self._log(f"Authentication failed")
            messagebox.showerror("Authentication Error", str(e))

        except Exception as e:
            self.connect_btn.configure(state="normal")
            self.provider_dropdown.configure(state="normal")
            self.status_label.configure(text="Connection failed", text_color="red")
            self._log(f"Connection failed: {str(e)}")
            messagebox.showerror("Connection Error", str(e))

    def _browse_folder(self):
        """Open folder browser dialog."""
        folder = filedialog.askdirectory(initialdir=self.download_path.get())
        if folder:
            self.download_path.set(folder)

    def _parse_date(self, date_str: str) -> datetime | None:
        """Parse date string to datetime."""
        if not date_str.strip():
            return None
        try:
            return datetime.strptime(date_str.strip(), "%Y-%m-%d")
        except ValueError:
            return None

    def _get_selected_extensions(self) -> list[str] | None:
        """Get list of selected file extensions."""
        selected_types = [
            name for name, var in self.file_type_vars.items() if var.get()
        ]

        if not selected_types or len(selected_types) == len(self.file_type_vars):
            return None  # All types

        return get_extensions_for_types(selected_types)

    def _get_selected_template_key(self) -> str:
        """Get the template key from dropdown selection."""
        selected_name = self.rename_dropdown.get()

        for key, info in RENAME_TEMPLATES.items():
            if info["name"] == selected_name:
                return key

        return "original"

    def _build_renamer(self):
        """Build renamer from selected template."""
        template_key = self._get_selected_template_key()
        renamer = create_renamer_from_template(template_key)
        renamer.set_options(lowercase=self.lowercase_var.get())
        return renamer

    def _start_search(self):
        """Start the search process in a background thread."""
        if not self.is_connected or not self.email_client:
            messagebox.showerror("Error", "Please connect to email server first")
            return

        self.search_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="Searching emails...")
        self.search_results = []

        self.working_thread = threading.Thread(target=self._search_worker)
        self.working_thread.daemon = True
        self.working_thread.start()

    def _search_worker(self):
        """Background worker for searching emails."""
        try:
            sender = self.sender_entry.get().strip() or None
            subject = self.subject_entry.get().strip() or None
            date_from = self.date_from_picker.get_datetime()
            date_to = self.date_to_picker.get_datetime()

            self._log(f"Searching emails - From: {sender or 'any'}, Subject: {subject or 'any'}")

            def search_progress(current, total):
                self.after(0, lambda: self.progress_label.configure(
                    text=f"Fetching emails: {current}/{total}"
                ))
                self.after(0, lambda: self.progress_bar.set(current / total))

            emails = self.email_client.search_emails(
                sender=sender,
                subject=subject,
                date_from=date_from,
                date_to=date_to,
                progress_callback=search_progress
            )

            self.search_results = emails

            # Count attachments
            total_attachments = sum(e.attachment_count for e in emails)

            self._log(f"Found {len(emails)} emails with {total_attachments} attachments")

            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self.progress_label.configure(text="Search complete"))
            self.after(0, lambda: self.results_label.configure(
                text=f"Found: {len(emails)} emails, {total_attachments} attachments",
                text_color="white"
            ))

            if emails:
                self.after(0, lambda: self.preview_btn.configure(state="normal"))
                self.after(0, lambda: self.download_btn.configure(state="normal"))
            else:
                self.after(0, lambda: self.preview_btn.configure(state="disabled"))
                self.after(0, lambda: self.download_btn.configure(state="disabled"))
                
        except Exception as e:
            error_msg = str(e)
            self._log(f"Search error: {error_msg}")
            self.after(0, lambda msg=error_msg: messagebox.showerror("Search Error", msg))
            self.after(0, lambda: self.progress_label.configure(text="Search failed"))

        finally:
            self.after(0, lambda: self.search_btn.configure(state="normal"))

    def _show_preview(self):
        """Show preview window with search results."""
        if not self.search_results:
            return

        PreviewWindow(
            self,
            self.search_results,
            on_download=self._download_selected
        )

    def _download_selected(self, emails: list):
        """Download attachments from selected emails."""
        self.search_results = emails
        self._start_download()

    def _start_download(self):
        """Start the download process in a background thread."""
        if not self.search_results:
            messagebox.showwarning("Warning", "No emails to download from. Run a search first.")
            return

        self.download_btn.configure(state="disabled")
        self.search_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="Preparing download...")

        self.working_thread = threading.Thread(target=self._download_worker)
        self.working_thread.daemon = True
        self.working_thread.start()

    def _download_worker(self):
        """Background worker for downloading attachments."""
        try:
            renamer = self._build_renamer()
            extensions = self._get_selected_extensions()

            downloader = AttachmentDownloader(
                self.download_path.get(),
                renamer,
                extensions
            )

            self._log(f"Downloading from {len(self.search_results)} emails...")

            def download_progress(current, total, filename):
                progress = current / total
                self.after(0, lambda: self.progress_label.configure(
                    text=f"Downloading: {filename}"
                ))
                self.after(0, lambda: self.progress_bar.set(progress))

            results = downloader.download_from_emails(
                self.search_results,
                download_progress
            )

            success_count = sum(1 for r in results if r.success)
            fail_count = len(results) - success_count

            for result in results:
                if result.success:
                    self._log(f"Saved: {result.saved_filename}")
                else:
                    self._log(f"Failed: {result.original_filename} - {result.error}")

            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self.progress_label.configure(text="Download complete"))
            self.after(0, lambda: self.stats_label.configure(
                text=f"Downloaded: {success_count} | Failed: {fail_count}"
            ))
            self._log(f"Complete: {success_count} files saved, {fail_count} failed")

        except Exception as e:
            error_msg = str(e)
            self._log(f"Download error: {error_msg}")
            self.after(0, lambda msg=error_msg: messagebox.showerror("Download Error", msg))

        finally:
            self.after(0, lambda: self.download_btn.configure(state="normal"))
            self.after(0, lambda: self.search_btn.configure(state="normal"))

    def on_closing(self):
        """Handle window close event."""
        if self.is_connected and self.email_client:
            self.email_client.disconnect()
        self.destroy()


def run_app():
    """Run the application."""
    app = EmailDownloaderApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
