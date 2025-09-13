import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, font
import requests
import imaplib
import email
from bs4 import BeautifulSoup
import re
import json
import time
from urllib.parse import urlparse, urljoin
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
import threading
from datetime import datetime, timedelta
import sqlite3
import os
import webbrowser


@dataclass
class UnsubscribeInfo:
    """Data for unsub info"""
    email_id: str
    sender: str
    subject: str
    unsubscribe_links: List[str]
    unsubscribe_email: Optional[str]
    list_unsubscribe_header: Optional[str]
    confidence_score: float
    email_date: datetime
    content_preview: str


class ModernTheme:
    """more modern UI theme with colors and styling"""

    # Colours for TKINTER
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    SECONDARY = "#64748b"
    SUCCESS = "#10b981"
    WARNING = "#f59e0b"
    DANGER = "#ef4444"
    BACKGROUND = "#f8fafc"
    SURFACE = "#ffffff"
    ON_SURFACE = "#1e293b"
    ON_SURFACE_VARIANT = "#64748b"
    OUTLINE = "#cbd5e1"

    @classmethod
    def configure_styles(cls):
        style = ttk.Style()

        style.configure(
            "Modern.TButton",
            padding=(20, 12),
            font=('Segoe UI', 10),
            borderwidth=0,
            focuscolor="none"
        )

        style.configure(
            "Primary.TButton",
            background=cls.PRIMARY,
            foreground="white",
            padding=(20, 12),
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            focuscolor="none"
        )

        style.map(
            "Primary.TButton",
            background=[('active', cls.PRIMARY_HOVER)]
        )

        style.configure(
            "Success.TButton",
            background=cls.SUCCESS,
            foreground="white",
            padding=(16, 10),
            font=('Segoe UI', 9, 'bold'),
            borderwidth=0
        )

        style.configure(
            "Danger.TButton",
            background=cls.DANGER,
            foreground="white",
            padding=(16, 10),
            font=('Segoe UI', 9, 'bold'),
            borderwidth=0
        )

        style.configure(
            "Card.TFrame",
            background=cls.SURFACE,
            relief="flat",
            borderwidth=1
        )

        style.configure(
            "Heading.TLabel",
            font=('Segoe UI', 24, 'bold'),
            foreground=cls.ON_SURFACE,
            background=cls.BACKGROUND
        )

        style.configure(
            "Subheading.TLabel",
            font=('Segoe UI', 14, 'bold'),
            foreground=cls.ON_SURFACE,
            background=cls.SURFACE
        )

        style.configure(
            "Body.TLabel",
            font=('Segoe UI', 10),
            foreground=cls.ON_SURFACE_VARIANT,
            background=cls.SURFACE
        )

        style.configure(
            "Modern.TEntry",
            padding=12,
            font=('Segoe UI', 10),
            borderwidth=1,
            relief="solid"
        )

        style.configure(
            "Modern.TNotebook",
            background=cls.BACKGROUND,
            borderwidth=0,
            tabmargins=0
        )

        style.configure(
            "Modern.TNotebook.Tab",
            padding=(20, 12),
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0
        )


class MailMuteAlgorithm:
    """ai-powered email content analyzer for better unsubscribe detection"""

    def __init__(self):
        self.unsubscribe_patterns = [
            r'unsubscribe',
            r'opt.?out',
            r'remove.*list',
            r'stop.*email',
            r'cancel.*subscription',
            r'manage.*preferences',
            r'email.*settings',
            r'notification.*settings',
            r'mailing.*list.*remove',
            r'leave.*list',
            r'turn.*off.*notifications'
        ]

        self.sender_blacklist = [
            'noreply', 'no-reply', 'donotreply', 'do-not-reply',
            'newsletter', 'marketing', 'promo', 'deals'
        ]

    def analyze_email_content(self, msg) -> UnsubscribeInfo:
        """Analyzes email content using AI-like pattern matching"""
        sender = self.extract_sender(msg)
        subject = msg.get('Subject', 'No Subject')
        email_id = msg.get('Message-ID', '')
        email_date = self.parse_email_date(msg.get('Date', ''))

        unsubscribe_links = []
        unsubscribe_email = None
        list_unsubscribe_header = msg.get('List-Unsubscribe', '')

        if list_unsubscribe_header:
            unsubscribe_email, header_links = self.parse_list_unsubscribe_header(list_unsubscribe_header)
            unsubscribe_links.extend(header_links)

        content_preview = ""
        html_content = self.extract_html_content(msg)
        if html_content:
            content_preview = self.extract_text_preview(html_content)
            links = self.extract_smart_unsubscribe_links(html_content)
            unsubscribe_links.extend(links)

        confidence_score = self.calculate_confidence_score(
            sender, subject, unsubscribe_links, content_preview, list_unsubscribe_header
        )

        return UnsubscribeInfo(
            email_id=email_id,
            sender=sender,
            subject=subject,
            unsubscribe_links=list(set(unsubscribe_links)),
            unsubscribe_email=unsubscribe_email,
            list_unsubscribe_header=list_unsubscribe_header,
            confidence_score=confidence_score,
            email_date=email_date,
            content_preview=content_preview[:200] + "..." if len(content_preview) > 200 else content_preview
        )

    def extract_sender(self, msg) -> str:
        """Extract sender information"""
        sender = msg.get('From', 'Unknown Sender')
        if '<' in sender and '>' in sender:
            sender = sender.split('<')[1].split('>')[0]
        return sender

    def parse_email_date(self, date_str: str) -> datetime:
        """Parse email date"""
        try:
            return email.utils.parsedate_to_datetime(date_str)
        except:
            return datetime.now()

    def parse_list_unsubscribe_header(self, header: str) -> Tuple[Optional[str], List[str]]:
        """Parse List-Unsubscribe header"""
        unsubscribe_email = None
        links = []
        mailto_match = re.search(r'<mailto:([^>]+)>', header)
        if mailto_match:
            unsubscribe_email = mailto_match.group(1)

        http_matches = re.findall(r'<(https?://[^>]+)>', header)
        links.extend(http_matches)

        return unsubscribe_email, links

    def extract_html_content(self, msg) -> str:
        """Extract HTML content from email"""
        html_content = ""

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload:
                        try:
                            html_content = payload.decode("utf-8")
                        except UnicodeDecodeError:
                            try:
                                html_content = payload.decode("latin-1")
                            except:
                                continue
                        break
        else:
            if msg.get_content_type() == "text/html":
                payload = msg.get_payload(decode=True)
                if payload:
                    try:
                        html_content = payload.decode("utf-8")
                    except UnicodeDecodeError:
                        try:
                            html_content = payload.decode("latin-1")
                        except:
                            pass

        return html_content

    def extract_text_preview(self, html_content: str) -> str:
        """Extract text preview from HTML"""
        try:
            soup = BeautifulSoup(html_content, "html.parser")
            return soup.get_text()[:500]
        except:
            return ""

    def extract_smart_unsubscribe_links(self, html_content: str) -> List[str]:
        """Smart extraction of unsubscribe links using AI-like pattern matching"""
        try:
            soup = BeautifulSoup(html_content, "html.parser")
            links = []

            for link in soup.find_all("a", href=True):
                href = link['href']
                link_text = link.get_text().strip().lower()

                for pattern in self.unsubscribe_patterns:
                    if re.search(pattern, href.lower()) or re.search(pattern, link_text):
                        if href.startswith(('http://', 'https://')):
                            links.append(href)
                        break
            for button in soup.find_all(["button", "input"], type="button"):
                button_text = button.get_text().strip().lower()
                onclick = button.get('onclick', '')

                for pattern in self.unsubscribe_patterns:
                    if re.search(pattern, button_text) or re.search(pattern, onclick):
                        url_match = re.search(r'https?://[^\s\'"]+', onclick)
                        if url_match:
                            links.append(url_match.group())
                        break

            return links
        except Exception:
            return []

    def calculate_confidence_score(self, sender: str, subject: str,
                                   unsubscribe_links: List[str], content_preview: str,
                                   list_unsubscribe_header: str) -> float:
        """Calculate confidence score for unsubscribe detection"""
        score = 0.0

        if list_unsubscribe_header:
            score += 0.4

        if unsubscribe_links:
            score += 0.3 + min(len(unsubscribe_links) * 0.1, 0.2)
        sender_lower = sender.lower()
        for blacklist_term in self.sender_blacklist:
            if blacklist_term in sender_lower:
                score += 0.1
                break
        subject_lower = subject.lower()
        for pattern in self.unsubscribe_patterns[:5]:
            if re.search(pattern, subject_lower):
                score += 0.1
                break
        content_lower = content_preview.lower()
        pattern_matches = sum(1 for pattern in self.unsubscribe_patterns
                              if re.search(pattern, content_lower))
        score += min(pattern_matches * 0.05, 0.2)

        return min(score, 1.0)


class DatabaseManager:
    """Manage SQLite database for storing unsubscribe history"""

    def __init__(self, db_path: str = "unsubscribe_history.db"):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Initialize database tables"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS unsubscribe_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id TEXT UNIQUE,
                sender TEXT,
                subject TEXT,
                unsubscribe_date TIMESTAMP,
                status TEXT,
                method TEXT,
                notes TEXT
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sender_stats (
                sender TEXT PRIMARY KEY,
                total_emails INTEGER,
                unsubscribed INTEGER,
                last_seen TIMESTAMP
            )
        ''')

        conn.commit()
        conn.close()

    def record_unsubscribe(self, unsubscribe_info: UnsubscribeInfo, status: str, method: str, notes: str = ""):
        """Record unsubscribe attempt"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT OR REPLACE INTO unsubscribe_history 
            (email_id, sender, subject, unsubscribe_date, status, method, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (unsubscribe_info.email_id, unsubscribe_info.sender, unsubscribe_info.subject,
              datetime.now(), status, method, notes))

        conn.commit()
        conn.close()

    def get_history(self, limit: int = 100):
        """Get unsubscribe history"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            SELECT sender, subject, unsubscribe_date, status, method, notes
            FROM unsubscribe_history
            ORDER BY unsubscribe_date DESC
            LIMIT ?
        ''', (limit,))

        results = cursor.fetchall()
        conn.close()
        return results

    def get_sender_stats(self):
        """Get sender statistics"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            SELECT sender, total_emails, unsubscribed, last_seen
            FROM sender_stats
            ORDER BY total_emails DESC
            LIMIT 20
        ''')

        results = cursor.fetchall()
        conn.close()
        return results

    def update_sender_stats(self, sender: str):
        """Update sender statistics"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT OR REPLACE INTO sender_stats (sender, total_emails, unsubscribed, last_seen)
            VALUES (?, 
                    COALESCE((SELECT total_emails FROM sender_stats WHERE sender = ?), 0) + 1,
                    COALESCE((SELECT unsubscribed FROM sender_stats WHERE sender = ?), 0),
                    ?)
        ''', (sender, sender, sender, datetime.now()))

        conn.commit()
        conn.close()


class SmartUnsubscriber:
    """Smart unsubscriber with multiple strategies"""

    def __init__(self, db_manager: DatabaseManager):
        self.db_manager = db_manager
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def unsubscribe(self, unsubscribe_info: UnsubscribeInfo) -> Dict[str, str]:
        """Attempt to unsubscribe using multiple methods"""
        results = {}

        if unsubscribe_info.list_unsubscribe_header:
            if unsubscribe_info.unsubscribe_email:
                result = self.unsubscribe_via_email(unsubscribe_info.unsubscribe_email)
                results['email'] = result
                self.db_manager.record_unsubscribe(unsubscribe_info, result, 'email')
        for i, link in enumerate(unsubscribe_info.unsubscribe_links):
            try:
                result = self.unsubscribe_via_link(link)
                results[f'link_{i + 1}'] = result
                self.db_manager.record_unsubscribe(unsubscribe_info, result, f'web_link_{i + 1}', link)
                if 'success' in result.lower():
                    break

            except Exception as e:
                results[f'link_{i + 1}'] = f"Error: {str(e)}"

        return results

    def unsubscribe_via_email(self, email_address: str) -> str:
        """Attempt to unsubscribe via email (placeholder - would need SMTP setup)"""
        return f"Email unsubscribe request would be sent to: {email_address}"

    def unsubscribe_via_link(self, link: str) -> str:
        """Attempt to unsubscribe via web link"""
        try:
            response = self.session.get(link, timeout=10, allow_redirects=True)

            if response.status_code == 200:
                if self.check_unsubscribe_success(response.text):
                    return f"Successfully unsubscribed via: {link}"
                else:
                    form_result = self.handle_unsubscribe_forms(response.text, link)
                    if form_result:
                        return form_result
                    else:
                        return f"Visited {link} but unclear if unsubscribed successfully"
            else:
                return f"Failed to visit {link} - Status Code: {response.status_code}"

        except requests.exceptions.Timeout:
            return f"Timeout accessing: {link}"
        except requests.exceptions.RequestException as e:
            return f"Error accessing {link}: {str(e)}"

    def check_unsubscribe_success(self, html_content: str) -> bool:
        """Check if HTML content indicates successful unsubscribe"""
        success_indicators = [
            'successfully unsubscribed',
            'removed from list',
            'unsubscribe successful',
            'email preferences updated',
            'subscription cancelled'
        ]

        content_lower = html_content.lower()
        return any(indicator in content_lower for indicator in success_indicators)

    def handle_unsubscribe_forms(self, html_content: str, base_url: str) -> Optional[str]:
        """Handle unsubscribe forms on the page"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            forms = soup.find_all('form')

            for form in forms:
                form_text = form.get_text().lower()
                if any(word in form_text for word in ['unsubscribe', 'remove', 'opt out']):
                    return f"Found unsubscribe form at {base_url} (auto-submission not implemented for safety)"

            return None
        except:
            return None


class ModernScrollableFrame(ttk.Frame):
    """A modern scrollable frame widget"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=ModernTheme.BACKGROUND)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="Card.TFrame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.bind_mousewheel()

    def bind_mousewheel(self):
        """Bind mousewheel to scroll"""

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind("<MouseWheel>", _on_mousewheel)


class MailMute(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MailMute")
        self.geometry("1200x800")
        self.configure(bg=ModernTheme.BACKGROUND)
        ModernTheme.configure_styles()

        try:
            self.iconbitmap("icon.ico")
        except:
            pass
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.provider_var = tk.StringVar(value="gmail")
        self.unsubscribe_infos = []
        self.analyzer = MailMuteAlgorithm()
        self.db_manager = DatabaseManager()
        self.unsubscriber = SmartUnsubscriber(self.db_manager)
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready to analyze your emails")
        self.imap_server = tk.StringVar()
        self.imap_port = tk.IntVar(value=993)
        self.use_ssl = tk.BooleanVar(value=True)

        self.create_main_interface()

    def create_main_interface(self):
        """Create the main interface with modern design"""
        main_container = ttk.Frame(self)
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        self.create_header(main_container)
        self.notebook = ttk.Notebook(main_container, style="Modern.TNotebook")
        self.notebook.pack(fill='both', expand=True, pady=(20, 0))

        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="📧 Email Analysis")
        self.create_main_tab(self.main_tab)

        self.history_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.history_tab, text="📊 History & Stats")
        self.create_history_tab(self.history_tab)

        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="⚙️ Settings")
        self.create_settings_tab(self.settings_tab)

        self.create_status_bar()

    def create_header(self, parent):
        """Create the application header"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill='x', pady=(0, 20))
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(fill='x')

        title_label = ttk.Label(
            title_frame,
            text="MailMute",
            style="Heading.TLabel"
        )
        title_label.pack(anchor='w')

        subtitle_label = ttk.Label(
            title_frame,
            text="Intelligently detects and unsubscribes from unwanted email lists.",
            style="Body.TLabel"
        )
        subtitle_label.pack(anchor='w', pady=(5, 0))

    def create_main_tab(self, parent):
        """Create main analysis tab with modern design"""
        scrollable_container = ModernScrollableFrame(parent)
        scrollable_container.pack(fill='both', expand=True)

        container = scrollable_container.scrollable_frame

        self.create_provider_card(container)

        self.create_credentials_card(container)
        self.create_analysis_options_card(container)

        self.create_action_buttons(container)
        self.create_results_card(container)

    def create_provider_card(self, parent):
        """Create email provider selection card"""
        card_frame = ttk.LabelFrame(parent, text="📨 Email Provider", padding=20)
        card_frame.pack(fill='x', pady=(0, 20))

        providers = [
            ("Gmail", "gmail", "🔐 Requires App Password"),
            ("Outlook", "outlook", "📧 Microsoft Account"),
            ("Yahoo", "yahoo", "🔑 Requires App Password"),
            ("Other IMAP", "other", "🛠️ Custom Settings")
        ]

        provider_frame = ttk.Frame(card_frame)
        provider_frame.pack(fill='x')

        for i, (text, value, desc) in enumerate(providers):
            col = i % 2
            row = i // 2

            radio_frame = ttk.Frame(provider_frame)
            radio_frame.grid(row=row, column=col, sticky='w', padx=20, pady=10)

            ttk.Radiobutton(
                radio_frame,
                text=text,
                variable=self.provider_var,
                value=value,
                command=self.on_provider_change
            ).pack(anchor='w')

            ttk.Label(radio_frame, text=desc, style="Body.TLabel").pack(anchor='w', padx=(20, 0))

    def create_credentials_card(self, parent):
        """Create credentials input card"""
        self.cred_frame = ttk.LabelFrame(parent, text="🔐 Email Credentials", padding=20)
        self.cred_frame.pack(fill='x', pady=(0, 20))

        email_frame = ttk.Frame(self.cred_frame)
        email_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(email_frame, text="Email Address:", style="Body.TLabel").pack(anchor='w')
        email_entry = ttk.Entry(
            email_frame,
            textvariable=self.username,
            width=50,
            style="Modern.TEntry",
            font=('Segoe UI', 11)
        )
        email_entry.pack(fill='x', pady=(5, 0))
        password_frame = ttk.Frame(self.cred_frame)
        password_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(password_frame, text="App Password:", style="Body.TLabel").pack(anchor='w')
        password_entry = ttk.Entry(
            password_frame,
            textvariable=self.password,
            width=50,
            show="*",
            style="Modern.TEntry",
            font=('Segoe UI', 11)
        )
        password_entry.pack(fill='x', pady=(5, 0))
        self.help_frame = ttk.Frame(self.cred_frame)
        self.help_frame.pack(fill='x')

        self.help_label = ttk.Label(
            self.help_frame,
            text="⚠️ Gmail requires App Password. Enable 2FA first, then generate App Password.",
            style="Body.TLabel",
            foreground=ModernTheme.WARNING
        )
        self.help_label.pack(anchor='w')

        help_button = ttk.Button(
            self.help_frame,
            text="📖 Setup Help",
            command=self.show_app_password_help,
            style="Modern.TButton"
        )
        help_button.pack(anchor='w', pady=(10, 0))

        self.create_imap_settings_card(parent)

    def create_imap_settings_card(self, parent):
        """Create custom IMAP settings card"""
        self.imap_frame = ttk.LabelFrame(parent, text="🛠️ Custom IMAP Settings", padding=20)

        settings_grid = ttk.Frame(self.imap_frame)
        settings_grid.pack(fill='x')
        ttk.Label(settings_grid, text="IMAP Server:", style="Body.TLabel").grid(
            row=0, column=0, sticky='w', pady=5
        )
        ttk.Entry(
            settings_grid,
            textvariable=self.imap_server,
            width=30,
            style="Modern.TEntry"
        ).grid(row=0, column=1, padx=(10, 0), pady=5, sticky='ew')
        ttk.Label(settings_grid, text="Port:", style="Body.TLabel").grid(
            row=1, column=0, sticky='w', pady=5
        )
        ttk.Entry(
            settings_grid,
            textvariable=self.imap_port,
            width=10,
            style="Modern.TEntry"
        ).grid(row=1, column=1, padx=(10, 0), pady=5, sticky='w')
        ttk.Checkbutton(
            settings_grid,
            text="Use SSL/TLS",
            variable=self.use_ssl
        ).grid(row=2, column=1, padx=(10, 0), pady=5, sticky='w')

        settings_grid.columnconfigure(1, weight=1)

    def create_analysis_options_card(self, parent):
        """Create analysis options card"""
        options_frame = ttk.LabelFrame(parent, text="🔍 Analysis Options", padding=20)
        options_frame.pack(fill='x', pady=(0, 20))
        grid_frame = ttk.Frame(options_frame)
        grid_frame.pack(fill='x')

        self.limit_var = tk.IntVar(value=50)
        self.confidence_threshold = tk.DoubleVar(value=0.5)
        ttk.Label(grid_frame, text="Email Limit:", style="Body.TLabel").grid(
            row=0, column=0, sticky='w', pady=10
        )

        limit_frame = ttk.Frame(grid_frame)
        limit_frame.grid(row=0, column=1, padx=(20, 0), pady=10, sticky='ew')

        ttk.Spinbox(
            limit_frame,
            from_=10,
            to=500,
            textvariable=self.limit_var,
            width=15,
            style="Modern.TEntry"
        ).pack(side='left')

        ttk.Label(
            limit_frame,
            text="Maximum emails to analyze",
            style="Body.TLabel"
        ).pack(side='left', padx=(10, 0))
        ttk.Label(grid_frame, text="Confidence Threshold:", style="Body.TLabel").grid(
            row=1, column=0, sticky='w', pady=10
        )

        confidence_frame = ttk.Frame(grid_frame)
        confidence_frame.grid(row=1, column=1, padx=(20, 0), pady=10, sticky='ew')

        self.confidence_scale = ttk.Scale(
            confidence_frame,
            from_=0.1,
            to=1.0,
            variable=self.confidence_threshold,
            orient='horizontal',
            length=200,
            command=self.update_confidence_label
        )
        self.confidence_scale.pack(side='left')

        self.confidence_label = ttk.Label(
            confidence_frame,
            text="0.50",
            style="Body.TLabel",
            font=('Segoe UI', 10, 'bold')
        )
        self.confidence_label.pack(side='left', padx=(10, 0))

        ttk.Label(
            confidence_frame,
            text="Higher = more selective",
            style="Body.TLabel"
        ).pack(side='left', padx=(10, 0))

        grid_frame.columnconfigure(1, weight=1)

    def create_action_buttons(self, parent):
        """Create action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill='x', pady=20)
        primary_frame = ttk.Frame(button_frame)
        primary_frame.pack(anchor='center')

        self.analyze_button = ttk.Button(
            primary_frame,
            text="🔍 Analyze Emails",
            command=self.start_analysis,
            style="Primary.TButton"
        )
        self.analyze_button.pack(side='left', padx=(0, 15))

        self.unsubscribe_button = ttk.Button(
            primary_frame,
            text="🚫 Smart Unsubscribe",
            command=self.start_unsubscribe,
            style="Success.TButton",
            state="disabled"
        )
        self.unsubscribe_button.pack(side='left')

    def create_results_card(self, parent):
        """Create results display card"""
        results_frame = ttk.LabelFrame(parent, text="📊 Analysis Results", padding=20)
        results_frame.pack(fill='both', expand=True, pady=(0, 20))
        info_frame = ttk.Frame(results_frame)
        info_frame.pack(fill='x', pady=(0, 15))

        self.results_info_label = ttk.Label(
            info_frame,
            text="Run analysis to see unsubscribe opportunities",
            style="Body.TLabel"
        )
        self.results_info_label.pack(anchor='w')
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill='both', expand=True)
        columns = ('sender', 'subject', 'confidence', 'links', 'date')
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            height=12
        )
        self.tree.heading('sender', text='📤 Sender')
        self.tree.heading('subject', text='📄 Subject')
        self.tree.heading('confidence', text='🎯 Confidence')
        self.tree.heading('links', text='🔗 Links')
        self.tree.heading('date', text='📅 Date')

        self.tree.column('sender', width=200, minwidth=150)
        self.tree.column('subject', width=300, minwidth=200)
        self.tree.column('confidence', width=100, minwidth=80, anchor='center')
        self.tree.column('links', width=80, minwidth=60, anchor='center')
        self.tree.column('date', width=120, minwidth=100, anchor='center')
        v_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)

        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        selection_frame = ttk.Frame(results_frame)
        selection_frame.pack(fill='x', pady=(15, 0))

        ttk.Label(
            selection_frame,
            text="💡 Tip: Select rows to unsubscribe from specific senders",
            style="Body.TLabel"
        ).pack(anchor='w')

    def create_history_tab(self, parent):
        """Create modern history and statistics tab"""
        scrollable_container = ModernScrollableFrame(parent)
        scrollable_container.pack(fill='both', expand=True, padx=20, pady=20)

        container = scrollable_container.scrollable_frame
        self.create_stats_cards(container)
        self.create_history_display(container)

    def create_stats_cards(self, parent):
        """Create statistics cards"""
        stats_frame = ttk.Frame(parent)
        stats_frame.pack(fill='x', pady=(0, 30))
        cards_frame = ttk.Frame(stats_frame)
        cards_frame.pack(fill='x')
        total_card = ttk.LabelFrame(cards_frame, text="📊 Total Unsubscribes", padding=20)
        total_card.grid(row=0, column=0, padx=(0, 15), pady=10, sticky='ew')

        self.total_unsubs_label = ttk.Label(
            total_card,
            text="0",
            font=('Segoe UI', 24, 'bold'),
            foreground=ModernTheme.PRIMARY
        )
        self.total_unsubs_label.pack()

        ttk.Label(total_card, text="Successful unsubscribes", style="Body.TLabel").pack()
        success_card = ttk.LabelFrame(cards_frame, text="✅ Success Rate", padding=20)
        success_card.grid(row=0, column=1, padx=15, pady=10, sticky='ew')

        self.success_rate_label = ttk.Label(
            success_card,
            text="0%",
            font=('Segoe UI', 24, 'bold'),
            foreground=ModernTheme.SUCCESS
        )
        self.success_rate_label.pack()

        ttk.Label(success_card, text="Unsubscribe success rate", style="Body.TLabel").pack()
        active_card = ttk.LabelFrame(cards_frame, text="📧 Most Active Sender", padding=20)
        active_card.grid(row=0, column=2, padx=(15, 0), pady=10, sticky='ew')

        self.active_sender_label = ttk.Label(
            active_card,
            text="None",
            font=('Segoe UI', 14, 'bold'),
            foreground=ModernTheme.WARNING
        )
        self.active_sender_label.pack()

        ttk.Label(active_card, text="Highest email volume", style="Body.TLabel").pack()

        cards_frame.columnconfigure(0, weight=1)
        cards_frame.columnconfigure(1, weight=1)
        cards_frame.columnconfigure(2, weight=1)

    def create_history_display(self, parent):
        """Create history display"""
        history_frame = ttk.LabelFrame(parent, text="📜 Recent History", padding=20)
        history_frame.pack(fill='both', expand=True)
        history_columns = ('sender', 'subject', 'date', 'status', 'method')
        self.history_tree = ttk.Treeview(
            history_frame,
            columns=history_columns,
            show='headings',
            height=15
        )
        self.history_tree.heading('sender', text='📤 Sender')
        self.history_tree.heading('subject', text='📄 Subject')
        self.history_tree.heading('date', text='📅 Date')
        self.history_tree.heading('status', text='🔄 Status')
        self.history_tree.heading('method', text='🛠️ Method')

        self.history_tree.column('sender', width=200)
        self.history_tree.column('subject', width=300)
        self.history_tree.column('date', width=150, anchor='center')
        self.history_tree.column('status', width=200)
        self.history_tree.column('method', width=100, anchor='center')

        history_scrollbar = ttk.Scrollbar(history_frame, orient='vertical', command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=history_scrollbar.set)

        self.history_tree.pack(side='left', fill='both', expand=True)
        history_scrollbar.pack(side='right', fill='y')
        self.load_history_data()

    def create_settings_tab(self, parent):
        """Create modern settings tab"""
        scrollable_container = ModernScrollableFrame(parent)
        scrollable_container.pack(fill='both', expand=True, padx=20, pady=20)

        container = scrollable_container.scrollable_frame
        self.create_app_settings(container)
        self.create_about_section(container)

    def create_app_settings(self, parent):
        """Create application settings"""
        settings_frame = ttk.LabelFrame(parent, text="⚙️ Application Settings", padding=20)
        settings_frame.pack(fill='x', pady=(0, 20))
        theme_frame = ttk.Frame(settings_frame)
        theme_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(theme_frame, text="🎨 Theme:", style="Body.TLabel").pack(anchor='w')
        theme_var = tk.StringVar(value="Modern Blue")
        theme_combo = ttk.Combobox(
            theme_frame,
            textvariable=theme_var,
            values=["Modern Blue", "Dark Mode (Coming Soon)", "Light Mode (Coming Soon)"],
            state="readonly",
            width=30
        )
        theme_combo.pack(anchor='w', pady=(5, 0))
        db_frame = ttk.Frame(settings_frame)
        db_frame.pack(fill='x', pady=(15, 0))

        ttk.Label(db_frame, text="🗄️ Database:", style="Body.TLabel").pack(anchor='w')

        db_info_frame = ttk.Frame(db_frame)
        db_info_frame.pack(fill='x', pady=(5, 0))

        ttk.Label(
            db_info_frame,
            text=f"Location: {os.path.abspath('unsubscribe_history.db')}",
            style="Body.TLabel"
        ).pack(anchor='w')

        db_button_frame = ttk.Frame(db_frame)
        db_button_frame.pack(anchor='w', pady=(10, 0))

        ttk.Button(
            db_button_frame,
            text="🗑️ Clear History",
            command=self.clear_history,
            style="Danger.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            db_button_frame,
            text="📁 Open Database Folder",
            command=lambda: os.startfile(os.path.dirname(os.path.abspath('unsubscribe_history.db'))),
            style="Modern.TButton"
        ).pack(side='left')

    def create_about_section(self, parent):
        """Create about section"""
        about_frame = ttk.LabelFrame(parent, text="ℹ️ About", padding=20)
        about_frame.pack(fill='x')

        about_text = """
MailMute

An intelligent email management tool that helps you automatically detect and unsubscribe from unwanted email lists using advanced pattern recognition and AI-powered analysis.

Features:
• AI-powered email analysis and pattern recognition
• Smart unsubscribe link detection
• Detailed statistics and history tracking
• Secure authentication with App Password support
• Multi-provider support (Gmail, Outlook, Yahoo, Custom IMAP)
• Confidence scoring for better accuracy

Developed with using Python, with Tkinter as frontend.

By William Costales
        """

        about_label = ttk.Label(
            about_frame,
            text=about_text.strip(),
            style="Body.TLabel",
            justify='left'
        )
        about_label.pack(anchor='w')

    def create_status_bar(self):
        """Create modern status bar"""
        status_frame = ttk.Frame(self, style="Card.TFrame")
        status_frame.pack(fill='x', side='bottom', padx=20, pady=(0, 20))
        status_content = ttk.Frame(status_frame)
        status_content.pack(fill='x', padx=15, pady=10)
        status_text_frame = ttk.Frame(status_content)
        status_text_frame.pack(side='left', fill='x', expand=True)

        ttk.Label(
            status_text_frame,
            text="Status:",
            style="Body.TLabel",
            font=('Segoe UI', 9, 'bold')
        ).pack(side='left')

        self.status_label = ttk.Label(
            status_text_frame,
            textvariable=self.status_var,
            style="Body.TLabel"
        )
        self.status_label.pack(side='left', padx=(10, 0))
        progress_frame = ttk.Frame(status_content)
        progress_frame.pack(side='right')

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            mode='determinate',
            length=200
        )
        self.progress_bar.pack(side='right')

    def update_confidence_label(self, value):
        """Update confidence threshold label"""
        self.confidence_label.config(text=f"{float(value):.2f}")

    def on_provider_change(self):
        """Handle provider selection change"""
        provider = self.provider_var.get()
        help_texts = {
            "gmail": (
                "⚠️ Gmail requires App Password. Enable 2FA first, then generate App Password.", ModernTheme.WARNING),
            "outlook": (
                "📧 Use your Microsoft account password or App Password if 2FA is enabled.", ModernTheme.PRIMARY),
            "yahoo": ("⚠️ Yahoo requires App Password for IMAP access.", ModernTheme.WARNING),
            "other": ("🛠️ Enter your email provider's IMAP settings below.", ModernTheme.PRIMARY)
        }

        text, color = help_texts.get(provider, ("", ModernTheme.ON_SURFACE_VARIANT))
        self.help_label.config(text=text, foreground=color)
        if provider == "other":
            self.imap_frame.pack(fill='x', pady=(0, 20), after=self.cred_frame)
        else:
            self.imap_frame.pack_forget()

    def show_app_password_help(self):
        """Show detailed help for setting up App Password"""
        help_window = tk.Toplevel(self)
        help_window.title("Gmail App Password Setup")
        help_window.geometry("700x600")
        help_window.configure(bg=ModernTheme.BACKGROUND)
        help_window.resizable(False, False)
        help_window.transient(self)
        help_window.grab_set()
        help_frame = ttk.Frame(help_window, padding=30)
        help_frame.pack(fill='both', expand=True)
        ttk.Label(
            help_frame,
            text="📖 Gmail App Password Setup Guide",
            font=('Segoe UI', 16, 'bold'),
            foreground=ModernTheme.ON_SURFACE
        ).pack(anchor='w', pady=(0, 20))

        help_text = """
Gmail App Password Setup (Updated 2024/2025)

Gmail requires App Passwords for IMAP access. Google has made these harder to find!

STEP 1: Enable 2-Factor Authentication (Required First!)
• Go to myaccount.google.com/security
• Find "2-Step Verification" and click "Get started"  
• Complete the 2FA setup (SMS, authenticator app, etc.)
• THIS IS REQUIRED - App passwords won't appear without 2FA

STEP 2: Find App Passwords (Google hid this!)
Method A - Search:
• Go to myaccount.google.com
• Use the SEARCH BOX at the top and type "App passwords"
• Click the "App passwords" result

Method B - Direct link:
• Go directly to: myaccount.google.com/apppasswords

Method C - If still can't find:
• Go to myaccount.google.com/security
• Scroll down and look for "App passwords" (may be hidden)

STEP 3: Generate App Password
• Click "Generate" or "Create app password"
• Select "Mail" or choose "Other" and type "Email Client"
• Copy the 16-character password (like: abcd efgh ijkl mnop)

STEP 4: Use in This App
• Email: your-email@gmail.com
• App Password: the 16-character code (NOT your regular password)

COMMON ISSUES:
• "App passwords" section missing = 2FA not enabled yet
• "Invalid credentials" = using regular password instead of App password
• Some Google Workspace accounts have different settings

WHY APP PASSWORDS:
• More secure than regular passwords
• Can be revoked individually
• Limited access scope
• Google's requirement for third-party email apps
        """
        text_frame = ttk.Frame(help_frame)
        text_frame.pack(fill='both', expand=True, pady=(0, 20))

        text_widget = scrolledtext.ScrolledText(
            text_frame,
            wrap=tk.WORD,
            padx=15,
            pady=15,
            font=('Segoe UI', 10),
            height=20
        )
        text_widget.pack(fill='both', expand=True)
        text_widget.insert('1.0', help_text.strip())
        text_widget.config(state='disabled')
        button_frame = ttk.Frame(help_frame)
        button_frame.pack(fill='x')

        ttk.Button(
            button_frame,
            text="🌐 Open Google Account",
            command=lambda: webbrowser.open("https://myaccount.google.com"),
            style="Modern.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            button_frame,
            text="🔐 Direct App Passwords",
            command=lambda: webbrowser.open("https://myaccount.google.com/apppasswords"),
            style="Modern.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            button_frame,
            text="🔒 Enable 2FA",
            command=lambda: webbrowser.open("https://myaccount.google.com/security"),
            style="Modern.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            button_frame,
            text="Close",
            command=help_window.destroy,
            style="Primary.TButton"
        ).pack(side='right')

    def start_analysis(self):
        """Start email analysis with validation"""
        if not self.username.get().strip() or not self.password.get().strip():
            messagebox.showerror("Error", "Please enter email credentials")
            return

        self.analyze_button.config(state="disabled")
        self.unsubscribe_button.config(state="disabled")
        threading.Thread(target=self.analyze_emails, daemon=True).start()

    def analyze_emails(self):
        """Analyze emails for unsubscribe opportunities"""
        try:
            self.status_var.set("🔌 Connecting to email server...")
            self.progress_var.set(5)
            mail = self.connect_to_email_server()
            if not mail:
                return

            self.status_var.set("🔍 Searching for marketing emails...")
            self.progress_var.set(15)
            search_criteria = [
                '(TEXT "unsubscribe")',
                '(TEXT "newsletter")',
                '(TEXT "marketing")',
                '(FROM "noreply")',
                '(FROM "no-reply")'
            ]

            all_email_ids = set()
            for i, criteria in enumerate(search_criteria):
                try:
                    self.status_var.set(f"🔍 Searching... ({i + 1}/{len(search_criteria)})")
                    _, search_data = mail.search(None, criteria)
                    if search_data[0]:
                        all_email_ids.update(search_data[0].split())
                    self.progress_var.set(15 + (i + 1) * 5)
                except:
                    continue

            email_ids = list(all_email_ids)[:self.limit_var.get()]

            if not email_ids:
                self.status_var.set("❌ No marketing emails found")
                self.after(0, lambda: messagebox.showinfo(
                    "No Results",
                    "No marketing emails found. Try adjusting your search criteria or check a different email folder."
                ))
                return

            self.status_var.set("🤖 Analyzing emails with AI...")
            self.unsubscribe_infos = []

            for i, num in enumerate(email_ids):
                try:
                    _, data = mail.fetch(num, "(RFC822)")
                    raw_email = data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    unsubscribe_info = self.analyzer.analyze_email_content(msg)

                    if unsubscribe_info.confidence_score >= self.confidence_threshold.get():
                        self.unsubscribe_infos.append(unsubscribe_info)

                        self.db_manager.update_sender_stats(unsubscribe_info.sender)

                    progress = 40 + (i / len(email_ids)) * 50
                    self.progress_var.set(progress)
                    self.status_var.set(f"🤖 Analyzed {i + 1}/{len(email_ids)} emails...")

                except Exception as e:
                    print(f"Error analyzing email {num}: {e}")
                    continue

            self.after(0, self.update_results_display)

            mail.logout()
            self.status_var.set(f"✅ Analysis complete! Found {len(self.unsubscribe_infos)} unsubscribe opportunities")
            self.progress_var.set(100)
            if self.unsubscribe_infos:
                self.after(0, lambda: self.unsubscribe_button.config(state="normal"))

        except Exception as e:
            self.status_var.set(f"❌ Error: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Analysis Error", f"Analysis failed: {str(e)}"))
        finally:
            self.after(0, lambda: self.analyze_button.config(state="normal"))

    def connect_to_email_server(self):
        """Connect to email server with support for multiple providers"""
        try:
            imap_settings = self.get_imap_settings()
            if not imap_settings:
                return None

            server, port, use_ssl = imap_settings
            if use_ssl:
                mail = imaplib.IMAP4_SSL(server, port)
            else:
                mail = imaplib.IMAP4(server, port)

            mail.login(self.username.get().strip(), self.password.get().strip())
            mail.select("inbox")
            return mail

        except imaplib.IMAP4.error as e:
            error_msg = str(e)
            provider = self.provider_var.get()

            if "Application-specific password" in error_msg or provider == "gmail":
                self.after(0, lambda: messagebox.showerror(
                    "🔐 Authentication Error",
                    "Gmail requires an App Password!\n\n"
                    "Quick Fix:\n"
                    "1. Enable 2-Factor Authentication on your Google Account\n"
                    "2. Generate an App Password for Mail\n"
                    "3. Use the 16-character App Password here\n\n"
                    "Click 'Setup Help' button for detailed instructions."
                ))
            elif "Invalid credentials" in error_msg:
                provider_help = {
                    "gmail": "Use your Gmail App Password (16 characters)",
                    "outlook": "Use your Microsoft account password or App Password if 2FA is enabled",
                    "yahoo": "Use your Yahoo App Password (required for IMAP)",
                    "other": "Check with your email provider for correct credentials"
                }

                self.after(0, lambda: messagebox.showerror(
                    "🔑 Invalid Credentials",
                    f"Authentication failed for {provider.title()}.\n\n"
                    f"💡 Tip: {provider_help.get(provider, 'Check your credentials')}\n\n"
                    f"Make sure you're using:\n"
                    f"• Your complete email address\n"
                    f"• The correct password/App Password"
                ))
            else:
                self.after(0, lambda: messagebox.showerror("Connection Error", f"Failed to connect: {error_msg}"))

            self.status_var.set("❌ Connection failed")
            self.progress_var.set(0)
            return None
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Unexpected error: {str(e)}"))
            self.status_var.set("❌ Connection failed")
            self.progress_var.set(0)
            return None

    def get_imap_settings(self):
        """Get IMAP settings based on selected provider"""
        provider = self.provider_var.get()

        settings = {
            "gmail": ("imap.gmail.com", 993, True),
            "outlook": ("outlook.office365.com", 993, True),
            "yahoo": ("imap.mail.yahoo.com", 993, True),
            "other": (self.imap_server.get(), self.imap_port.get(), self.use_ssl.get())
        }

        if provider == "other" and not self.imap_server.get():
            messagebox.showerror("Error", "Please enter IMAP server details")
            return None

        return settings.get(provider)

    def update_results_display(self):
        """Update the results treeview with modern styling"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.unsubscribe_infos:
            self.results_info_label.config(text="No unsubscribe opportunities found with current settings")
            return
        self.results_info_label.config(
            text=f"Found {len(self.unsubscribe_infos)} unsubscribe opportunities • "
                 f"Select rows and click 'Smart Unsubscribe' to proceed"
        )

        for info in self.unsubscribe_infos:
            subject = info.subject[:60] + "..." if len(info.subject) > 60 else info.subject
            confidence = f"{info.confidence_score:.0%}"
            date_str = info.email_date.strftime("%m/%d/%Y")

            item = self.tree.insert('', 'end', values=(
                info.sender,
                subject,
                confidence,
                len(info.unsubscribe_links),
                date_str
            ))
            if info.confidence_score >= 0.8:
                self.tree.set(item, 'confidence', f"🟢 {confidence}")
            elif info.confidence_score >= 0.6:
                self.tree.set(item, 'confidence', f"🟡 {confidence}")
            else:
                self.tree.set(item, 'confidence', f"🟠 {confidence}")

    def start_unsubscribe(self):
        """Start smart unsubscribe process"""
        if not self.unsubscribe_infos:
            messagebox.showinfo("Info", "Please analyze emails first")
            return

        selected_items = self.tree.selection()
        if not selected_items:
            result = messagebox.askyesno(
                "Unsubscribe All?",
                f"No specific emails selected.\n\n"
                f"Do you want to unsubscribe from all {len(self.unsubscribe_infos)} detected opportunities?",
                icon="question"
            )
            if result:
                selected_items = self.tree.get_children()
            else:
                return
        count = len(selected_items)
        result = messagebox.askyesno(
            "Confirm Unsubscribe",
            f"Are you sure you want to attempt unsubscribing from {count} email list(s)?\n\n"
            f"This action will:\n"
            f"• Visit unsubscribe links automatically\n"
            f"• Log all attempts in your history\n"
            f"• May take several minutes to complete",
            icon="warning"
        )

        if not result:
            return

        self.unsubscribe_button.config(state="disabled")
        self.analyze_button.config(state="disabled")

        threading.Thread(target=self.perform_unsubscribe, args=(selected_items,), daemon=True).start()

    def perform_unsubscribe(self, selected_items):
        results = []
        successful_count = 0

        total_items = len(selected_items)

        for i, item in enumerate(selected_items):
            try:
                item_index = self.tree.index(item)
                unsubscribe_info = self.unsubscribe_infos[item_index]

                self.status_var.set(f"🚫 Unsubscribing from {unsubscribe_info.sender}... ({i + 1}/{total_items})")

                result = self.unsubscriber.unsubscribe(unsubscribe_info)

                success = any('success' in str(r).lower() for r in result.values())
                if success:
                    successful_count += 1

                result_summary = []
                for method, outcome in result.items():
                    result_summary.append(f"  {method}: {outcome}")

                results.append({
                    'sender': unsubscribe_info.sender,
                    'success': success,
                    'details': '\n'.join(result_summary)
                })

            except Exception as e:
                results.append({
                    'sender': unsubscribe_info.sender if 'unsubscribe_info' in locals() else 'Unknown',
                    'success': False,
                    'details': f"Error: {str(e)}"
                })

            progress = (i + 1) / total_items * 100
            self.progress_var.set(progress)
        self.after(0, lambda: self.show_unsubscribe_results(results, successful_count, total_items))
        self.status_var.set(f"✅ Unsubscribe process completed: {successful_count}/{total_items} successful")
        self.progress_var.set(100)
        self.after(0, lambda: self.analyze_button.config(state="normal"))
        self.after(0, lambda: self.unsubscribe_button.config(state="normal"))

    def show_unsubscribe_results(self, results, successful_count, total_count):
        """Show detailed unsubscribe results in a modern dialog"""
        results_window = tk.Toplevel(self)
        results_window.title("🚫 Unsubscribe Results")
        results_window.geometry("800x600")
        results_window.configure(bg=ModernTheme.BACKGROUND)
        results_window.transient(self)

        header_frame = ttk.Frame(results_window, padding=20)
        header_frame.pack(fill='x')
        success_rate = (successful_count / total_count * 100) if total_count > 0 else 0

        summary_frame = ttk.Frame(header_frame)
        summary_frame.pack(fill='x', pady=(0, 20))

        ttk.Label(
            summary_frame,
            text="📊 Unsubscribe Summary",
            font=('Segoe UI', 16, 'bold'),
            foreground=ModernTheme.ON_SURFACE
        ).pack(anchor='w')

        stats_frame = ttk.Frame(summary_frame)
        stats_frame.pack(fill='x', pady=(10, 0))

        success_label = ttk.Label(
            stats_frame,
            text=f"✅ Successful: {successful_count}/{total_count} ({success_rate:.1f}%)",
            font=('Segoe UI', 12, 'bold'),
            foreground=ModernTheme.SUCCESS if successful_count > 0 else ModernTheme.DANGER
        )
        success_label.pack(anchor='w')
        results_frame = ttk.LabelFrame(results_window, text="📋 Detailed Results", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20)

        results_text = scrolledtext.ScrolledText(
            results_frame,
            wrap=tk.WORD,
            width=90,
            height=20,
            font=('Consolas', 9),
            padx=10,
            pady=10
        )
        results_text.pack(fill='both', expand=True, pady=(0, 15))

        for i, result in enumerate(results, 1):
            status_icon = "✅" if result['success'] else "❌"
            status_text = "SUCCESS" if result['success'] else "FAILED"

            results_text.insert(tk.END, f"{i}. {status_icon} {result['sender']} - {status_text}\n")
            results_text.insert(tk.END, f"{result['details']}\n\n")

        results_text.configure(state="disabled")

        button_frame = ttk.Frame(results_window, padding=20)
        button_frame.pack(fill='x')

        ttk.Button(
            button_frame,
            text="📄 Export Results",
            command=lambda: self.export_results(results),
            style="Modern.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            button_frame,
            text="🔄 Refresh Analysis",
            command=lambda: [results_window.destroy(), self.start_analysis()],
            style="Modern.TButton"
        ).pack(side='left', padx=(0, 10))

        ttk.Button(
            button_frame,
            text="Close",
            command=results_window.destroy,
            style="Primary.TButton"
        ).pack(side='right')

    def export_results(self, results):
        """Export unsubscribe results to file"""
        from tkinter import filedialog

        filename = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )

        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("AI Email Unsubscriber Results\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

                    for i, result in enumerate(results, 1):
                        status = "SUCCESS" if result['success'] else "FAILED"
                        f.write(f"{i}. {result['sender']} - {status}\n")
                        f.write(f"{result['details']}\n\n")

                messagebox.showinfo("Export Successful", f"Results exported to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Export Failed", f"Failed to export results:\n{str(e)}")

    def load_history_data(self):
        """Load and display history data"""
        try:

            history_data = self.db_manager.get_history(100)

            for item in self.history_tree.get_children():
                self.history_tree.delete(item)

            for record in history_data:
                sender, subject, date, status, method, notes = record

                try:
                    if isinstance(date, str):
                        date_obj = datetime.fromisoformat(date)
                    else:
                        date_obj = date
                    formatted_date = date_obj.strftime("%m/%d/%Y %H:%M")
                except:
                    formatted_date = str(date)

                short_subject = subject[:50] + "..." if len(subject) > 50 else subject
                status_icon = "✅" if "success" in status.lower() else "❌"
                status_display = f"{status_icon} {status}"

                self.history_tree.insert('', 'end', values=(
                    sender,
                    short_subject,
                    formatted_date,
                    status_display,
                    method
                ))

            self.update_statistics()

        except Exception as e:
            print(f"Error loading history: {e}")

    def update_statistics(self):
        """Update statistics display"""
        try:

            history_data = self.db_manager.get_history(1000)
            sender_stats = self.db_manager.get_sender_stats()

            total_attempts = len(history_data)
            successful_attempts = sum(1 for record in history_data if "success" in record[3].lower())
            success_rate = (successful_attempts / total_attempts * 100) if total_attempts > 0 else 0

            self.total_unsubs_label.config(text=str(successful_attempts))
            self.success_rate_label.config(text=f"{success_rate:.1f}%")

            if sender_stats:
                most_active = sender_stats[0][0]
                self.active_sender_label.config(text=most_active[:30] + "..." if len(most_active) > 30 else most_active)
            else:
                self.active_sender_label.config(text="None")

        except Exception as e:
            print(f"Error updating statistics: {e}")

    def clear_history(self):
        """Clear all history data"""
        result = messagebox.askyesno(
            "Clear History",
            "Are you sure you want to clear all unsubscribe history?\n\n"
            "This action cannot be undone.",
            icon="warning"
        )

        if result:
            try:

                conn = sqlite3.connect(self.db_manager.db_path)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM unsubscribe_history")
                cursor.execute("DELETE FROM sender_stats")
                conn.commit()
                conn.close()

                self.load_history_data()

                messagebox.showinfo("History Cleared", "All history data has been cleared successfully.")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to clear history:\n{str(e)}")


if __name__ == "__main__":
    try:
        app = MailMute()
        app.mainloop()
    except Exception as e:
        print(f"Application error: {e}")
        import traceback

        traceback.print_exc()
