"""
EPUB Editor - A desktop application for creating, editing, and exporting E-Books.

Required pip packages:
    pip install customtkinter ebooklib Pillow python-docx reportlab beautifulsoup4 lxml

Usage:
    python epub_editor.py
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import json
import os
import base64
import html as html_module
from pathlib import Path
from PIL import Image, ImageTk
import io
import threading

# ── EPUB ──────────────────────────────────────────────────────────────────────
try:
    from ebooklib import epub
    import ebooklib
    EPUB_OK = True
except ImportError:
    EPUB_OK = False

# ── BeautifulSoup (EPUB import) ───────────────────────────────────────────────
try:
    from bs4 import BeautifulSoup
    BS4_OK = True
except ImportError:
    BS4_OK = False

# ── PDF ───────────────────────────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ── DOCX ──────────────────────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_OK = True
except ImportError:
    DOCX_OK = False


# ══════════════════════════════════════════════════════════════════════════════
#  COLOUR PALETTE & THEME
# ══════════════════════════════════════════════════════════════════════════════
COLORS = {
    "bg":           "#0F1117",   # near-black background
    "surface":      "#1A1D27",   # panel / card surface
    "surface2":     "#22263A",   # slightly lighter surface
    "border":       "#2E3347",   # subtle borders
    "accent":       "#5C6BFF",   # electric indigo accent
    "accent_hover": "#7B89FF",
    "accent_dim":   "#3A4599",
    "text":         "#E8EAF6",   # primary text
    "text_dim":     "#8B90A8",   # secondary / muted text
    "danger":       "#FF5370",
    "success":      "#4CAF87",
    "chapter_sel":  "#2A2F4E",   # selected chapter row
}

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

FONT_HEADING  = ("Segoe UI", 18, "bold")
FONT_SUBHEAD  = ("Segoe UI", 13, "bold")
FONT_LABEL    = ("Segoe UI", 11)
FONT_SMALL    = ("Segoe UI", 10)
FONT_MONO     = ("Consolas", 11)
FONT_CHAPTER  = ("Segoe UI", 12)
FONT_EDITOR   = ("Georgia", 13)


# ══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
class Chapter:
    """Holds a single chapter's name and plain-text content."""
    def __init__(self, name: str = "", content: str = ""):
        self.name    = name
        self.content = content

    def to_dict(self) -> dict:
        return {"name": self.name, "content": self.content}

    @classmethod
    def from_dict(cls, d: dict) -> "Chapter":
        return cls(d.get("name", ""), d.get("content", ""))


class Book:
    """Complete book project."""
    def __init__(self):
        self.title    : str           = ""
        self.author   : str           = ""
        self.cover_b64: str           = ""   # base64-encoded image
        self.chapters : list[Chapter] = []

    # ── serialisation ────────────────────────────────────────────────────────
    def to_dict(self) -> dict:
        return {
            "title":    self.title,
            "author":   self.author,
            "cover":    self.cover_b64,
            "chapters": [c.to_dict() for c in self.chapters],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Book":
        b          = cls()
        b.title    = d.get("title",  "")
        b.author   = d.get("author", "")
        b.cover_b64= d.get("cover",  "")
        b.chapters = [Chapter.from_dict(c) for c in d.get("chapters", [])]
        return b

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "Book":
        with open(path, "r", encoding="utf-8") as f:
            return cls.from_dict(json.load(f))

    # ── next chapter name ─────────────────────────────────────────────────────
    def next_chapter_name(self) -> str:
        return f"Kapitel {len(self.chapters) + 1}"


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _text_to_html(text: str) -> str:
    """Convert plain text to simple HTML paragraphs."""
    paragraphs = text.split("\n\n")
    parts = []
    for p in paragraphs:
        p = p.strip()
        if p:
            escaped = html_module.escape(p).replace("\n", "<br/>")
            parts.append(f"<p>{escaped}</p>")
    return "\n".join(parts) if parts else "<p></p>"



def export_epub(book: Book, path: str) -> None:
    """Export the Book as EPUB."""
    if not EPUB_OK:
        raise RuntimeError("ebooklib is not installed.")

    eb = epub.EpubBook()
    eb.set_identifier("bookid")
    eb.set_title(book.title or "Untitled")
    eb.set_language("de")
    eb.add_author(book.author or "Unknown")

    chapters = []
    for i, ch in enumerate(book.chapters):
        content = (ch.content or "").strip()
        if not content:
            content = " "

        chapter = epub.EpubHtml(
            title=ch.name or f"Kapitel {i+1}",
            file_name=f"chapter_{i+1}.xhtml",
            lang="de"
        )

        chapter.content = f"""
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head><title>{html_module.escape(ch.name or f'Kapitel {i+1}')}</title></head>
        <body>
        <h1>{html_module.escape(ch.name or f'Kapitel {i+1}')}</h1>
        <p>{html_module.escape(content).replace(chr(10), '<br/>')}</p>
        </body>
        </html>
        """

        eb.add_item(chapter)
        chapters.append(chapter)

    if not chapters:
        raise RuntimeError("Keine Kapitel vorhanden.")

    eb.toc = tuple(chapters)
    eb.spine = ["nav"] + chapters

    eb.add_item(epub.EpubNcx())
    eb.add_item(epub.EpubNav())

    epub.write_epub(path, eb)

def export_pdf(book: Book, path: str) -> None:
    """Export the Book as a PDF using ReportLab."""
    if not PDF_OK:
        raise RuntimeError("reportlab is not installed.")

    doc = SimpleDocTemplate(path, pagesize=A4,
                            leftMargin=3*cm, rightMargin=3*cm,
                            topMargin=3*cm, bottomMargin=3*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BookTitle", parent=styles["Title"],
        fontSize=28, spaceAfter=12, alignment=TA_CENTER
    )
    author_style = ParagraphStyle(
        "BookAuthor", parent=styles["Normal"],
        fontSize=14, spaceAfter=40, alignment=TA_CENTER, textColor="#555555"
    )
    ch_title_style = ParagraphStyle(
        "ChTitle", parent=styles["Heading1"],
        fontSize=18, spaceBefore=30, spaceAfter=12
    )
    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"],
        fontSize=11, leading=18, spaceAfter=8
    )

    story = []
    story.append(Paragraph(html_module.escape(book.title or "Untitled"), title_style))
    story.append(Paragraph(html_module.escape(book.author or ""), author_style))

    for ch in book.chapters:
        story.append(Paragraph(html_module.escape(ch.name), ch_title_style))
        for para in ch.content.split("\n\n"):
            para = para.strip()
            if para:
                story.append(Paragraph(html_module.escape(para).replace("\n", "<br/>"), body_style))
        story.append(Spacer(1, 12))

    doc.build(story)


def export_docx(book: Book, path: str) -> None:
    """Export the Book as a DOCX file."""
    if not DOCX_OK:
        raise RuntimeError("python-docx is not installed.")

    doc = Document()
    # Title page
    t = doc.add_heading(book.title or "Untitled", level=0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a = doc.add_paragraph(book.author or "")
    a.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    for ch in book.chapters:
        doc.add_heading(ch.name, level=1)
        for para in ch.content.split("\n\n"):
            para = para.strip()
            if para:
                doc.add_paragraph(para)
        doc.add_paragraph()

    doc.save(path)


def import_epub(path: str) -> Book:
    """Import an EPUB file into a Book object."""
    if not EPUB_OK:
        raise RuntimeError("ebooklib is not installed.")
    if not BS4_OK:
        raise RuntimeError("beautifulsoup4 is not installed.")

    eb   = epub.read_epub(path)
    book = Book()

    # ── metadata ─────────────────────────────────────────────────────────────
    titles  = eb.get_metadata("DC", "title")
    authors = eb.get_metadata("DC", "creator")
    book.title  = titles[0][0]  if titles  else ""
    book.author = authors[0][0] if authors else ""

    # ── cover ─────────────────────────────────────────────────────────────────
    for item in eb.get_items_of_type(ebooklib.ITEM_IMAGE):
        name = item.file_name.lower()
        if "cover" in name or name.endswith((".jpg", ".jpeg", ".png")):
            book.cover_b64 = base64.b64encode(item.content).decode()
            break

    # ── chapters ─────────────────────────────────────────────────────────────
    for item in eb.get_items_of_type(ebooklib.ITEM_DOCUMENT):
        soup  = BeautifulSoup(item.content, "lxml")
        # Extract heading as chapter name
        h_tag = soup.find(["h1", "h2", "h3"])
        name  = h_tag.get_text(strip=True) if h_tag else item.get_name()
        # Remove heading from body text
        if h_tag:
            h_tag.decompose()
        # Build plain text
        lines = []
        for elem in soup.find_all(["p", "div"]):
            t = elem.get_text(" ", strip=True)
            if t:
                lines.append(t)
        content = "\n\n".join(lines)
        if content.strip() or name:
            book.chapters.append(Chapter(name=name, content=content))

    return book


# ══════════════════════════════════════════════════════════════════════════════
#  UI WIDGETS
# ══════════════════════════════════════════════════════════════════════════════

class ChapterRow(ctk.CTkFrame):
    """A single row in the chapter sidebar."""

    def __init__(self, master, chapter: Chapter, index: int,
                 on_select, on_rename, on_delete, on_move_up, on_move_down,
                 selected: bool = False, **kwargs):
        super().__init__(
            master,
            fg_color=COLORS["chapter_sel"] if selected else COLORS["surface"],
            corner_radius=8,
            **kwargs,
        )
        self._chapter     = chapter
        self._index       = index
        self._on_select   = on_select
        self._on_rename   = on_rename
        self._on_delete   = on_delete
        self._on_move_up  = on_move_up
        self._on_move_down= on_move_down

        self.grid_columnconfigure(0, weight=1)

        # ── label (clickable) ─────────────────────────────────────────────
        self._label = ctk.CTkLabel(
            self, text=chapter.name, anchor="w",
            text_color=COLORS["text"] if selected else COLORS["text_dim"],
            font=FONT_CHAPTER, padx=10,
        )
        self._label.grid(row=0, column=0, sticky="ew", pady=6)
        self._label.bind("<Button-1>", lambda _: on_select(index))
        self.bind("<Button-1>", lambda _: on_select(index))

        # ── action buttons (appear on hover) ──────────────────────────────
        self._btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._btn_frame.grid(row=0, column=1, padx=(0, 6))

        btn_cfg = dict(width=24, height=24, corner_radius=4,
                       fg_color="transparent",
                       text_color=COLORS["text_dim"],
                       hover_color=COLORS["surface2"])

        ctk.CTkButton(self._btn_frame, text="↑", **btn_cfg,
                      command=lambda: on_move_up(index)).pack(side="left", padx=1)
        ctk.CTkButton(self._btn_frame, text="↓", **btn_cfg,
                      command=lambda: on_move_down(index)).pack(side="left", padx=1)
        ctk.CTkButton(self._btn_frame, text="✎", **btn_cfg,
                      command=lambda: on_rename(index)).pack(side="left", padx=1)
        ctk.CTkButton(self._btn_frame, text="✕",
                      **{**btn_cfg, "text_color": COLORS["danger"],
                         "hover_color": "#3a1520"},
                      command=lambda: on_delete(index)).pack(side="left", padx=1)


class StatusBar(ctk.CTkFrame):
    """Thin status bar at the bottom of the window."""

    def __init__(self, master, **kwargs):
        super().__init__(master, height=28, fg_color=COLORS["surface"],
                         corner_radius=0, **kwargs)
        self._label = ctk.CTkLabel(
            self, text="Bereit.", anchor="w",
            text_color=COLORS["text_dim"], font=FONT_SMALL,
        )
        self._label.pack(side="left", padx=12)

    def set(self, msg: str) -> None:
        self._label.configure(text=msg)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════

class EPUBEditorApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self._book          : Book           = Book()
        self._selected_idx  : int | None     = None
        self._project_path  : str | None     = None
        self._cover_img_tk  : ImageTk.PhotoImage | None = None

        self._configure_window()
        self._build_ui()
        self._refresh_chapter_list()
        self._status.set("Neues Projekt gestartet. Bereit.")

    # ── window setup ─────────────────────────────────────────────────────────
    def _configure_window(self) -> None:
        self.title("EPUB Editor")
        self.geometry("1280x820")
        self.minsize(900, 600)
        self.configure(fg_color=COLORS["bg"])

        # menubar
        menubar = tk.Menu(self, bg=COLORS["surface"], fg=COLORS["text"],
                          activebackground=COLORS["accent"],
                          activeforeground=COLORS["text"],
                          relief="flat", bd=0)
        self.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0,
                            bg=COLORS["surface"], fg=COLORS["text"],
                            activebackground=COLORS["accent"],
                            activeforeground=COLORS["text"])
        menubar.add_cascade(label=" Datei ", menu=file_menu)
        file_menu.add_command(label="Neues Projekt",       command=self._new_project)
        file_menu.add_command(label="Projekt öffnen…",    command=self._open_project)
        file_menu.add_command(label="Projekt speichern",   command=self._save_project)
        file_menu.add_command(label="Projekt speichern unter…", command=self._save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="EPUB importieren…",  command=self._import_epub)
        file_menu.add_separator()
        file_menu.add_command(label="Beenden",             command=self.quit)

        export_menu = tk.Menu(menubar, tearoff=0,
                              bg=COLORS["surface"], fg=COLORS["text"],
                              activebackground=COLORS["accent"],
                              activeforeground=COLORS["text"])
        menubar.add_cascade(label=" Exportieren ", menu=export_menu)
        export_menu.add_command(label="Als EPUB exportieren…",  command=self._export_epub)
        export_menu.add_command(label="Als PDF exportieren…",   command=self._export_pdf)
        export_menu.add_command(label="Als DOCX exportieren…",  command=self._export_docx)

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── top toolbar ───────────────────────────────────────────────────
        self._build_toolbar()

        # ── main content area ─────────────────────────────────────────────
        content = ctk.CTkFrame(self, fg_color=COLORS["bg"])
        content.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)

        # LEFT: sidebar
        self._build_sidebar(content)

        # MIDDLE divider
        div = ctk.CTkFrame(content, width=1, fg_color=COLORS["border"])
        div.grid(row=0, column=1, sticky="ns", padx=0)

        # RIGHT: editor
        self._build_editor(content)

        # ── status bar ────────────────────────────────────────────────────
        self._status = StatusBar(self)
        self._status.grid(row=2, column=0, sticky="ew")

    def _build_toolbar(self) -> None:
        bar = ctk.CTkFrame(self, height=60, fg_color=COLORS["surface"],
                           corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_columnconfigure(99, weight=1)  # push everything left

        # App logo / name
        ctk.CTkLabel(bar, text="  📖  EPUB Editor",
                     font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["text"]).grid(row=0, column=0, padx=16, pady=12)

        # Separator
        ctk.CTkFrame(bar, width=1, height=30, fg_color=COLORS["border"]).grid(
            row=0, column=1, padx=8)

        btn_style = dict(height=36, corner_radius=8, font=FONT_LABEL,
                         fg_color=COLORS["surface2"],
                         hover_color=COLORS["border"],
                         text_color=COLORS["text"])

        ctk.CTkButton(bar, text="＋ Neu",    width=90,  **btn_style,
                      command=self._new_project).grid(row=0, column=2, padx=4)
        ctk.CTkButton(bar, text="📂 Öffnen", width=110, **btn_style,
                      command=self._open_project).grid(row=0, column=3, padx=4)
        ctk.CTkButton(bar, text="💾 Speichern", width=120, **btn_style,
                      command=self._save_project).grid(row=0, column=4, padx=4)

        ctk.CTkFrame(bar, width=1, height=30, fg_color=COLORS["border"]).grid(
            row=0, column=5, padx=8)

        ctk.CTkButton(bar, text="📥 EPUB Import", width=140, **btn_style,
                      command=self._import_epub).grid(row=0, column=6, padx=4)

        acc_style = dict(height=36, corner_radius=8, font=FONT_LABEL,
                         fg_color=COLORS["accent"],
                         hover_color=COLORS["accent_hover"],
                         text_color="#ffffff")

        ctk.CTkButton(bar, text="📤 EPUB Export", width=140, **acc_style,
                      command=self._export_epub).grid(row=0, column=7, padx=(4, 16))

    def _build_sidebar(self, parent) -> None:
        sidebar = ctk.CTkFrame(parent, width=280, fg_color=COLORS["surface"],
                               corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_rowconfigure(3, weight=1)
        sidebar.grid_propagate(False)

        # ── Metadata section ──────────────────────────────────────────────
        meta = ctk.CTkFrame(sidebar, fg_color=COLORS["surface2"], corner_radius=10)
        meta.pack(fill="x", padx=12, pady=(16, 8))

        ctk.CTkLabel(meta, text="METADATA", font=FONT_SMALL,
                     text_color=COLORS["text_dim"]).pack(anchor="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(meta, text="Titel", font=FONT_SMALL,
                     text_color=COLORS["text_dim"]).pack(anchor="w", padx=12)
        self._title_var = ctk.StringVar()
        self._title_entry = ctk.CTkEntry(
            meta, textvariable=self._title_var,
            placeholder_text="Buchtitel …",
            fg_color=COLORS["border"], border_color=COLORS["border"],
            text_color=COLORS["text"], font=FONT_LABEL, height=34,
        )
        self._title_entry.pack(fill="x", padx=12, pady=(2, 6))
        self._title_var.trace_add("write", self._on_metadata_change)

        ctk.CTkLabel(meta, text="Autor", font=FONT_SMALL,
                     text_color=COLORS["text_dim"]).pack(anchor="w", padx=12)
        self._author_var = ctk.StringVar()
        self._author_entry = ctk.CTkEntry(
            meta, textvariable=self._author_var,
            placeholder_text="Autor/in …",
            fg_color=COLORS["border"], border_color=COLORS["border"],
            text_color=COLORS["text"], font=FONT_LABEL, height=34,
        )
        self._author_entry.pack(fill="x", padx=12, pady=(2, 6))
        self._author_var.trace_add("write", self._on_metadata_change)

        # Cover image
        self._cover_btn = ctk.CTkButton(
            meta, text="🖼  Cover auswählen", height=32,
            fg_color=COLORS["border"], hover_color=COLORS["surface"],
            text_color=COLORS["text_dim"], font=FONT_SMALL, corner_radius=6,
            command=self._pick_cover,
        )
        self._cover_btn.pack(fill="x", padx=12, pady=(4, 8))

        self._cover_preview = ctk.CTkLabel(meta, text="", height=0)
        self._cover_preview.pack()

        # ── Chapter section header ─────────────────────────────────────────
        ch_header = ctk.CTkFrame(sidebar, fg_color="transparent")
        ch_header.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(ch_header, text="KAPITEL", font=FONT_SMALL,
                     text_color=COLORS["text_dim"]).pack(side="left")
        ctk.CTkButton(ch_header, text="+ Neu", width=64, height=26,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
                      text_color="#fff", font=FONT_SMALL, corner_radius=6,
                      command=self._add_chapter).pack(side="right")

        # ── Scrollable chapter list ────────────────────────────────────────
        self._chapter_scroll = ctk.CTkScrollableFrame(
            sidebar, fg_color="transparent", corner_radius=0,
        )
        self._chapter_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _build_editor(self, parent) -> None:
        editor_wrap = ctk.CTkFrame(parent, fg_color=COLORS["bg"], corner_radius=0)
        editor_wrap.grid(row=0, column=2, sticky="nsew")
        editor_wrap.grid_rowconfigure(1, weight=1)
        editor_wrap.grid_columnconfigure(0, weight=1)

        # ── Chapter title bar ─────────────────────────────────────────────
        self._ch_title_frame = ctk.CTkFrame(
            editor_wrap, height=50, fg_color=COLORS["surface"], corner_radius=0)
        self._ch_title_frame.grid(row=0, column=0, sticky="ew")
        self._ch_title_frame.grid_columnconfigure(0, weight=1)

        self._ch_title_label = ctk.CTkLabel(
            self._ch_title_frame, text="Kein Kapitel ausgewählt",
            font=FONT_SUBHEAD, text_color=COLORS["text_dim"], anchor="w",
        )
        self._ch_title_label.grid(row=0, column=0, padx=20, pady=12, sticky="ew")

        self._word_count_label = ctk.CTkLabel(
            self._ch_title_frame, text="",
            font=FONT_SMALL, text_color=COLORS["text_dim"], anchor="e",
        )
        self._word_count_label.grid(row=0, column=1, padx=20)

        # ── Text editor ───────────────────────────────────────────────────
        editor_frame = ctk.CTkFrame(editor_wrap, fg_color=COLORS["bg"],
                                    corner_radius=0)
        editor_frame.grid(row=1, column=0, sticky="nsew", padx=24, pady=16)
        editor_frame.grid_rowconfigure(0, weight=1)
        editor_frame.grid_columnconfigure(0, weight=1)

        self._text_editor = tk.Text(
            editor_frame,
            wrap="word",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            insertbackground=COLORS["accent"],
            selectbackground=COLORS["accent_dim"],
            selectforeground=COLORS["text"],
            font=FONT_EDITOR,
            relief="flat",
            padx=24, pady=20,
            spacing1=4, spacing3=4,
            undo=True,
            state="disabled",
        )
        self._text_editor.grid(row=0, column=0, sticky="nsew")

        scrollbar = ctk.CTkScrollbar(editor_frame,
                                     command=self._text_editor.yview,
                                     fg_color=COLORS["bg"],
                                     button_color=COLORS["border"],
                                     button_hover_color=COLORS["accent"])
        scrollbar.grid(row=0, column=1, sticky="ns")
        self._text_editor.configure(yscrollcommand=scrollbar.set)

        # word-count update
        self._text_editor.bind("<<Modified>>", self._on_text_modified)

        # ── Empty state message ───────────────────────────────────────────
        self._empty_label = ctk.CTkLabel(
            editor_wrap,
            text="Wähle ein Kapitel aus\noder erstelle ein neues.",
            font=("Segoe UI", 15),
            text_color=COLORS["text_dim"],
        )
        self._empty_label.place(relx=0.5, rely=0.5, anchor="center")

    # ── chapter list rendering ────────────────────────────────────────────────
    def _refresh_chapter_list(self) -> None:
        for widget in self._chapter_scroll.winfo_children():
            widget.destroy()

        for i, ch in enumerate(self._book.chapters):
            row = ChapterRow(
                self._chapter_scroll, ch, i,
                on_select    = self._select_chapter,
                on_rename    = self._rename_chapter,
                on_delete    = self._delete_chapter,
                on_move_up   = self._move_chapter_up,
                on_move_down = self._move_chapter_down,
                selected     = (i == self._selected_idx),
            )
            row.pack(fill="x", pady=3)

    # ── chapter operations ────────────────────────────────────────────────────
    def _add_chapter(self) -> None:
        ch = Chapter(name=self._book.next_chapter_name(), content="")
        self._book.chapters.append(ch)
        self._selected_idx = len(self._book.chapters) - 1
        self._refresh_chapter_list()
        self._open_chapter(self._selected_idx)
        self._status.set(f'Kapitel "{ch.name}" hinzugefügt.')

    def _select_chapter(self, index: int) -> None:
        self._save_current_text()
        self._selected_idx = index
        self._refresh_chapter_list()
        self._open_chapter(index)

    def _open_chapter(self, index: int) -> None:
        if index < 0 or index >= len(self._book.chapters):
            return
        ch = self._book.chapters[index]
        self._empty_label.place_forget()
        self._text_editor.configure(state="normal")
        self._text_editor.delete("1.0", "end")
        self._text_editor.insert("1.0", ch.content)
        self._ch_title_label.configure(text=ch.name, text_color=COLORS["text"])
        self._update_word_count()

    def _save_current_text(self) -> None:
        if self._selected_idx is None:
            return
        if self._selected_idx >= len(self._book.chapters):
            return
        content = self._text_editor.get("1.0", "end-1c")
        self._book.chapters[self._selected_idx].content = content

    def _rename_chapter(self, index: int) -> None:
        ch = self._book.chapters[index]
        new_name = simpledialog.askstring(
            "Kapitel umbenennen", "Neuer Name:",
            initialvalue=ch.name, parent=self,
        )
        if new_name and new_name.strip():
            ch.name = new_name.strip()
            self._refresh_chapter_list()
            if index == self._selected_idx:
                self._ch_title_label.configure(text=ch.name)
            self._status.set(f'Kapitel umbenannt: "{ch.name}"')

    def _delete_chapter(self, index: int) -> None:
        ch = self._book.chapters[index]
        if not messagebox.askyesno(
            "Kapitel löschen",
            f'"{ch.name}" wirklich löschen?',
            parent=self,
        ):
            return
        self._book.chapters.pop(index)
        if self._selected_idx == index:
            self._selected_idx = None
            self._text_editor.configure(state="disabled")
            self._text_editor.delete("1.0", "end")
            self._ch_title_label.configure(
                text="Kein Kapitel ausgewählt", text_color=COLORS["text_dim"])
            self._word_count_label.configure(text="")
            self._empty_label.place(relx=0.5, rely=0.5, anchor="center")
        elif self._selected_idx is not None and self._selected_idx > index:
            self._selected_idx -= 1
        self._refresh_chapter_list()
        self._status.set(f'Kapitel "{ch.name}" gelöscht.')

    def _move_chapter_up(self, index: int) -> None:
        if index <= 0:
            return
        self._save_current_text()
        chs = self._book.chapters
        chs[index - 1], chs[index] = chs[index], chs[index - 1]
        if self._selected_idx == index:
            self._selected_idx -= 1
        elif self._selected_idx == index - 1:
            self._selected_idx += 1
        self._refresh_chapter_list()

    def _move_chapter_down(self, index: int) -> None:
        if index >= len(self._book.chapters) - 1:
            return
        self._save_current_text()
        chs = self._book.chapters
        chs[index + 1], chs[index] = chs[index], chs[index + 1]
        if self._selected_idx == index:
            self._selected_idx += 1
        elif self._selected_idx == index + 1:
            self._selected_idx -= 1
        self._refresh_chapter_list()

    # ── metadata change ───────────────────────────────────────────────────────
    def _on_metadata_change(self, *_) -> None:
        self._book.title  = self._title_var.get()
        self._book.author = self._author_var.get()

    # ── cover image ───────────────────────────────────────────────────────────
    def _pick_cover(self) -> None:
        path = filedialog.askopenfilename(
            title="Cover auswählen",
            filetypes=[("Bilder", "*.jpg *.jpeg *.png"), ("Alle Dateien", "*.*")],
            parent=self,
        )
        if not path:
            return
        with open(path, "rb") as f:
            raw = f.read()
        self._book.cover_b64 = base64.b64encode(raw).decode()
        # show thumbnail
        img = Image.open(io.BytesIO(raw))
        img.thumbnail((240, 140))
        self._cover_img_tk = ImageTk.PhotoImage(img)
        self._cover_preview.configure(image=self._cover_img_tk, height=140)
        self._cover_btn.configure(text="🖼  Cover ändern")
        self._status.set("Cover geladen.")

    # ── word count ────────────────────────────────────────────────────────────
    def _on_text_modified(self, _event=None) -> None:
        self._text_editor.edit_modified(False)
        self._update_word_count()

    def _update_word_count(self) -> None:
        text  = self._text_editor.get("1.0", "end-1c")
        words = len(text.split()) if text.strip() else 0
        chars = len(text)
        self._word_count_label.configure(
            text=f"{words:,} Wörter · {chars:,} Zeichen"
        )

    # ── project management ────────────────────────────────────────────────────
    def _new_project(self) -> None:
        if not messagebox.askyesno(
            "Neues Projekt",
            "Nicht gespeicherte Änderungen gehen verloren. Fortfahren?",
            parent=self,
        ):
            return
        self._book         = Book()
        self._selected_idx = None
        self._project_path = None
        self._title_var.set("")
        self._author_var.set("")
        self._cover_preview.configure(image="", height=0)
        self._cover_btn.configure(text="🖼  Cover auswählen")
        self._text_editor.configure(state="disabled")
        self._text_editor.delete("1.0", "end")
        self._ch_title_label.configure(
            text="Kein Kapitel ausgewählt", text_color=COLORS["text_dim"])
        self._word_count_label.configure(text="")
        self._empty_label.place(relx=0.5, rely=0.5, anchor="center")
        self._refresh_chapter_list()
        self.title("EPUB Editor")
        self._status.set("Neues Projekt erstellt.")

    def _open_project(self) -> None:
        path = filedialog.askopenfilename(
            title="Projekt öffnen",
            filetypes=[("EPUB Editor Projekt", "*.epubproj"), ("JSON", "*.json"),
                       ("Alle Dateien", "*.*")],
            parent=self,
        )
        if not path:
            return
        try:
            self._book         = Book.load(path)
            self._project_path = path
            self._selected_idx = None
            self._title_var.set(self._book.title)
            self._author_var.set(self._book.author)
            # restore cover preview
            if self._book.cover_b64:
                raw = base64.b64decode(self._book.cover_b64)
                img = Image.open(io.BytesIO(raw))
                img.thumbnail((240, 140))
                self._cover_img_tk = ImageTk.PhotoImage(img)
                self._cover_preview.configure(image=self._cover_img_tk, height=140)
                self._cover_btn.configure(text="🖼  Cover ändern")
            self._text_editor.configure(state="disabled")
            self._text_editor.delete("1.0", "end")
            self._empty_label.place(relx=0.5, rely=0.5, anchor="center")
            self._refresh_chapter_list()
            self.title(f"EPUB Editor – {Path(path).name}")
            self._status.set(f"Projekt geladen: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Projekt konnte nicht geladen werden:\n{e}",
                                 parent=self)

    def _save_project(self) -> None:
        self._save_current_text()
        if self._project_path:
            self._do_save(self._project_path)
        else:
            self._save_project_as()

    def _save_project_as(self) -> None:
        self._save_current_text()
        path = filedialog.asksaveasfilename(
            title="Projekt speichern unter",
            defaultextension=".epubproj",
            filetypes=[("EPUB Editor Projekt", "*.epubproj"), ("JSON", "*.json")],
            parent=self,
        )
        if path:
            self._do_save(path)

    def _do_save(self, path: str) -> None:
        try:
            self._book.save(path)
            self._project_path = path
            self.title(f"EPUB Editor – {Path(path).name}")
            self._status.set(f"Gespeichert: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Speichern fehlgeschlagen:\n{e}",
                                 parent=self)

    # ── import ────────────────────────────────────────────────────────────────
    def _import_epub(self) -> None:
        path = filedialog.askopenfilename(
            title="EPUB importieren",
            filetypes=[("EPUB Dateien", "*.epub"), ("Alle Dateien", "*.*")],
            parent=self,
        )
        if not path:
            return
        self._status.set("Importiere EPUB …")

        def do_import():
            try:
                book = import_epub(path)
                self.after(0, lambda: self._apply_imported_book(book, path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror(
                    "Import-Fehler", str(e), parent=self))
                self.after(0, lambda: self._status.set("Import fehlgeschlagen."))

        threading.Thread(target=do_import, daemon=True).start()

    def _apply_imported_book(self, book: Book, source_path: str) -> None:
        self._book         = book
        self._selected_idx = None
        self._project_path = None
        self._title_var.set(book.title)
        self._author_var.set(book.author)
        if book.cover_b64:
            raw = base64.b64decode(book.cover_b64)
            img = Image.open(io.BytesIO(raw))
            img.thumbnail((240, 140))
            self._cover_img_tk = ImageTk.PhotoImage(img)
            self._cover_preview.configure(image=self._cover_img_tk, height=140)
            self._cover_btn.configure(text="🖼  Cover ändern")
        self._text_editor.configure(state="disabled")
        self._text_editor.delete("1.0", "end")
        self._empty_label.place(relx=0.5, rely=0.5, anchor="center")
        self._refresh_chapter_list()
        self._status.set(
            f"EPUB importiert: {Path(source_path).name} "
            f"({len(book.chapters)} Kapitel)"
        )

    # ── export ────────────────────────────────────────────────────────────────
    def _pre_export_check(self) -> bool:
        self._save_current_text()
        if not self._book.chapters:
            messagebox.showwarning("Keine Kapitel",
                                   "Das Buch hat keine Kapitel.", parent=self)
            return False
        return True

    def _export_epub(self) -> None:
        if not self._pre_export_check():
            return
        path = filedialog.asksaveasfilename(
            title="Als EPUB exportieren",
            defaultextension=".epub",
            filetypes=[("EPUB", "*.epub")],
            initialfile=f"{self._book.title or 'buch'}.epub",
            parent=self,
        )
        if not path:
            return
        self._run_export(export_epub, path, "EPUB")

    def _export_pdf(self) -> None:
        if not self._pre_export_check():
            return
        path = filedialog.asksaveasfilename(
            title="Als PDF exportieren",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"{self._book.title or 'buch'}.pdf",
            parent=self,
        )
        if not path:
            return
        self._run_export(export_pdf, path, "PDF")

    def _export_docx(self) -> None:
        if not self._pre_export_check():
            return
        path = filedialog.asksaveasfilename(
            title="Als DOCX exportieren",
            defaultextension=".docx",
            filetypes=[("Word Dokument", "*.docx")],
            initialfile=f"{self._book.title or 'buch'}.docx",
            parent=self,
        )
        if not path:
            return
        self._run_export(export_docx, path, "DOCX")

    def _run_export(self, fn, path: str, fmt: str) -> None:
        self._status.set(f"Exportiere als {fmt} …")

        def do_export():
            try:
                fn(self._book, path)
                self.after(0, lambda: self._status.set(
                    f"{fmt} exportiert: {Path(path).name}"))
                self.after(0, lambda: messagebox.showinfo(
                    "Export erfolgreich",
                    f"Datei gespeichert:\n{path}", parent=self))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror(
                    "Export-Fehler", str(e), parent=self))
                self.after(0, lambda: self._status.set("Export fehlgeschlagen."))

        threading.Thread(target=do_export, daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = EPUBEditorApp()
    app.mainloop()
