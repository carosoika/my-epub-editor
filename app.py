import streamlit as st
from ebooklib import epub
import io
import html as html_module

# --- SEITENKONFIGURATION ---
st.set_page_config(
    page_title="EPUB Studio Web",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SESSION STATE INITIALISIERUNG ---
# Speicher für die Kapitel im Web-Browser-Sitzungsverlauf
if "chapters" not in st.session_state:
    st.session_state.chapters = [{"name": "Kapitel 1", "content": ""}]
if "selected_chapter_index" not in st.session_state:
    st.session_state.selected_chapter_index = 0

# --- TITEL ---
st.title("📚 EPUB Studio Web")
st.caption("Erstelle und exportiere deine E-Books direkt im Browser.")

# --- SIDEBAR: METADATEN & KAPITELLISTE ---
with st.sidebar:
    st.header("1. Buch-Metadaten")
    book_title = st.text_input("Buchtitel", value="Mein neues Buch")
    book_author = st.text_input("Autor", value="Max Mustermann")
    
    # Cover-Upload
    cover_file = st.file_uploader("Cover-Bild hochladen (Optional)", type=["jpg", "jpeg", "png"])
    if cover_file:
        st.image(cover_file, caption="Cover Vorschau", use_container_width=True)
    
    st.write("---")
    st.header("2. Kapitelverwaltung")
    
    # Kapitel hinzufügen
    if st.button("➕ Neues Kapitel hinzufügen", use_container_width=True):
        next_num = len(st.session_state.chapters) + 1
        st.session_state.chapters.append({"name": f"Kapitel {next_num}", "content": ""})
        st.session_state.selected_chapter_index = len(st.session_state.chapters) - 1
        st.rerun()

    # Kapitelauswahl via Radio-Buttons (Simuliert die Sidebar-Liste)
    chapter_names = [f"{i+1}. {ch['name']}" for i, ch in enumerate(st.session_state.chapters)]
    
    if chapter_names:
        selected_ch_format = st.radio(
            "Wähle ein Kapitel zum Bearbeiten:",
            options=chapter_names,
            index=st.session_state.selected_chapter_index
        )
        # Index zurückgewinnen
        st.session_state.selected_chapter_index = chapter_names.index(selected_ch_format)
    
    # Lösch-Button
    if len(st.session_state.chapters) > 1:
        if st.button("🗑️ Aktuelles Kapitel löschen", type="primary", use_container_width=True):
            st.session_state.chapters.pop(st.session_state.selected_chapter_index)
            st.session_state.selected_chapter_index = max(0, st.session_state.selected_chapter_index - 1)
            st.rerun()

# --- HAUPTBEREICH: EDITOR & EXPORT ---
if st.session_state.chapters:
    current_idx = st.session_state.selected_chapter_index
    current_chapter = st.session_state.chapters[current_idx]
    
    st.subheader(f"Bearbeite: {current_chapter['name']}")
    
    # Kapitel umbenennen
    new_name = st.text_input("Kapitelname ändern", value=current_chapter['name'])
    st.session_state.chapters[current_idx]['name'] = new_name
    
    # Text-Editor (Einfaches großes Textfeld wie gewünscht)
    new_content = st.text_area(
        "Kapitelinhalt (Hier Text hineinkopieren):",
        value=current_chapter['content'],
        height=400,
        placeholder="Schreibe oder füge hier deinen Text ein..."
    )
    st.session_state.chapters[current_idx]['content'] = new_content

    # --- EPUB EXPORT LOGIK (GEFIXT OHNE "DOCUMENT IS EMPTY" FEHLER) ---
    st.write("---")
    st.subheader("3. Buch Exportieren")
    
    if st.button("🚀 EPUB Datei generieren", type="secondary"):
        try:
            eb = epub.EpubBook()
            eb.set_title(book_title)
            eb.set_language('de')
            eb.add_author(book_author)
            
            # Cover hinzufügen falls vorhanden
            if cover_file:
                cover_bytes = cover_file.read()
                eb.set_cover("cover.jpg", cover_bytes)
            
            spine_items = ['nav']
            toc_items = []
            
            # Sicheres Generieren der Kapitel-XHTMLs
            for i, ch in enumerate(st.session_state.chapters):
                # Verhindert leere Dokumente, falls kein Text eingegeben wurde
                safe_content = ch['content'] if ch['content'].strip() else "Inhalt folgt..."
                
                # Einfache Absätze für Zeilenumbrüche im E-Reader generieren
                body_html = "".join(f"<p>{html_module.escape(p)}</p>" for p in safe_content.split("\n") if p.strip())
                
                epub_ch = epub.EpubHtml(
                    title=ch['name'],
                    file_name=f"chapter_{i+1:03d}.xhtml",
                    lang="de"
                )
                
                # Valide HTML-Struktur, damit ebooklib/lxml niemals abstürzt
                epub_ch.content = f"""<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="de">
<head>
    <title>{html_module.escape(ch['name'])}</title>
</head>
<body>
    <h1>{html_module.escape(ch['name'])}</h1>
    {body_html}
</body>
</html>"""
                
                eb.add_item(epub_ch)
                spine_items.append(epub_ch)
                toc_items.append(epub_ch)
            
            # Inhaltsverzeichnis und Navigation definieren
            eb.toc = toc_items
            eb.add_item(epub.EpubNav())
            eb.add_item(epub.EpubNcx())
            eb.spine = spine_items
            
            # Speicher-Buffer für den Browser-Download nutzen statt lokaler Festplatte
            buffer = io.BytesIO()
            epub.write_epub(buffer, eb, {})
            buffer.seek(0)
            
            # Download Button anzeigen
            st.success("EPUB erfolgreich generiert! Klicke unten auf Download.")
            st.download_button(
                label="📥 EPUB Herunterladen",
                data=buffer,
                file_name=f"{book_title.replace(' ', '_')}.epub",
                mime="application/epub+zip",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Fehler beim Erstellen der EPUB: {str(e)}")