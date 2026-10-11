# -*- coding: utf-8 -*-
"""
AAHM Kursant System
Android DOCX import uses the Android Storage Access Framework (SAF), which
works with modern Android scoped storage and supports selecting multiple files.
"""
import os
import re
import json
import shutil
import traceback
import datetime
import zipfile
import xml.etree.ElementTree as ET

from kivy.app import App
from kivy.metrics import dp
from kivy.utils import get_color_from_hex, platform
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.textinput import TextInput
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout

# DOCX support uses the Python standard library; no lxml/native build needed.
DOCX_AVAILABLE = True
DOCX_ERROR = ""

Window.clearcolor = get_color_from_hex("#08111F")


def write_log(message, exc=None):
    """Write a durable diagnostic log inside the app's private storage."""
    try:
        app = App.get_running_app()
        folder = app.user_data_dir if app else os.path.dirname(os.path.abspath(__file__))
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "aahm.log")
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n[{}] {}\n".format(datetime.datetime.now().isoformat(timespec="seconds"), message))
            if exc is not None:
                f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except Exception:
        pass


def dark_input(widget):
    """Apply the app's dark palette to Kivy TextInput widgets."""
    widget.background_normal = ""
    widget.background_active = ""
    widget.background_color = get_color_from_hex(PANEL2)
    widget.foreground_color = get_color_from_hex(WHITE)
    widget.cursor_color = get_color_from_hex(GOLD)
    widget.hint_text_color = get_color_from_hex(MUTED)
    widget.selection_color = (0.23, 0.42, 0.68, 0.8)
    return widget

BG = "#08111F"
PANEL = "#111F33"
PANEL2 = "#192B43"
WHITE = "#F4F7FB"
MUTED = "#9EADC1"
GOLD = "#E5B85C"
GREEN = "#2E9B72"
RED = "#D85C62"
BLUE = "#3B82F6"


FIELDS = [
    ("Adı", "adi", ["adı", "adi"]),
    ("Soyadı", "soyadi", ["soyadı", "soyadi"]),
    ("Ata adı", "ata_adi", ["ata adı", "ata adi"]),
    ("Doğum tarixi", "dogum_tarixi", ["doğulduğu tarix gün,ay,il", "doğulduğu tarix", "doğum tarixi"]),
    ("Vəsiqə seriya və FIN kodu", "fin", ["vəsiqə seriya və fin kodu", "vəsiqə seriya və fin", "fin kodu", "fin"]),
    ("Ümumvətəndaş pasportu", "pasport", ["ümumvətəndaş pasport"]),
    ("Doğulduğu yer", "doguldugu_yer", ["doğulduğu yer"]),
    ("Yaşayış yeri", "yasayis_yeri", ["yaşayış yeri"]),
    ("Qeydiyyat ünvanı", "qeydiyyat_unvani", ["qeydiyyatda olduğu ünvan"]),
    ("Qoşun növü", "qosun_novu", ["qoşun növü"]),
    ("Mülki ixtisas", "ixtisas", ["mülki ixtisas"]),
    ("İntizam balı", "intizam_bali", ["intizam balı"]),
    ("Tədris ortalama balı", "ortalama_bali", ["tədris ortalama balı"]),
    ("Fiziki hazırlığı", "fiziki_hazirliq", ["fiziki hazırlığı"]),
    ("İtirdiyi intizam balı", "itirdiyi_intizam_bali", ["kurs ərzində itirdiyi intizam balı"]),
    ("Kəsildiyi fənlər", "kesildiyi_fenler", ["son semestr imtahanı ərzində kəsildiyi fənnlər", "son semestr imtahanı ərzində kəsildiyi fənlər", "kəsildiyi fənnlər", "kəsildiyi fənlər"]),
    ("Əlaqə nömrəsi", "elaqe", ["əlaqə nömrəsi"]),
    ("Oxuduğu məktəb və ya lisey", "mekteb", ["oxuduğu məktəb və ya lisey"]),
    ("Hərbi İnstituta qəbul balı", "qebul_bali", ["hərbi instituta qəbul balı"]),
    ("Hərbi İnstituta qəbul tarixi", "qebul_tarixi", ["hərbi instituta qəbul tarixi"]),
    ("Qan qrupu", "qan_qrupu", ["qan qrupu"]),
    ("Risk qrupu", "risk_qrupu", ["risk qrupu"]),
    ("Siqaret", "siqaret", ["siqaret"]),
    ("İçki", "icki", ["içki"]),
    ("Sevimli məşğuliyyəti", "mesguliyyet", ["sevimli məşğuliyyəti"]),
    ("Ata", "ata", ["ata"]),
    ("Ata doğum tarixi", "ata_dogum_tarixi", ["ata doğum tarixi gün/ay/il", "ata doğum tarixi"]),
    ("Ata iş yeri", "ata_is_yeri", ["ata iş yeri"]),
    ("Ana", "ana", ["ana"]),
    ("Ana doğum tarixi", "ana_dogum_tarixi", ["ana doğum tarixi gün,ay,il", "ana doğum tarixi"]),
    ("Ana iş yeri", "ana_is_yeri", ["ana iş yeri"]),
    ("Bacı", "baci", ["bacı"]),
    ("Bacı doğum tarixi", "baci_dogum_tarixi", ["bacı doğum tarixi gün,ay,il", "bacı doğum tarixi"]),
    ("Bacı iş yeri", "baci_is_yeri", ["bacı iş yeri"]),
    ("Qardaş", "qardas", ["qardaş"]),
    ("Qardaş doğum tarixi", "qardas_dogum_tarixi", ["qardaş doğum tarixi gün,ay,il", "qardaş doğum tarixi"]),
    ("Qardaş iş yeri", "qardas_is_yeri", ["qardaş iş yeri"]),
    ("Ailədə vəfat edən", "vefat", ["ailədə vəfat edən"]),
    ("Boşanma halı (himayə)", "bosanma", ["boşanma halı (himayə)", "boşanma halı"]),
    ("Ailədə şəhid, qazi, əlillik", "aile_statusu", ["ailədə şəhid,qazi,əlillik", "ailədə şəhid, qazi, əlillik"]),
    ("Keçirdiyi xəstəliklər", "xestelikler", ["keçirdiyi xəstəliklər"]),
    ("Keçirdiyi travmalar", "travmalar", ["keçirdiyi travmalar"]),
    ("Kurs", "kurs", ["kurs"]),
    ("Tabor", "tabor", ["tabor"]),
    ("Bölük", "boluk", ["bölük", "boluk"]),
    ("Tağım", "taqim", ["tağım", "taqim"]),
    ("Qrup", "qrup", ["qrup"]),
    ("Yaxa №", "yaxa_no", ["yaxa №", "yaxa no"]),
    ("Hadisə haqqında məlumat", "hadise", ["hadisə haqqında məlumat"]),
]
DISPLAY_LABELS = {key: label for label, key, _ in FIELDS}


def data_path():
    folder = App.get_running_app().user_data_dir
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "kursantlar.json")


def load_data():
    try:
        with open(data_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except (OSError, ValueError, TypeError):
        return []


def save_data(data):
    path = data_path()
    temp = path + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(temp, path)


def btn(text, bg=PANEL2, fg=WHITE, height=dp(48)):
    return Button(
        text=text, size_hint_y=None, height=height,
        background_normal="", background_down="",
        background_color=get_color_from_hex(bg),
        color=get_color_from_hex(fg), bold=True, font_size="14sp"
    )


class Panel(BoxLayout):
    def __init__(self, bg=PANEL, **kwargs):
        super().__init__(**kwargs)
        with self.canvas.before:
            Color(*get_color_from_hex(bg))
            self.shape = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(16)])
        self.bind(pos=self._update, size=self._update)

    def _update(self, *_):
        self.shape.pos, self.shape.size = self.pos, self.size


def normalise(text):
    text = str(text).lower().replace(":", "").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def person_name_key(person):
    """Stable full-name key; accents/case/extra spaces do not cause false differences."""
    import unicodedata
    parts = []
    for key in ("adi", "soyadi", "ata_adi"):
        value = str(person.get(key, "") or "").casefold().strip()
        value = " ".join(value.split())
        value = "".join(ch for ch in unicodedata.normalize("NFKC", value))
        parts.append(value)
    # Require all three components so missing patronymics don't cause false duplicates.
    return tuple(parts) if all(parts) else None


def alias_map():
    result = []
    for _, key, aliases in FIELDS:
        for alias in aliases:
            result.append((normalise(alias), key))
    return sorted(result, key=lambda pair: len(pair[0]), reverse=True)


def _docx_blocks(path):
    if not zipfile.is_zipfile(path):
        raise ValueError("Fayl düzgün DOCX sənədi deyil.")
    with zipfile.ZipFile(path, "r") as z:
        try:
            root = ET.fromstring(z.read("word/document.xml"))
        except KeyError:
            raise ValueError("DOCX sənədinin məzmunu tapılmadı.")
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    body = root.find("w:body", ns)
    blocks = []
    if body is None:
        return blocks

    def ptext(p):
        return "".join((n.text or "") for n in p.findall(".//w:t", ns)).strip()

    for node in list(body):
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "p":
            text = ptext(node)
            if text:
                blocks.append(("paragraph", text))
        elif tag == "tbl":
            rows = []
            for tr in node.findall("w:tr", ns):
                cells = []
                for tc in tr.findall("w:tc", ns):
                    cells.append(" ".join(filter(None, [ptext(x) for x in tc.findall(".//w:p", ns)])).strip())
                rows.append(cells)
            blocks.append(("table", rows))
    return blocks


def parse_docx(path):
    found = {}
    aliases = alias_map()

    def match_label(text):
        label = re.sub(r"^\d+[\.)]\s*", "", normalise(text))
        for alias, key in aliases:
            if label == alias:
                return key
        return None

    for kind, content in _docx_blocks(path):
        if kind == "table":
            for cells in content:
                for i, raw in enumerate(cells):
                    if not raw:
                        continue
                    label, value = raw, ""
                    if ":" in raw:
                        label, value = raw.split(":", 1)
                        value = value.strip()
                    key = match_label(label)
                    if key:
                        if not value:
                            value = next((x.strip() for x in cells[i + 1:] if x.strip()), "")
                        if value:
                            found[key] = value
        elif ":" in content:
            label, value = content.split(":", 1)
            key = match_label(label)
            if key and value.strip():
                found[key] = value.strip()

    if not found.get("adi") and not found.get("soyadi"):
        raise ValueError("Adı/Soyadı tapılmadı. Word cədvəlində etiketləri yoxlayın.")
    found.setdefault("adi", "Naməlum")
    found.setdefault("soyadi", "Kursant")
    found.setdefault("kurs", "Hazırlıq kursu")
    found.setdefault("tabor", "Kursant hazırlıq")
    found.setdefault("boluk", "3-cü kursant")
    return found


def _xml_escape(value):
    return (str(value or "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def write_docx(path, title, rows):
    """Write a basic Office Open XML document without python-docx/lxml."""
    def para(text, title_style=False):
        ppr = '<w:pPr><w:pStyle w:val="Title"/></w:pPr>' if title_style else ""
        return '<w:p>' + ppr + '<w:r><w:t xml:space="preserve">' + _xml_escape(text) + '</w:t></w:r></w:p>'

    tr = ['<w:tr><w:tc><w:p><w:r><w:t>Parametr</w:t></w:r></w:p></w:tc>'
          '<w:tc><w:p><w:r><w:t>Məlumat</w:t></w:r></w:p></w:tc></w:tr>']
    for label, value in rows:
        tr.append('<w:tr><w:tc>' + para(label) + '</w:tc><w:tc>' + para(value) + '</w:tc></w:tr>')

    docxml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
      + para(title, True) + '<w:tbl><w:tblPr><w:tblBorders>'
      + ''.join(f'<w:{edge} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
                for edge in ("top","left","bottom","right","insideH","insideV"))
      + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
      + ''.join(tr) + '</w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
      '</w:body></w:document>')
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      '</Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
      '<w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
      '<w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
      '</w:styles>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", docxml)
        z.writestr("word/styles.xml", styles)


# -------- Android native document picker (SAF), with multi-select --------
class AndroidDocPicker:
    REQUEST_CODE = 4817
    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    @staticmethod
    def open():
        from jnius import autoclass
        from android import activity

        Intent = autoclass("android.content.Intent")
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        intent.setType(AndroidDocPicker.MIME)
        intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, True)
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        activity.bind(on_activity_result=AndroidDocPicker._on_result)
        PythonActivity.mActivity.startActivityForResult(intent, AndroidDocPicker.REQUEST_CODE)

    @staticmethod
    def _on_result(request_code, result_code, intent):
        if request_code != AndroidDocPicker.REQUEST_CODE:
            return
        try:
            from android import activity
            activity.unbind(on_activity_result=AndroidDocPicker._on_result)
        except Exception:
            pass

        app = App.get_running_app()
        screen = app.root.get_screen("import")
        if intent is None:
            screen.set_status("Fayl seçimi ləğv edildi.", RED)
            return

        try:
            Activity = __import__("jnius").autoclass("android.app.Activity")
            if result_code != Activity.RESULT_OK:
                screen.set_status("Fayl seçilmədi.", RED)
                return

            uris = []
            clip = intent.getClipData()
            if clip is not None:
                for i in range(clip.getItemCount()):
                    uris.append(clip.getItemAt(i).getUri())
            else:
                uri = intent.getData()
                if uri is not None:
                    uris.append(uri)

            if not uris:
                screen.set_status("Seçilmiş fayl tapılmadı.", RED)
                return

            screen.copy_selected_uris(uris)
        except Exception as exc:
            write_log("Word faylı seçmə xətası", exc)
            screen.set_status("Fayl seçmə xətası: " + str(exc)[:160], RED)



class AndroidExportPicker:
    """Let the user choose a visible Documents/Downloads parent folder, then create
    an 'Arayış-məruzə' subfolder and save the generated DOCX there via SAF."""
    REQUEST_CODE = 4818
    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    pending_screen = None
    pending_temp_path = None
    pending_filename = None

    @staticmethod
    def open(screen, temp_path, filename):
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        AndroidExportPicker.pending_screen = screen
        AndroidExportPicker.pending_temp_path = temp_path
        AndroidExportPicker.pending_filename = filename
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT_TREE)
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION |
                        Intent.FLAG_GRANT_WRITE_URI_PERMISSION |
                        Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION)
        activity.bind(on_activity_result=AndroidExportPicker._on_result)
        PythonActivity.mActivity.startActivityForResult(intent, AndroidExportPicker.REQUEST_CODE)

    @staticmethod
    def _on_result(request_code, result_code, intent):
        if request_code != AndroidExportPicker.REQUEST_CODE:
            return
        try:
            from android import activity
            activity.unbind(on_activity_result=AndroidExportPicker._on_result)
        except Exception:
            pass
        screen = AndroidExportPicker.pending_screen
        temp_path = AndroidExportPicker.pending_temp_path
        filename = AndroidExportPicker.pending_filename
        try:
            Activity = __import__("jnius").autoclass("android.app.Activity")
            if result_code != Activity.RESULT_OK or intent is None:
                if screen:
                    screen.status.color = get_color_from_hex(GOLD)
                    screen.status.text = "Qovluq seçimi ləğv edildi; Word faylı saxlanmadı."
                return
            from jnius import autoclass
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            resolver = PythonActivity.mActivity.getContentResolver()
            DocumentsContract = autoclass("android.provider.DocumentsContract")
            tree_uri = intent.getData()
            # Retain permission for future Android access when supported.
            try:
                PythonActivity.mActivity.getContentResolver().takePersistableUriPermission(
                    tree_uri, IntentFlags.read_write())
            except Exception:
                pass
            tree_id = DocumentsContract.getTreeDocumentId(tree_uri)
            parent_uri = DocumentsContract.buildDocumentUriUsingTree(tree_uri, tree_id)
            folder_uri = None
            children_uri = DocumentsContract.buildChildDocumentsUriUsingTree(tree_uri, tree_id)
            cursor = resolver.query(children_uri, None, None, None, None)
            if cursor is not None:
                try:
                    name_col = cursor.getColumnIndex("_display_name")
                    id_col = cursor.getColumnIndex("document_id")
                    while cursor.moveToNext():
                        if cursor.getString(name_col) == "Arayış-məruzə":
                            folder_uri = DocumentsContract.buildDocumentUriUsingTree(
                                tree_uri, cursor.getString(id_col))
                            break
                finally:
                    cursor.close()
            if folder_uri is None:
                folder_uri = DocumentsContract.createDocument(
                    resolver, parent_uri, "vnd.android.document/directory", "Arayış-məruzə")
            if folder_uri is None:
                raise IOError("Arayış-məruzə qovluğu yaradıla bilmədi.")
            # Replace an existing file with the same name so repeated exports update it.
            child_uri = DocumentsContract.buildChildDocumentsUriUsingTree(
                tree_uri, DocumentsContract.getDocumentId(folder_uri))
            cursor = resolver.query(child_uri, None, None, None, None)
            if cursor is not None:
                try:
                    name_col = cursor.getColumnIndex("_display_name")
                    id_col = cursor.getColumnIndex("document_id")
                    while cursor.moveToNext():
                        if cursor.getString(name_col) == filename:
                            old_uri = DocumentsContract.buildDocumentUriUsingTree(
                                tree_uri, cursor.getString(id_col))
                            DocumentsContract.deleteDocument(resolver, old_uri)
                            break
                finally:
                    cursor.close()
            file_uri = DocumentsContract.createDocument(resolver, folder_uri,
                                                        AndroidExportPicker.MIME, filename)
            if file_uri is None:
                raise IOError("Word faylı yaradıla bilmədi.")
            source = open(temp_path, "rb")
            output = resolver.openOutputStream(file_uri, "w")
            if output is None:
                source.close()
                raise IOError("Faylı seçilmiş qovluğa yazmaq olmadı.")
            try:
                while True:
                    chunk = source.read(8192)
                    if not chunk:
                        break
                    # Stream bytes individually for compatibility with PyJNIus Java arrays.
                    for byte_value in chunk:
                        output.write(int(byte_value) & 0xFF)
                output.flush()
            finally:
                source.close()
                output.close()
            if screen:
                screen.status.color = get_color_from_hex(GREEN)
                screen.status.text = "Saxlanıldı: Arayış-məruzə/" + filename
        except Exception as exc:
            write_log("Word faylını telefona saxlamaq alınmadı", exc)
            if screen:
                screen.status.color = get_color_from_hex(RED)
                screen.status.text = "Saxlama xətası: " + str(exc)[:130]
        finally:
            try:
                if temp_path and os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            AndroidExportPicker.pending_screen = None
            AndroidExportPicker.pending_temp_path = None
            AndroidExportPicker.pending_filename = None


class IntentFlags:
    @staticmethod
    def read_write():
        from jnius import autoclass
        Intent = autoclass("android.content.Intent")
        return Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION

def copy_android_uri(uri, destination):
    """Copy SAF content:// URI into a normal app-private .docx file."""
    # Do not import jarray: some packaged PyJNIus builds do not expose it.
    # InputStream.read() without arguments returns one byte as an integer.
    from jnius import autoclass
    PythonActivity = autoclass("org.kivy.android.PythonActivity")
    resolver = PythonActivity.mActivity.getContentResolver()
    stream = resolver.openInputStream(uri)
    if stream is None:
        raise IOError("Android faylı açmadı.")
    try:
        with open(destination, "wb") as output:
            chunk = bytearray()
            while True:
                value = stream.read()
                if value == -1:
                    break
                chunk.append(int(value) & 0xFF)
                # Write in small batches to avoid holding the whole document in memory.
                if len(chunk) >= 8192:
                    output.write(chunk)
                    chunk.clear()
            if chunk:
                output.write(chunk)
    finally:
        stream.close()


# -------------------- LOGIN --------------------
class Login(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        root = BoxLayout(orientation="vertical", padding=dp(24), spacing=dp(14))
        root.add_widget(Label(text="", size_hint_y=.12))

        hero = Panel(orientation="vertical", padding=dp(12), spacing=dp(5),
                     size_hint_y=None, height=dp(190), bg="#13243A")
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aahm_logo.png")
        if os.path.isfile(logo_path):
            logo = Image(source=logo_path, size_hint=(1, None), height=dp(115),
                         allow_stretch=True, keep_ratio=True)
            hero.add_widget(logo)
        else:
            hero.add_widget(Label(text="AAHM", font_size="36sp", bold=True,
                                  color=get_color_from_hex(GOLD)))
        hero.add_widget(Label(text="KURSANT İDARƏETMƏ SİSTEMİ",
                              font_size="12sp", color=get_color_from_hex(MUTED)))
        root.add_widget(hero)
        root.add_widget(Label(text="İdarəetmə panelinə xoş gəlmisiniz",
                              color=get_color_from_hex(WHITE), size_hint_y=None, height=dp(42)))

        self.username = dark_input(TextInput(hint_text="İstifadəçi adı", multiline=False,
                                  size_hint_y=None, height=dp(50)))
        self.password = dark_input(TextInput(hint_text="Şifrə", password=True, multiline=False,
                                  size_hint_y=None, height=dp(50)))
        root.add_widget(self.username)
        root.add_widget(self.password)

        self.message = Label(text="", color=get_color_from_hex(RED),
                             size_hint_y=None, height=dp(32))
        root.add_widget(self.message)
        login_btn = btn("DAXİL OL", GOLD, "#101827", dp(54))
        login_btn.bind(on_release=self.login)
        root.add_widget(login_btn)
        root.add_widget(Label(text="Sınaq giriş: admin / admin",
                              color=get_color_from_hex(MUTED), font_size="12sp",
                              size_hint_y=None, height=dp(28)))
        self.add_widget(root)

    def login(self, *_):
        if self.username.text.strip() == "admin" and self.password.text == "admin":
            self.message.text = ""
            self.manager.current = "main"
        else:
            self.message.text = "İstifadəçi adı və ya şifrə yanlışdır."


# -------------------- BACKUP / RESTORE VIA ANDROID FILE PICKER --------------------
class AndroidBackup:
    REQUEST_EXPORT = 4821
    REQUEST_IMPORT = 4822

    @staticmethod
    def export_backup():
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        intent = Intent(Intent.ACTION_CREATE_DOCUMENT)
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        intent.setType("application/json")
        intent.putExtra(Intent.EXTRA_TITLE, "AAHM_kursant_backup.json")
        intent.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
        activity.bind(on_activity_result=AndroidBackup._on_export_result)
        PythonActivity.mActivity.startActivityForResult(intent, AndroidBackup.REQUEST_EXPORT)

    @staticmethod
    def import_backup():
        from jnius import autoclass
        from android import activity
        Intent = autoclass("android.content.Intent")
        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        # Some Android file providers hide .json files when MIME is application/json.
        # Use a wildcard picker and validate the selected file contents below.
        intent.setType("*/*")
        intent.putExtra(Intent.EXTRA_MIME_TYPES, ["application/json", "text/json", "text/plain", "application/octet-stream"])
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        activity.bind(on_activity_result=AndroidBackup._on_import_result)
        PythonActivity.mActivity.startActivityForResult(intent, AndroidBackup.REQUEST_IMPORT)

    @staticmethod
    def _unbind(callback):
        try:
            from android import activity
            activity.unbind(on_activity_result=callback)
        except Exception:
            pass

    @staticmethod
    def _on_export_result(request_code, result_code, intent):
        if request_code != AndroidBackup.REQUEST_EXPORT:
            return
        AndroidBackup._unbind(AndroidBackup._on_export_result)
        screen = App.get_running_app().root.get_screen("main")
        try:
            Activity = __import__("jnius").autoclass("android.app.Activity")
            if intent is None or result_code != Activity.RESULT_OK:
                screen.set_backup_status("Backup yaradılması ləğv edildi.", GOLD)
                return
            uri = intent.getData()
            if uri is None:
                raise ValueError("Fayl yeri seçilmədi.")
            from jnius import autoclass
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            stream = PythonActivity.mActivity.getContentResolver().openOutputStream(uri)
            if stream is None:
                raise IOError("Backup faylı yazmaq üçün açılmadı.")
            payload = json.dumps(load_data(), ensure_ascii=False, indent=2).encode("utf-8")
            try:
                for byte in payload:
                    stream.write(int(byte) & 0xFF)
                stream.flush()
            finally:
                stream.close()
            screen.set_backup_status(f"Backup yaradıldı: {len(load_data())} kursant.", GREEN)
        except Exception as exc:
            screen.set_backup_status("Backup xətası: " + str(exc)[:140], RED)

    @staticmethod
    def _on_import_result(request_code, result_code, intent):
        if request_code != AndroidBackup.REQUEST_IMPORT:
            return
        AndroidBackup._unbind(AndroidBackup._on_import_result)
        screen = App.get_running_app().root.get_screen("main")
        try:
            Activity = __import__("jnius").autoclass("android.app.Activity")
            if intent is None or result_code != Activity.RESULT_OK:
                screen.set_backup_status("Backup seçimi ləğv edildi.", GOLD)
                return
            uri = intent.getData()
            if uri is None:
                raise ValueError("Fayl seçilmədi.")
            from jnius import autoclass
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            stream = PythonActivity.mActivity.getContentResolver().openInputStream(uri)
            if stream is None:
                raise IOError("Backup faylı oxunmadı.")
            raw = bytearray()
            try:
                while True:
                    value = stream.read()
                    if value == -1:
                        break
                    raw.append(int(value) & 0xFF)
                    if len(raw) > 25 * 1024 * 1024:
                        raise ValueError("Backup faylı çox böyükdür.")
            finally:
                stream.close()
            restored = json.loads(raw.decode("utf-8-sig"))
            # Accept the original list format and common wrapped backup formats.
            if isinstance(restored, dict):
                restored = restored.get("kursantlar", restored.get("data", restored.get("students")))
            if not isinstance(restored, list) or any(not isinstance(x, dict) for x in restored):
                raise ValueError("Backup JSON formatı uyğun deyil. AAHM backup JSON faylını seçin.")
            current = load_data()
            known = {person_name_key(x) for x in current if person_name_key(x)}
            added = 0
            for person in restored:
                key = person_name_key(person)
                if key and key in known:
                    continue
                current.append(person)
                added += 1
                if key:
                    known.add(key)
            save_data(current)
            screen.refresh()
            screen.set_backup_status(f"Bərpa edildi: {added} yeni kursant; təkrarlar buraxıldı.", GREEN)
        except Exception as exc:
            screen.set_backup_status("Bərpa xətası: " + str(exc)[:140], RED)


# -------------------- MAIN --------------------
class Main(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        root = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(9))

        top = BoxLayout(size_hint_y=None, height=dp(50))
        top.add_widget(Label(text="AAHM  |  Kursantlar", bold=True, font_size="20sp",
                             color=get_color_from_hex(WHITE), halign="left"))
        logout = btn("Çıxış", "#74343C", height=dp(40))
        logout.size_hint_x = None
        logout.width = dp(78)
        logout.bind(on_release=lambda *_: setattr(self.manager, "current", "login"))
        top.add_widget(logout)
        root.add_widget(top)

        self.search = dark_input(TextInput(hint_text="Ad, soyad və ata adı ilə axtar...",
                                multiline=False, size_hint_y=None, height=dp(46)))
        self.search.bind(text=lambda *_: self.refresh())
        root.add_widget(self.search)

        self.count = Label(text="KURSANTLAR: 0", color=get_color_from_hex(MUTED),
                           size_hint_y=None, height=dp(25))
        root.add_widget(self.count)

        scroll = ScrollView()
        self.listing = GridLayout(cols=1, spacing=dp(8), size_hint_y=None, padding=[0, dp(3)])
        self.listing.bind(minimum_height=self.listing.setter("height"))
        scroll.add_widget(self.listing)
        root.add_widget(scroll)

        add = btn("WORD FAYLLARINI İDXAL ET", GREEN, height=dp(54))
        add.bind(on_release=lambda *_: setattr(self.manager, "current", "import"))
        root.add_widget(add)

        backup_row = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
        backup_out = btn("BACKUP YARAT", BLUE, height=dp(46))
        backup_out.bind(on_release=self.create_backup)
        backup_in = btn("BACKUP-DAN BƏRPA ET", PANEL2, height=dp(46))
        backup_in.bind(on_release=self.restore_backup)
        backup_row.add_widget(backup_out)
        backup_row.add_widget(backup_in)
        root.add_widget(backup_row)
        self.backup_status = Label(text="Backup faylını başqa telefona göndərərək kursantları köçürə bilərsiniz.",
                                   color=get_color_from_hex(MUTED), font_size="11sp",
                                   size_hint_y=None, height=dp(34))
        root.add_widget(self.backup_status)
        log_button = btn("LOG FAYLINI SAXLA", PANEL2, height=dp(42))
        log_button.bind(on_release=self.export_log)
        root.add_widget(log_button)
        self.add_widget(root)

    def export_log(self, *_):
        try:
            path = os.path.join(App.get_running_app().user_data_dir, "aahm.log")
            if not os.path.isfile(path):
                write_log("Log faylı yaradıldı; tətbiq işləyir.")
            if platform == "android":
                from jnius import autoclass
                from android import activity
                Intent = autoclass("android.content.Intent")
                PythonActivity = autoclass("org.kivy.android.PythonActivity")
                intent = Intent(Intent.ACTION_CREATE_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                intent.setType("text/plain")
                intent.putExtra(Intent.EXTRA_TITLE, "AAHM-log.txt")
                activity.bind(on_activity_result=Main._on_log_result)
                Main.pending_log_path = path
                PythonActivity.mActivity.startActivityForResult(intent, 4899)
            else:
                self.set_backup_status("Log faylı: " + path, GREEN)
        except Exception as exc:
            write_log("Log faylını saxlamaq alınmadı", exc)
            self.set_backup_status("Log xətası: " + str(exc)[:100], RED)

    pending_log_path = None

    @staticmethod
    def _on_log_result(request_code, result_code, intent):
        if request_code != 4899:
            return
        try:
            from android import activity
            activity.unbind(on_activity_result=Main._on_log_result)
        except Exception:
            pass
        try:
            if intent is None or Main.pending_log_path is None:
                return
            from jnius import autoclass
            Activity = autoclass("android.app.Activity")
            if result_code != Activity.RESULT_OK:
                return
            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            resolver = PythonActivity.mActivity.getContentResolver()
            uri = intent.getData()
            with open(Main.pending_log_path, "rb") as source:
                output = resolver.openOutputStream(uri, "w")
                if output is None:
                    raise IOError("Log faylı yazıla bilmədi")
                try:
                    while True:
                        chunk = source.read(4096)
                        if not chunk:
                            break
                        for value in chunk:
                            output.write(int(value) & 255)
                    output.flush()
                finally:
                    output.close()
        except Exception as exc:
            write_log("Log ixrac xətası", exc)
        finally:
            Main.pending_log_path = None

    def on_enter(self, *_):
        self.refresh()

    def refresh(self):
        if not hasattr(self, "listing"):
            return
        self.listing.clear_widgets()
        query = self.search.text.casefold().strip()
        count = 0
        for person in load_data():
            name = " ".join(str(person.get(k, "") or "") for k in ("soyadi", "adi", "ata_adi")).strip()
            if query and query not in name.casefold():
                continue
            count += 1
            item = btn(" " + (name or "Adsız kursant") + "   ›", PANEL2, height=dp(58))
            item.halign = "left"
            item.bind(on_release=lambda _b, p=person: self.open_person(p))
            self.listing.add_widget(item)
        self.count.text = f"KURSANTLAR: {count}"
        if count == 0:
            self.listing.add_widget(Label(text="Siyahı boşdur. Word fayllarını idxal edin.",
                                          color=get_color_from_hex(MUTED),
                                          size_hint_y=None, height=dp(70)))

    def set_backup_status(self, message, color=MUTED):
        self.backup_status.text = message
        self.backup_status.color = get_color_from_hex(color)

    def create_backup(self, *_):
        if platform != "android":
            self.set_backup_status("Backup fayl seçicisi Android-də işləyir.", GOLD)
            return
        try:
            AndroidBackup.export_backup()
        except Exception as exc:
            self.set_backup_status("Backup açıla bilmədi: " + str(exc)[:120], RED)

    def restore_backup(self, *_):
        if platform != "android":
            self.set_backup_status("Backup fayl seçicisi Android-də işləyir.", GOLD)
            return
        try:
            AndroidBackup.import_backup()
        except Exception as exc:
            self.set_backup_status("Bərpa açıla bilmədi: " + str(exc)[:120], RED)

    def open_person(self, person):
        self.manager.get_screen("detail").show_person(person)
        self.manager.current = "detail"


# -------------------- IMPORT MULTIPLE DOCX --------------------
class Import(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.selected_paths = []

        root = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(10))
        bar = BoxLayout(size_hint_y=None, height=dp(46))
        back = btn("‹ Geri", height=dp(42))
        back.size_hint_x = None
        back.width = dp(82)
        back.bind(on_release=lambda *_: setattr(self.manager, "current", "main"))
        bar.add_widget(back)
        bar.add_widget(Label(text="Word fayllarını idxal et", color=get_color_from_hex(WHITE),
                             bold=True, font_size="17sp"))
        root.add_widget(bar)

        hero = Panel(orientation="vertical", padding=dp(14), spacing=dp(6),
                     size_hint_y=None, height=dp(115), bg="#13243A")
        hero.add_widget(Label(text="DOCX  •  ÇOXSEÇİM", color=get_color_from_hex(GOLD),
                              bold=True, font_size="18sp"))
        hero.add_widget(Label(text="Bir dəfəyə bir neçə Word sənədi seçə bilərsiniz.",
                              color=get_color_from_hex(MUTED), font_size="12sp"))
        root.add_widget(hero)

        choose = btn("TELEFONDAN WORD FAYLLARI SEÇ", BLUE, height=dp(54))
        choose.bind(on_release=self.choose_files)
        root.add_widget(choose)

        self.file_count = Label(text="Hələ fayl seçilməyib",
                                color=get_color_from_hex(MUTED),
                                size_hint_y=None, height=dp(30))
        root.add_widget(self.file_count)

        scroll = ScrollView()
        self.file_list = GridLayout(cols=1, spacing=dp(6), size_hint_y=None, padding=[0, dp(4)])
        self.file_list.bind(minimum_height=self.file_list.setter("height"))
        scroll.add_widget(self.file_list)
        root.add_widget(scroll)

        bottom = Panel(orientation="vertical", padding=dp(10), spacing=dp(8),
                       size_hint_y=None, height=dp(118), bg=PANEL)
        self.status = Label(text="DOCX fayllarını seçmək üçün yuxarıdakı düyməyə basın.",
                            color=get_color_from_hex(MUTED), font_size="12sp")
        bottom.add_widget(self.status)
        actions = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
        clear = btn("Təmizlə", "#74343C", height=dp(46))
        clear.size_hint_x = .36
        clear.bind(on_release=self.clear_files)
        actions.add_widget(clear)
        self.import_button = btn("HAMISINI ƏLAVƏ ET", GREEN, height=dp(46))
        self.import_button.bind(on_release=self.import_all)
        actions.add_widget(self.import_button)
        bottom.add_widget(actions)
        root.add_widget(bottom)
        self.add_widget(root)

    def on_enter(self, *_):
        self.render_files()

    def choose_files(self, *_):
        if not DOCX_AVAILABLE:
            self.set_status("python-docx yüklənməyib: " + DOCX_ERROR[:100], RED)
            # Still allow selecting files; import will show a useful error.
        if platform == "android":
            try:
                AndroidDocPicker.open()
            except Exception as exc:
                self.set_status("Android fayl seçicisi açıla bilmədi: " + str(exc)[:150], RED)
        else:
            self.set_status("Çoxseçimli Android fayl seçicisi yalnız Android-də işləyir.", GOLD)

    def copy_selected_uris(self, uris):
        folder = os.path.join(App.get_running_app().user_data_dir, "incoming_docx")
        os.makedirs(folder, exist_ok=True)
        added = 0
        errors = []
        for index, uri in enumerate(uris):
            try:
                name = str(uri.getLastPathSegment() or f"document_{index + 1}.docx")
                name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
                if not name.lower().endswith(".docx"):
                    name += ".docx"
                destination = os.path.join(folder, f"{index}_{name}")
                copy_android_uri(uri, destination)
                # Check ZIP/DOCX signature before adding it.
                with open(destination, "rb") as f:
                    signature = f.read(4)
                if signature[:2] != b"PK":
                    os.remove(destination)
                    raise ValueError("Fayl DOCX formatında görünmür.")
                if destination not in self.selected_paths:
                    self.selected_paths.append(destination)
                    added += 1
            except Exception as exc:
                errors.append(str(exc)[:90])
        self.render_files()
        if added:
            self.set_status(f"{added} fayl seçildi. İdxal üçün aşağıdakı düyməyə basın.", GREEN)
        elif errors:
            self.set_status("Fayllar oxunmadı: " + "; ".join(errors[:2]), RED)
        else:
            self.set_status("Yeni fayl seçilmədi.", GOLD)

    def render_files(self):
        self.file_list.clear_widgets()
        self.file_count.text = f"Seçilmiş fayllar: {len(self.selected_paths)}"
        for path in self.selected_paths:
            row = Panel(orientation="horizontal", padding=[dp(10), dp(6)],
                        spacing=dp(8), size_hint_y=None, height=dp(54), bg=PANEL2)
            name = os.path.basename(path)
            label = Label(text="DOCX: " + name, color=get_color_from_hex(WHITE),
                          font_size="12sp", halign="left", valign="middle")
            label.bind(size=lambda widget, size: setattr(widget, "text_size", size))
            row.add_widget(label)
            self.file_list.add_widget(row)

    def set_status(self, text, color=MUTED):
        self.status.text = text
        self.status.color = get_color_from_hex(color)

    def clear_files(self, *_):
        self.selected_paths = []
        self.render_files()
        self.set_status("Seçilmiş fayllar siyahısı təmizləndi.", MUTED)

    def import_all(self, *_):
        if not self.selected_paths:
            self.set_status("Əvvəlcə telefondan bir və ya bir neçə DOCX faylı seçin.", RED)
            return
        if not DOCX_AVAILABLE:
            self.set_status("Word kitabxanası yoxdur: " + DOCX_ERROR[:130], RED)
            return

        data = load_data()
        imported = 0
        skipped = 0
        failures = []

        for path in list(self.selected_paths):
            try:
                person = parse_docx(path)
                full_name_key = person_name_key(person)
                duplicate = bool(full_name_key) and any(
                    person_name_key(old) == full_name_key for old in data
                )
                if duplicate:
                    skipped += 1
                    continue
                data.append(person)
                imported += 1
            except Exception as exc:
                failures.append(os.path.basename(path) + ": " + str(exc)[:70])

        try:
            save_data(data)
        except Exception as exc:
            self.set_status("Məlumatları saxlamaq mümkün olmadı: " + str(exc)[:120], RED)
            return

        self.manager.get_screen("main").refresh()
        summary = f"Əlavə edildi: {imported}; təkrar: {skipped}; xəta: {len(failures)}."
        if failures:
            summary += " " + failures[0][:120]
        self.set_status(summary, GREEN if not failures else GOLD)
        if imported or skipped:
            self.selected_paths = []
            self.render_files()


# -------------------- DETAILS + WORD EXPORT --------------------
class Detail(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.person = {}
        root = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))

        bar = BoxLayout(size_hint_y=None, height=dp(44))
        back = btn("‹ Geri", height=dp(40))
        back.size_hint_x = None
        back.width = dp(80)
        back.bind(on_release=lambda *_: setattr(self.manager, "current", "main"))
        bar.add_widget(back)
        bar.add_widget(Label(text="Kursant məlumatları", color=get_color_from_hex(WHITE),
                             bold=True, font_size="18sp"))
        root.add_widget(bar)

        scroll = ScrollView()
        self.rows = GridLayout(cols=1, spacing=dp(5), size_hint_y=None, padding=[0, dp(3)])
        self.rows.bind(minimum_height=self.rows.setter("height"))
        scroll.add_widget(self.rows)
        root.add_widget(scroll)

        export_button = btn("WORD ARAYIŞ-MƏRUZƏSİ YARAT", GOLD, "#101827", dp(50))
        export_button.bind(on_release=self.export)
        root.add_widget(export_button)
        delete_button = btn("KURSANTI SİYAHIDAN SİL", "#74343C", height=dp(48))
        delete_button.bind(on_release=self.delete_person)
        root.add_widget(delete_button)
        self.status = Label(text="", color=get_color_from_hex(MUTED),
                            size_hint_y=None, height=dp(46), font_size="11sp")
        root.add_widget(self.status)
        self.add_widget(root)

    def show_person(self, person):
        self.person = dict(person)
        self.rows.clear_widgets()

        name = " ".join(str(self.person.get(k, "") or "").strip()
                        for k in ("soyadi", "adi", "ata_adi")
                        if str(self.person.get(k, "") or "").strip())
        header = Panel(orientation="vertical", padding=dp(12), spacing=dp(4),
                       size_hint_y=None, height=dp(110), bg="#13243A")
        header.add_widget(Label(text=name or "Adsız kursant",
                                color=get_color_from_hex(GOLD), font_size="19sp", bold=True))
        for label, key in [("Kurs", "kurs"), ("Tabor", "tabor"),
                           ("Bölük", "boluk"), ("Tağım", "taqim"), ("Qrup", "qrup")]:
            header.add_widget(Label(text=f"{label}: {self.person.get(key, '—') or '—'}",
                                    color=get_color_from_hex(MUTED), font_size="12sp"))
        self.rows.add_widget(header)

        hidden = {"adi", "soyadi", "ata_adi", "kurs", "tabor", "boluk", "taqim", "qrup"}
        for key, value in self.person.items():
            if key in hidden:
                continue
            row = Panel(orientation="horizontal", padding=[dp(10), dp(7)],
                        spacing=dp(6), size_hint_y=None, height=dp(60), bg=PANEL)
            left = Label(text=DISPLAY_LABELS.get(key, key), color=get_color_from_hex(MUTED),
                         font_size="12sp", size_hint_x=.46, halign="left", valign="middle")
            right = Label(text=str(value or "—"), color=get_color_from_hex(WHITE),
                          font_size="13sp", size_hint_x=.54, halign="left", valign="middle")
            left.bind(size=lambda w, s: setattr(w, "text_size", s))
            right.bind(size=lambda w, s: setattr(w, "text_size", s))
            row.add_widget(left)
            row.add_widget(right)
            self.rows.add_widget(row)
        self.status.text = ""

    def delete_person(self, *_):
        current_key = person_name_key(self.person)
        data = load_data()
        # Delete only the selected record matching all three name fields.
        new_data = []
        removed = False
        for person in data:
            if not removed and current_key and person_name_key(person) == current_key:
                removed = True
                continue
            new_data.append(person)
        if not removed:
            self.status.color = get_color_from_hex(RED)
            self.status.text = "Kursant siyahıda tapılmadı; silinmədi."
            return
        try:
            save_data(new_data)
        except Exception as exc:
            self.status.color = get_color_from_hex(RED)
            self.status.text = "Silinmə xətası: " + str(exc)[:100]
            return
        self.manager.get_screen("main").refresh()
        self.manager.current = "main"

    def export(self, *_):
        temp_path = None
        try:
            parts = [str(self.person.get(k, "") or "").strip()
                     for k in ("adi", "soyadi", "ata_adi")]
            safe_parts = [re.sub(r"[\\/:*?\"<>|]+", "_", part).strip(" .")
                          for part in parts if part]
            filename = (" ".join(safe_parts) or "KURSANT") .upper() + ".DOCX"
            folder = os.path.join(App.get_running_app().user_data_dir, "exports_tmp")
            os.makedirs(folder, exist_ok=True)
            temp_path = os.path.join(folder, filename)
            rows = [(DISPLAY_LABELS.get(key, key), str(value or ""))
                    for key, value in self.person.items()]
            write_docx(temp_path, "ARAYIŞ-MƏRUZƏ", rows)
            # Fail early with a clear message if the generated DOCX package is invalid.
            with zipfile.ZipFile(temp_path, "r") as docx_check:
                required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
                if not required.issubset(set(docx_check.namelist())):
                    raise ValueError("Word sənədinin daxili quruluşu natamamdır.")
                ET.fromstring(docx_check.read("word/document.xml"))
            if platform == "android":
                self.status.color = get_color_from_hex(MUTED)
                self.status.text = "Saxlanacaq yeri seçin (məsələn, Documents)."
                AndroidExportPicker.open(self, temp_path, filename)
            else:
                folder = os.path.join(App.get_running_app().user_data_dir, "Arayış-məruzə")
                os.makedirs(folder, exist_ok=True)
                path = os.path.join(folder, filename)
                shutil.copyfile(temp_path, path)
                os.remove(temp_path)
                self.status.color = get_color_from_hex(GREEN)
                self.status.text = "Saxlanıldı: " + path
        except Exception as exc:
            try:
                if temp_path and os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
            self.status.color = get_color_from_hex(RED)
            self.status.text = "Word yaratma xətası: " + str(exc)[:150]



class AAHMApp(App):
    def build(self):
        self.title = "AAHM Kursant Sistemi"
        manager = ScreenManager()
        manager.add_widget(Login(name="login"))
        manager.add_widget(Main(name="main"))
        manager.add_widget(Import(name="import"))
        manager.add_widget(Detail(name="detail"))
        return manager


if __name__ == "__main__":
    try:
        AAHMApp().run()
    except Exception as exc:
        write_log("Tətbiqdə tutulmamış xəta", exc)
        traceback.print_exc()
