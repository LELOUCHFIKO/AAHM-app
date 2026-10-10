import json
import os
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.textinput import TextInput
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.graphics import Color, Rectangle
from docx import Document

DATA_FILE = 'kursantlar.json'

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return [
        {
            "adi": "Elvin", "soyadi": "Məmmədov", "ata_adi": "Rəhim",
            "dogum_tarixi": "12.05.2005", "fin": "1234567",
            "kurs": "Hazırlıq kursu", "tabor": "Kursant hazırlıq",
            "boluk": "3-cü kursant", "elaqe": "+994501234567",
            "qan_qrupu": "II (A) Rh+", "ata": "Rəhim Məmmədov", "ana": "Nərmin Məmmədova"
        }
    ]

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# Arxa fon rəngi üçün köməkçi sinif
class BackgroundBoxLayout(BoxLayout):
    def __init__(self, bg_color=(0.08, 0.12, 0.18, 1), **kwargs):
        super(BackgroundBoxLayout, self).__init__(**kwargs)
        with self.canvas.before:
            Color(*bg_color)
            self.rect = Rectangle(size=self.size, pos=self.pos)
        self.bind(size=self._update_rect, pos=self._update_rect)

    def _update_rect(self, instance, value):
        self.rect.size = instance.size
        self.rect.pos = instance.pos

# 1. GİRİŞ (LOGIN) EKRANI
class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super(LoginScreen, self).__init__(**kwargs)
        layout = BackgroundBoxLayout(orientation='vertical', padding=30, spacing=20, bg_color=(0.06, 0.09, 0.14, 1))
        
        # AAHM emblem uploaded by the user
        logo_path = os.path.join(os.path.dirname(__file__), "aahm_logo.png")
        layout.add_widget(Image(source=logo_path, size_hint_y=None, height=150, allow_stretch=True, keep_ratio=True))
        layout.add_widget(Label(text="AAHM", font_size=32, bold=True, color=(0.9, 0.9, 0.9, 1), size_hint_y=None, height=50))
        layout.add_widget(Label(text="Hərbi İnstitut - Arayış İdarəetmə", font_size=16, color=(0.6, 0.7, 0.8, 1), size_hint_y=None, height=30))
        
        # Boşluq
        layout.add_widget(Label(size_hint_y=None, height=20))
        
        self.user_input = TextInput(hint_text="İstifadəçi adı (admin)", multiline=False, size_hint_y=None, height=50)
        self.pass_input = TextInput(hint_text="Şifrə (admin)", password=True, multiline=False, size_hint_y=None, height=50)
        
        layout.add_widget(self.user_input)
        layout.add_widget(self.pass_input)
        
        self.error_label = Label(text="", color=(1, 0.3, 0.3, 1), size_hint_y=None, height=30)
        layout.add_widget(self.error_label)
        
        btn_login = Button(text="DAXİL OL", size_hint_y=None, height=55, background_color=(0.1, 0.5, 0.8, 1), bold=True)
        btn_login.bind(on_press=self.check_login)
        layout.add_widget(btn_login)
        
        self.add_widget(layout)

    def check_login(self, instance):
        if self.user_input.text == "admin" and self.pass_input.text == "admin":
            self.error_label.text = ""
            self.user_input.text = ""
            self.pass_input.text = ""
            self.manager.current = 'main'
        else:
            self.error_label.text = "Səhv istifadəçi adı və ya şifrə! (admin/admin)"

# 2. ƏSAS EKRAN (KURSANTLAR SİYAHISI VƏ AXTARIŞ)
class MainScreen(Screen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        layout = BackgroundBoxLayout(orientation='vertical', padding=20, spacing=15, bg_color=(0.08, 0.12, 0.18, 1))
        
        top_bar = BoxLayout(size_hint_y=None, height=50, spacing=10)
        top_bar.add_widget(Label(text="Kursant Arayış-Məruzələri", font_size=20, bold=True, color=(1, 1, 1, 1)))
        
        btn_logout = Button(text="Çıxış", size_hint_x=None, width=80, background_color=(0.8, 0.2, 0.2, 1))
        btn_logout.bind(on_press=self.logout)
        top_bar.add_widget(btn_logout)
        layout.add_widget(top_bar)
        
        self.search_input = TextInput(hint_text="🔍 Kursant axtar (Ad, Soyad)...", size_hint_y=None, height=45, multiline=False)
        self.search_input.bind(text=self.filter_list)
        layout.add_widget(self.search_input)
        
        self.scroll = ScrollView()
        self.list_layout = GridLayout(cols=1, spacing=10, size_hint_y=None, padding=5)
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        self.scroll.add_widget(self.list_layout)
        layout.add_widget(self.scroll)
        
        btn_import = Button(text="📁 Word Faylından Kursant Əlavə Et (.docx)", size_hint_y=None, height=55, background_color=(0.1, 0.6, 0.3, 1), bold=True)
        btn_import.bind(on_press=self.go_to_import)
        layout.add_widget(btn_import)
        
        self.add_widget(layout)
        self.update_list()

    def update_list(self, search_text=""):
        self.list_layout.clear_widgets()
        data = load_data()
        for item in data:
            full_name = f"{item.get('soyadi', '')} {item.get('adi', '')} {item.get('ata_adi', '')}"
            if search_text.lower() in full_name.lower() or search_text == "":
                btn = Button(text=f"  👤 {full_name}", size_hint_y=None, height=55, background_color=(0.15, 0.22, 0.32, 1), color=(1,1,1,1))
                btn.bind(on_press=lambda inst, k=item: self.open_detail(k))
                self.list_layout.add_widget(btn)

    def filter_list(self, instance, value):
        self.update_list(value)

    def open_detail(self, kursant_data):
        detail_screen = self.manager.get_screen('detail')
        detail_screen.display_kursant(kursant_data)
        self.manager.current = 'detail'

    def go_to_import(self, instance):
        self.manager.current = 'import'

    def logout(self, instance):
        self.manager.current = 'login'

    def on_enter(self):
        self.update_list()

# 3. WORD FAYLINDAN AVTOMATİK ƏLAVƏ ETMƏ EKRANI
class ImportScreen(Screen):
    def __init__(self, **kwargs):
        super(ImportScreen, self).__init__(**kwargs)
        layout = BackgroundBoxLayout(orientation='vertical', padding=25, spacing=20, bg_color=(0.08, 0.12, 0.18, 1))
        
        layout.add_widget(Label(text="Word Faylından Əlavə Et", font_size=22, bold=True, color=(1,1,1,1), size_hint_y=None, height=40))
        layout.add_widget(Label(text="Word faylını telefonun 'Download' qovluğuna qoyun\nvə faylın adını yazın (məsələn: arayis.docx)", color=(0.7, 0.7, 0.7, 1), size_hint_y=None, height=50))
        
        self.file_input = TextInput(hint_text="Faylın adı (məs: arayis.docx)", multiline=False, size_hint_y=None, height=50)
        layout.add_widget(self.file_input)
        
        self.status_label = Label(text="", color=(0.2, 1, 0.2, 1), size_hint_y=None, height=40)
        layout.add_widget(self.status_label)
        
        btn_parse = Button(text="📥 Faylı Oxu və Əlavə Et", size_hint_y=None, height=55, background_color=(0.1, 0.6, 0.3, 1), bold=True)
        btn_parse.bind(on_press=self.parse_and_save_word)
        layout.add_widget(btn_parse)
        
        btn_back = Button(text="◀ Geri qayıt", size_hint_y=None, height=50, background_color=(0.4, 0.4, 0.4, 1))
        btn_back.bind(on_press=self.go_back)
        layout.add_widget(btn_back)
        
        self.add_widget(layout)

    def parse_and_save_word(self, instance):
        filename = self.file_input.text.strip()
        if not filename:
            self.status_label.color = (1, 0.3, 0.3, 1)
            self.status_label.text = "Zəhmət olmasa fayl adını daxil edin!"
            return
        
        full_path = f"/sdcard/Download/{filename}"
        if not os.path.exists(full_path):
            full_path = filename  # Cari qovluqda yoxla
            if not os.path.exists(full_path):
                self.status_label.color = (1, 0.3, 0.3, 1)
                self.status_label.text = "Fayl tapılmadı! Download qovluğunu yoxlayın."
                return

        try:
            doc = Document(full_path)
            new_data = {
                "kurs": "Hazırlıq kursu",
                "tabor": "Kursant hazırlıq",
                "boluk": "3-cü kursant"
            }
            
            # Word cədvəlini oxuyub məlumatları çıxarırıq
            for table in doc.tables:
                for row in table.rows:
                    cells = row.cells
                    if len(cells) >= 2:
                        key = cells[0].text.strip().replace(':', '').lower()
                        val = cells[1].text.strip()
                        if 'adı' in key and 'ata' not in key and 'soy' not in key:
                            new_data['adi'] = val
                        elif 'soyadı' in key:
                            new_data['soyadi'] = val
                        elif 'ata adı' in key:
                            new_data['ata_adi'] = val
                        elif 'doğulduğu tarix' in key:
                            new_data['dogum_tarixi'] = val
                        elif 'vəsiqə' in key or 'fin' in key:
                            new_data['fin'] = val
                        elif 'əlaqə' in key:
                            new_data['elaqe'] = val
                        elif 'qan qrupu' in key:
                            new_data['qan_qrupu'] = val
                        elif key == 'ata':
                            new_data['ata'] = val
                        elif key == 'ana':
                            new_data['ana'] = val

            # Əgər cədvəldən tapılmasa, standart ad verək
            if not new_data.get('adi'): new_data['adi'] = "Naməlum"
            if not new_data.get('soyadi'): new_data['soyadi'] = "Kursant"

            data = load_data()
            data.append(new_data)
            save_data(data)

            self.status_label.color = (0.2, 1, 0.2, 1)
            self.status_label.text = "Uğurla əlavə olundu!"
            self.file_input.text = ""
        except Exception as e:
            self.status_label.color = (1, 0.3, 0.3, 1)
            self.status_label.text = f"Xəta baş verdi: {str(e)[:30]}"

    def go_back(self, instance):
        self.manager.current = 'main'

# 4. KURSANT DETAL VƏ WORD-Ə YÜKLƏMƏ EKRANI
class DetailScreen(Screen):
    def __init__(self, **kwargs):
        super(DetailScreen, self).__init__(**kwargs)
        layout = BackgroundBoxLayout(orientation='vertical', padding=20, spacing=10, bg_color=(0.08, 0.12, 0.18, 1))
        
        top_bar = BoxLayout(size_hint_y=None, height=50, spacing=10)
        btn_back = Button(text="◀ Geri", size_hint_x=0.3, background_color=(0.3, 0.3, 0.4, 1))
        btn_back.bind(on_press=self.go_back)
        top_bar.add_widget(btn_back)
        top_bar.add_widget(Label(text="Arayış-Məruzə Sənədi", bold=True, color=(1,1,1,1)))
        layout.add_widget(top_bar)
        
        self.scroll = ScrollView()
        self.content_layout = GridLayout(cols=1, spacing=8, size_hint_y=None, padding=5)
        self.content_layout.bind(minimum_height=self.content_layout.setter('height'))
        self.scroll.add_widget(self.content_layout)
        layout.add_widget(self.scroll)
        
        btn_export = Button(text="📥 Word (.docx) Faylına Yüklə", size_hint_y=None, height=55, background_color=(0.8, 0.4, 0.1, 1), bold=True)
        btn_export.bind(on_press=self.export_to_word)
        layout.add_widget(btn_export)
        
        self.add_widget(layout)
        self.current_kursant = {}

    def display_kursant(self, data):
        self.current_kursant = data
        self.content_layout.clear_widgets()
        
        fields = [
            ("Adı", data.get("adi", "")),
            ("Soyadı", data.get("soyadi", "")),
            ("Ata adı", data.get("ata_adi", "")),
            ("Doğulduğu tarix", data.get("dogum_tarixi", "")),
            ("Vəsiqə Seriya və FİN", data.get("fin", "")),
            ("Kurs", data.get("kurs", "Hazırlıq kursu")),
            ("Tabor", data.get("tabor", "Kursant hazırlıq")),
            ("Bölük", data.get("boluk", "3-cü kursant")),
            ("Əlaqə nömrəsi", data.get("elaqe", "")),
            ("Qan qrupu", data.get("qan_qrupu", "")),
            ("Ata", data.get("ata", "")),
            ("Ana", data.get("ana", ""))
        ]
        
        for k, v in fields:
            lbl = Label(text=f"[b]{k}:[/b] {v}", markup=True, size_hint_y=None, height=40, halign='left', valign='middle', color=(0.9, 0.9, 0.9, 1))
            lbl.bind(size=lambda s, w: s.setter('text_size')(s, (w, None)))
            self.content_layout.add_widget(lbl)

    def export_to_word(self, instance):
        doc = Document()
        doc.add_heading('ARAYIŞ-MƏRUZƏ', 0)
        
        d = self.current_kursant
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Parametr'
        hdr_cells[1].text = 'Məlumat'
        
        rows_data = [
            ("Adı", d.get("adi", "")),
            ("Soyadı", d.get("soyadi", "")),
            ("Ata adı", d.get("ata_adi", "")),
            ("Doğulduğu tarix gün,ay,il", d.get("dogum_tarixi", "")),
            ("Vəsiqə Seriya və Fin kodu", d.get("fin", "")),
            ("Kurs", d.get("kurs", "Hazırlıq kursu")),
            ("Tabor", d.get("tabor", "Kursant hazırlıq")),
            ("Bölük", d.get("boluk", "3-cü kursant")),
            ("Əlaqə nömrəsi", d.get("elaqe", "")),
            ("Qan qrupu", d.get("qan_qrupu", "")),
            ("Ata", d.get("ata", "")),
            ("Ana", d.get("ana", ""))
        ]
        
        for key, val in rows_data:
            row_cells = table.add_row().cells
            row_cells[0].text = key
            row_cells[1].text = val
            
        filename = f"/sdcard/Download/Arayis_{d.get('soyadi', 'Kursant')}.docx"
        try:
            doc.save(filename)
        except:
            doc.save("arayis.docx")

    def go_back(self, instance):
        self.manager.current = 'main'

class ArayisApp(App):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(LoginScreen(name='login'))
        sm.add_widget(MainScreen(name='main'))
        sm.add_widget(ImportScreen(name='import'))
        sm.add_widget(DetailScreen(name='detail'))
        return sm

if __name__ == '__main__':
    ArayisApp().run()