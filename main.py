import os
import shutil
import gc
import sys
import csv
import hashlib
import sqlite3
import tempfile
import json
import re
import difflib
import textwrap
from collections import defaultdict, Counter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QDoubleSpinBox, QComboBox,
    QFrame, QAbstractItemView, QStackedWidget, QStyledItemDelegate, QInputDialog,
    QDialog, QScrollArea, QFileDialog, QGridLayout, QCompleter,
    QRadioButton, QButtonGroup, QStackedLayout, QTextEdit, QDateEdit, QProgressBar,
    QFormLayout, QProgressDialog, QListWidget,
    QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt, QTimer, QPoint, QPointF, QEvent, QStringListModel, QRegularExpression, QDate, QThread, pyqtSignal, QRectF
from PyQt6.QtGui import QFont, QColor, QPixmap, QPainter, QPolygon, QPainterPath, QValidator, QTextDocument, QRegularExpressionValidator, QTextCursor, QBrush, QPen, QDoubleValidator
try:
    import openpyxl
    OPENPYXL_VAR = True
except ImportError:
    OPENPYXL_VAR = False
try:
    import cv2
    import pytesseract
    OCR_VAR = True
except ImportError:
    cv2 = None
    pytesseract = None
    OCR_VAR = False
if OCR_VAR:
    if os.path.exists("/opt/homebrew/bin/tesseract"):
        pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
    elif os.path.exists("/usr/local/bin/tesseract"):
        pytesseract.pytesseract.tesseract_cmd = "/usr/local/bin/tesseract"
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from isim_sozlugu import ERKEK_ISIMLERI, SOYISIMLER, ISIM_DUZELTME_HARITASI
from wp_musteri_esleme import MUSTERISIZ, MusteriEsleyici, wp_ek_borc


class UcretsizEvrakOkuyucu:
    """OpenCV + Tesseract: rakam ve Türkçe yazı çapraz doğrulamalı çek/senet okuyucu."""

    @staticmethod
    def goruntuyu_temizle(foto_yolu):
        img = cv2.imread(foto_yolu, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return None
        img_blur = cv2.GaussianBlur(img, (5, 5), 0)
        return cv2.adaptiveThreshold(
            img_blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 11, 2,
        )

    @staticmethod
    def sayiyi_turkce_okunusa_cevir(sayi):
        birler = ["", "BİR", "İKİ", "ÜÇ", "DÖRT", "BEŞ", "ALTI", "YEDİ", "SEKİZ", "DOKUZ"]
        onlar = ["", "ON", "YİRMİ", "OTUZ", "KIRK", "ELLİ", "ALTMIŞ", "YETMİŞ", "SEKSEN", "DOKSAN"]

        def uclu_grup_okunusu(n):
            y, o, b = n // 100, (n % 100) // 10, n % 10
            res = ""
            if y == 1:
                res += "YÜZ"
            elif y > 1:
                res += birler[y] + "YÜZ"
            res += onlar[o] + birler[b]
            return res

        sayi = int(sayi)
        if sayi == 0:
            return "SIFIR"
        milyon = sayi // 1000000
        binler = (sayi % 1000000) // 1000
        kalan = sayi % 1000
        sonuc = ""
        if milyon > 0:
            sonuc += uclu_grup_okunusu(milyon) + "MİLYON"
        if binler == 1:
            sonuc += "BİN"
        elif binler > 1:
            sonuc += uclu_grup_okunusu(binler) + "BİN"
        if kalan > 0:
            sonuc += uclu_grup_okunusu(kalan)
        return sonuc

    @classmethod
    def evrak_analiz_et(cls, foto_yolu):
        sonuclar = {
            "tutar_rakam": 0.0,
            "beklenen_yazi": "",
            "vade_tarihi": "",
            "banka_adi": "",
            "guvenli_mi": False,
            "hata_mesaji": "",
            "ham_metin": "",
        }
        if not OCR_VAR:
            sonuclar["hata_mesaji"] = "OCR Motoru (Tesseract) başlatılamadı. Terminalden 'brew install tesseract tesseract-lang' yazdığınıza emin olun."
            return sonuclar
        try:
            temiz_foto = cls.goruntuyu_temizle(foto_yolu)
            if temiz_foto is None:
                sonuclar["hata_mesaji"] = "Fotoğraf okunamadı."
                return sonuclar
            okunan_metin = pytesseract.image_to_string(temiz_foto, lang="tur+eng", config=r"--oem 3 --psm 6")
        except Exception:
            sonuclar["hata_mesaji"] = "OCR Motoru (Tesseract) başlatılamadı. Terminalden 'brew install tesseract tesseract-lang' yazdığınıza emin olun."
            return sonuclar
        sonuclar["ham_metin"] = okunan_metin
        okunan_metin_upper = okunan_metin.upper()
        fiyat_m = re.search(r"#?\s*(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)\s*(?:#|TL|TRL)", okunan_metin_upper)
        if fiyat_m:
            try:
                ham_rakam = fiyat_m.group(1).replace(".", "").replace(",", ".")
                sonuclar["tutar_rakam"] = float(ham_rakam)
            except ValueError:
                pass
        if sonuclar["tutar_rakam"] > 0:
            beklenen_yazi = cls.sayiyi_turkce_okunusa_cevir(sonuclar["tutar_rakam"])
            sonuclar["beklenen_yazi"] = beklenen_yazi

            def normalize_et(metin):
                return (metin.replace(" ", "").replace("İ", "I").replace("Ü", "U")
                        .replace("Ö", "O").replace("Ş", "S").replace("Ç", "C").replace("Ğ", "G"))

            if normalize_et(beklenen_yazi) in normalize_et(okunan_metin_upper):
                sonuclar["guvenli_mi"] = True
                sonuclar["hata_mesaji"] = "Kusursuz. Rakam ve el yazısı tutarı birebir eşleşti."
            else:
                sonuclar["hata_mesaji"] = (
                    f"Rakam {sonuclar['tutar_rakam']:,.2f} ₺ olarak okundu ancak çekte "
                    f"'{beklenen_yazi}' karşılığı BULUNAMADI!"
                )
        else:
            sonuclar["hata_mesaji"] = "Çek üzerindeki rakam alanı (#...#) net olarak okunamadı."
        tarih_m = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", okunan_metin)
        if tarih_m:
            sonuclar["vade_tarihi"] = f"{tarih_m.group(1)}.{tarih_m.group(2)}.{tarih_m.group(3)}"
        bankalar = [
            "GARANTI", "AKBANK", "IS BANKASI", "YAPI KREDI", "ZIRAAT", "HALKBANK",
            "VAKIFBANK", "QNB", "DENIZBANK", "TEB", "KUVEYT TURK",
        ]
        duz = okunan_metin_upper.replace(" ", "")
        for b in bankalar:
            if b.replace(" ", "") in duz:
                sonuclar["banka_adi"] = b
                break
        return sonuclar


class EvrakEkleDialog(QDialog):
    """Tarih kısıtlaması, büyük harf ve sayı doğrulaması içeren Evrak Ekle/Düzenle penceresi"""

    def __init__(self, parent=None, evrak_id=None):
        super().__init__(parent)
        self.evrak_id = evrak_id
        self.foto_yolu = None
        self.setWindowTitle("Finansal Evrak Düzenle / Yeni Giriş" if evrak_id else "Yeni Finansal Evrak Girişi")
        self.setFixedSize(450, 430)
        form = QFormLayout(self)
        self.cmb_tip = QComboBox()
        self.cmb_tip.addItems(["ÇEK", "SENET"])
        self.dt_vade = QDateEdit()
        self.dt_vade.setCalendarPopup(True)
        self.dt_vade.setDate(QDate.currentDate())
        self.dt_vade.setMinimumDate(QDate.currentDate())
        self.inp_tutar = QLineEdit()
        self.inp_tutar.setPlaceholderText("Örn: 200000 yazın, kendi koysun")
        self.inp_tutar.textChanged.connect(self.format_tutar_input)
        self.inp_sahip = BuyukHarfKutusu("Keşideci veya Müşteri Adı")
        self.cmb_banka = QComboBox()
        self.cmb_banka.setEditable(True)
        self.cmb_banka.addItems([
            "Banka Seçiniz...",
            "Ziraat Bankası",
            "Türkiye İş Bankası",
            "Garanti BBVA",
            "Akbank",
            "Yapı Kredi",
            "VakıfBank",
            "Halkbank",
            "QNB Finansbank",
            "DenizBank",
            "Türk Ekonomi Bankası (TEB)",
            "Kuveyt Türk Katılım Bankası",
            "Albaraka Türk Katılım Bankası",
            "Türkiye Finans Katılım Bankası",
            "Odeabank",
            "Anadolubank",
            "Şekerbank",
            "Fibabanka",
            "Alternatif Bank",
            "Burgan Bank",
            "ICBC Turkey Bank",
        ])
        self.inp_sehir = BuyukHarfKutusu("Örn: İSTANBUL")
        self.btn_tara = QPushButton("📸 FOTOĞRAFTAN OTOMATİK DOLDUR (Yapay Zeka)")
        self.btn_tara.setStyleSheet("background: #8b5cf6; color: white; font-weight: bold; padding: 8px; border-radius: 4px;")
        self.btn_tara.clicked.connect(self.fotograf_tara)
        self.btn_kaydet = QPushButton("KAYDET VE GÜNCELLE")
        self.btn_kaydet.setStyleSheet("background: #10b981; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_kaydet.clicked.connect(self.dogrula_ve_kaydet)
        form.addRow(self.btn_tara)
        form.addRow("Evrak Tipi:", self.cmb_tip)
        form.addRow("Vade Tarihi *:", self.dt_vade)
        form.addRow("Tutar (₺) *:", self.inp_tutar)
        form.addRow("Keşideci / Müşteri:", self.inp_sahip)
        form.addRow("Banka Adı *:", self.cmb_banka)
        form.addRow("Şehir:", self.inp_sehir)
        form.addRow(self.btn_kaydet)
        if self.evrak_id:
            self.verileri_yukle()
        lbl_uyari = QLabel("ℹ️ Yeni girişlerde bugünden önceki tarihler seçilemez.")
        lbl_uyari.setStyleSheet("color: #64748b; font-size: 11px;")
        form.addRow(lbl_uyari)

    def verileri_yukle(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(
            "SELECT evrak_tipi, vade_tarihi, tutar, sahibi_kesideci, banka_adi, sehir FROM finans_evraklar WHERE id = ?",
            (self.evrak_id,),
        )
        row = c.fetchone()
        conn.close()
        if not row:
            return
        tip, vade, tutar, sahip, banka, sehir = row
        self.cmb_tip.setCurrentText(tip or "ÇEK")
        if vade:
            self.dt_vade.setMinimumDate(QDate(2020, 1, 1))
            qd = QDate.fromString(str(vade)[:10], "yyyy-MM-dd")
            if qd.isValid():
                self.dt_vade.setDate(qd)
        self.inp_tutar.setText(str(int(round(float(tutar or 0)))))
        self.inp_sahip.setText(str(sahip or ""))
        idx = self.cmb_banka.findText(banka or "", Qt.MatchFlag.MatchContains)
        if idx >= 0:
            self.cmb_banka.setCurrentIndex(idx)
        elif banka:
            self.cmb_banka.setCurrentText(banka)
        self.inp_sehir.setText(str(sehir or ""))

    def dogrula_ve_kaydet(self):
        tutar_str = self.inp_tutar.text().strip()
        banka_secimi = self.cmb_banka.currentText()
        vade_str = self.dt_vade.date().toString("yyyy-MM-dd")
        tutar_val = turk_para_coz(tutar_str)
        if tutar_val <= 0:
            QMessageBox.warning(self, "Hatalı Tutar", "⚠️ Miktar (Tutar) alanı boş bırakılamaz ve 0'dan büyük olmalıdır!")
            self.inp_tutar.setFocus()
            return
        if not banka_secimi or "Seçiniz" in banka_secimi:
            QMessageBox.warning(self, "Eksik Banka", "⚠️ Banka seçimi zorunludur. Lütfen listeden bir banka seçin!")
            self.cmb_banka.setFocus()
            return
        if not vade_str:
            QMessageBox.warning(self, "Eksik Tarih", "⚠️ Vade tarihi boş bırakılamaz!")
            return
        self.accept()

    def format_tutar_input(self, text):
        self.inp_tutar.blockSignals(True)
        rakamlar = "".join(c for c in text if c.isdigit())
        if not rakamlar:
            self.inp_tutar.setText("")
        else:
            formatli = f"{int(rakamlar):,}".replace(",", ".")
            self.inp_tutar.setText(formatli)
        self.inp_tutar.blockSignals(False)

    def fotograf_tara(self):
        yol, _ = QFileDialog.getOpenFileName(self, "Evrak Fotoğrafı Seç", "", "Resim (*.png *.jpg *.jpeg)")
        if not yol:
            return
        self.foto_yolu = yol
        analiz = UcretsizEvrakOkuyucu.evrak_analiz_et(yol)
        if not analiz["guvenli_mi"]:
            QMessageBox.critical(
                self, "⚠️ SIFIR HATA GÜVENLİK SİSTEMİ",
                f"Uyarı: {analiz['hata_mesaji']}\nLütfen alanları manuel doldurun.",
            )
            return
        QMessageBox.information(self, "✅ Başarılı", "Çapraz doğrulama başarılı.")
        self.inp_tutar.setText(str(int(round(float(analiz["tutar_rakam"] or 0)))))
        if analiz["vade_tarihi"]:
            qd = QDate.fromString(analiz["vade_tarihi"], "dd.MM.yyyy")
            if qd.isValid() and qd >= QDate.currentDate():
                self.dt_vade.setDate(qd)
        if analiz["banka_adi"]:
            idx = self.cmb_banka.findText(analiz["banka_adi"], Qt.MatchFlag.MatchContains)
            if idx >= 0:
                self.cmb_banka.setCurrentIndex(idx)
            else:
                self.cmb_banka.setCurrentText(analiz["banka_adi"])
        self.cmb_tip.setCurrentText("ÇEK")


class TopluEvrakOkuyucuThread(QThread):
    progress_guncelle = pyqtSignal(int, int, str)
    islem_bitti = pyqtSignal(int, list)

    def __init__(self, dosya_yollari):
        super().__init__()
        self.dosya_yollari = dosya_yollari

    def run(self):
        basarili_kayit_sayisi = 0
        hatali_evraklar = []
        conn = sqlite3.connect(DB_NAME, timeout=30)
        c = conn.cursor()
        for i, yol in enumerate(self.dosya_yollari):
            if self.isInterruptionRequested():
                break
            dosya_adi = os.path.basename(yol)
            self.progress_guncelle.emit(i + 1, len(self.dosya_yollari), dosya_adi)
            try:
                analiz = UcretsizEvrakOkuyucu.evrak_analiz_et(yol)
                if analiz["guvenli_mi"]:
                    vade = analiz["vade_tarihi"]
                    if vade:
                        try:
                            vade_dt = datetime.strptime(vade, "%d.%m.%Y").strftime("%Y-%m-%d")
                        except ValueError:
                            vade_dt = datetime.now().strftime("%Y-%m-%d")
                    else:
                        vade_dt = datetime.now().strftime("%Y-%m-%d")
                    c.execute(
                        """INSERT INTO finans_evraklar (evrak_tipi, vade_tarihi, tutar, banka_adi, foto_yolu, durum)
                           VALUES ('ÇEK', ?, ?, ?, ?, 'BEKLİYOR')""",
                        (vade_dt, analiz["tutar_rakam"], analiz["banka_adi"], yol),
                    )
                    basarili_kayit_sayisi += 1
                else:
                    hatali_evraklar.append((dosya_adi, analiz["hata_mesaji"]))
            except Exception as e:
                hatali_evraklar.append((dosya_adi, f"Beklenmeyen Hata: {str(e)}"))
        conn.commit()
        conn.close()
        self.islem_bitti.emit(basarili_kayit_sayisi, hatali_evraklar)


class TopluSonucDialog(QDialog):
    """Hangi çekin neresinde hata/uyuşmazlık olduğunu tek tek gösteren detaylı rapor penceresi"""

    def __init__(self, basarili_sayisi, hatali_listesi, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Toplu Evrak İşlem Raporu & Detaylar")
        self.resize(600, 450)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        lbl_baslik = QLabel(
            f"📊 Toplu Yükleme Tamamlandı\n"
            f"✅ Başarılı Kaydedilen: {basarili_sayisi} | ⚠️ İncelenmesi Gereken: {len(hatali_listesi)}"
        )
        lbl_baslik.setStyleSheet("font-size: 13.5px; font-weight: bold; color: #1e293b;")
        layout.addWidget(lbl_baslik)
        self.tbl_detay = QTableWidget()
        self.tbl_detay.setColumnCount(2)
        self.tbl_detay.setHorizontalHeaderLabels(["Dosya Adı", "Hata veya Uyuşmazlık Nedeni"])
        self.tbl_detay.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.tbl_detay.setColumnWidth(0, 160)
        self.tbl_detay.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.tbl_detay.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl_detay.setRowCount(len(hatali_listesi))
        for r, (dosya, sebep) in enumerate(hatali_listesi):
            self.tbl_detay.setItem(r, 0, QTableWidgetItem(dosya))
            item_sebep = QTableWidgetItem(sebep)
            item_sebep.setForeground(QColor("#dc2626"))
            self.tbl_detay.setItem(r, 1, item_sebep)
        layout.addWidget(self.tbl_detay)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_kapat = QPushButton("✖ İptal Et / Kapat")
        self.btn_kapat.setFixedHeight(36)
        self.btn_kapat.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_kapat.setStyleSheet(
            "QPushButton { background-color: #f1f5f9; color: #334155; font-weight: bold; "
            "border: 1px solid #cbd5e1; border-radius: 6px; padding: 0 20px; }"
            "QPushButton:hover { background-color: #e2e8f0; }"
        )
        self.btn_kapat.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_kapat)
        layout.addLayout(btn_layout)


class EmailAyarlariDialog(QDialog):
    """Varsayılan Gmail ayarları otomatik gelen E-Posta Bildirim Penceresi"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("E-Posta Bildirim ve Alıcı Ayarları")
        self.setFixedSize(500, 520)
        self.layout = QFormLayout(self)

        self.inp_sunucu = QLineEdit("smtp.gmail.com")
        self.inp_port = QLineEdit("587")

        self.inp_gonderen = QLineEdit()
        self.inp_gonderen.setPlaceholderText("Sizin e-posta adresiniz (Örn: ad soyad@gmail.com)")

        self.inp_sifre = QLineEdit()
        self.inp_sifre.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_sifre.setPlaceholderText("Gmail 16 haneli uygulama şifresi")

        self.alicilar = []
        for i in range(1, 6):
            inp = QLineEdit()
            inp.setPlaceholderText(f"{i}. Alıcı E-Posta Adresi (İsteğe bağlı)")
            self.alicilar.append(inp)

        self.btn_kaydet = QPushButton("KAYDET")
        self.btn_kaydet.setStyleSheet("background: #10b981; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_kaydet.clicked.connect(self.ayarlari_kaydet)

        self.layout.addRow(QLabel("<b>⚙️ Varsayılan Ayarlar Otomatik Tanımlıdır</b>"))
        self.layout.addRow("Gönderen E-Posta:", self.inp_gonderen)
        self.layout.addRow("E-Posta Şifresi:", self.inp_sifre)
        self.layout.addRow(QLabel("<b>--- Bildirim Gidecek Alıcılar (En Fazla 5) ---</b>"))

        for idx, inp in enumerate(self.alicilar, 1):
            self.layout.addRow(f"Alıcı {idx}:", inp)

        self.layout.addRow(self.btn_kaydet)
        self.verileri_yukle()

    def verileri_yukle(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT smtp_sunucu, smtp_port, gonderen_email, gonderen_sifre, alici_1, alici_2, alici_3, alici_4, alici_5 FROM email_ayarlari LIMIT 1")
        row = c.fetchone()
        conn.close()
        if row:
            if row[0]: self.inp_sunucu.setText(row[0])
            if row[1]: self.inp_port.setText(str(row[1]))
            self.inp_gonderen.setText(row[2] or "")
            self.inp_sifre.setText(row[3] or "")
            for i in range(5):
                if row[4+i]:
                    self.alicilar[i].setText(row[4+i])

    def ayarlari_kaydet(self):
        sunucu = self.inp_sunucu.text().strip() or "smtp.gmail.com"
        port = self.inp_port.text().strip() or "587"
        gonderen = self.inp_gonderen.text().strip()
        sifre = self.inp_sifre.text().strip()

        a1 = self.alicilar[0].text().strip()
        a2 = self.alicilar[1].text().strip()
        a3 = self.alicilar[2].text().strip()
        a4 = self.alicilar[3].text().strip()
        a5 = self.alicilar[4].text().strip()

        if not gonderen or not sifre:
            QMessageBox.warning(self, "Eksik Bilgi", "Lütfen gönderen e-posta adresinizi ve şifrenizi giriniz.")
            return

        try:
            port_val = int(port)
        except:
            port_val = 587

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM email_ayarlari")
        c.execute("""
            INSERT INTO email_ayarlari (smtp_sunucu, smtp_port, gonderen_email, gonderen_sifre, alici_1, alici_2, alici_3, alici_4, alici_5)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (sunucu, port_val, gonderen, sifre, a1, a2, a3, a4, a5))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "Başarılı", "E-posta bildirim ayarları başarıyla kaydedildi.")
        self.accept()


def vade_hatirlatici_mailleri_gonder():
    """Tarih formatı esnek, hem açılışta hem yeni çek girildiğinde çalışan kontrol"""
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()

        c.execute("SELECT smtp_sunucu, smtp_port, gonderen_email, gonderen_sifre, alici_1, alici_2, alici_3, alici_4, alici_5 FROM email_ayarlari LIMIT 1")
        ayar = c.fetchone()
        if not ayar or not ayar[0] or not ayar[2] or not ayar[3]:
            print("⚠️ E-posta ayarları eksik veya şifre girilmemiş.")
            conn.close()
            return

        sunucu, port, gonderen, sifre = ayar[0], int(ayar[1] or 587), ayar[2], ayar[3]
        alicilar = [a.strip() for a in ayar[4:9] if a and "@" in a]
        if not alicilar:
            print("⚠️ Kayıtlı alıcı e-posta adresi bulunamadı.")
            conn.close()
            return

        bugun = datetime.now().date()

        c.execute("SELECT id, evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi_kesideci, foto_yolu FROM finans_evraklar WHERE UPPER(durum) = 'BEKLİYOR'")
        evraklar = c.fetchall()
        conn.close()

        secilen_evraklar = []
        ek_fotolar = []
        for e_id, tip, vade, tutar, banka, sahip, foto_yolu in evraklar:
            if not vade:
                continue

            vade_dt = None
            vade_str = str(vade).strip()
            for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
                try:
                    vade_dt = datetime.strptime(vade_str, fmt).date()
                    break
                except ValueError:
                    pass

            if not vade_dt:
                print(f"⚠️ Geçersiz tarih formatı (ID: {e_id}): {vade_str}")
                continue

            kalan_gun = (vade_dt - bugun).days
            if kalan_gun in [0, 1, 7]:
                secilen_evraklar.append((tip, vade_dt.strftime("%d.%m.%Y"), tutar, banka, sahip, kalan_gun))
                if foto_yolu and os.path.exists(foto_yolu):
                    ek_fotolar.append(foto_yolu)

        if not secilen_evraklar:
            print("ℹ️ Vadesi bugün, yarın veya 7 gün sonra olan bekleyen evrak yok.")
            return

        konu = f"🚨 Vade Bildirimi: {len(secilen_evraklar)} Adet Evrak"
        govde = "Vadesi yaklaşan veya bugün dolan evraklar:\n\n"

        for tip, v_tar, tutar, banka, sahip, k_gun in secilen_evraklar:
            durum_metni = "🚨 BUGÜN VADESİ DOLUYOR!" if k_gun == 0 else (f"⏳ Yarın ({v_tar})" if k_gun == 1 else f"📅 {k_gun} Gün Kaldı ({v_tar})")
            govde += f"• [{durum_metni}] {tip} | Tutar: {float(tutar or 0):,.2f} ₺ | Banka: {banka} | Keşideci: {sahip}\n"

        govde += "\nBenimPOS Otomasyon Sistemi"

        msg = MIMEMultipart()
        msg['From'] = gonderen
        msg['Subject'] = konu
        msg.attach(MIMEText(govde, 'plain', 'utf-8'))
        for resim_yolu in ek_fotolar:
            with open(resim_yolu, "rb") as f:
                img_eki = MIMEImage(f.read())
                img_eki.add_header("Content-Disposition", "attachment", filename=os.path.basename(resim_yolu))
                msg.attach(img_eki)

        server = smtplib.SMTP(sunucu, port, timeout=10)
        server.starttls()
        server.login(gonderen, sifre)
        server.sendmail(gonderen, alicilar, msg.as_string())
        server.quit()

        print(f"✅ Vade hatırlatma maili {len(alicilar)} alıcıya başarıyla gönderildi!")

    except Exception as e:
        print(f"❌ Mail Gönderme Hatası: {e}")


_ok_ikonlari = {}

def ok_ikonu(yon="asagi", renk="#6c757d"):
    """Üçgen ok görseli üretir ve dosya yolunu döndürür.

    Qt'nin stil motoru CSS üçgen hilesini (0 boyut + saydam kenarlık)
    desteklemiyor; ayrıca bir alt kontrol (::drop-down, ::up-button) stille
    özelleştirildiğinde yerel ok da çizilmiyor. Bu yüzden oku bir kez çizip
    stile görsel olarak veriyoruz. 2 kat çözünürlükte çizilir ki retina
    ekranda keskin görünsün.
    """
    anahtar = (yon, renk)
    if anahtar not in _ok_ikonlari:
        pix = QPixmap(20, 12)
        pix.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(renk))
        if yon == "asagi":
            ucgen = [QPoint(0, 0), QPoint(20, 0), QPoint(10, 12)]
        else:
            ucgen = [QPoint(0, 12), QPoint(20, 12), QPoint(10, 0)]
        painter.drawPolygon(QPolygon(ucgen))
        painter.end()
        yol = os.path.join(tempfile.gettempdir(), f"plastikpos_ok_{yon}_{renk.lstrip('#')}.png")
        pix.save(yol)
        _ok_ikonlari[anahtar] = yol
    return _ok_ikonlari[anahtar]

def buyuk_harf(metin):
    """Türkçe'ye uygun büyük harf. Python'un upper() metodu 'i' harfini 'I'
    yapar; veritabanındaki 'SİYAH POM' gibi kayıtlarla tutarlı olmak için
    'i' harfini önce 'İ'ye çeviriyoruz.
    """
    return metin.replace("i", "İ").upper()


def turk_para_coz(t_str):
    """'125.000', '200.000' veya '1.250.000,50' gibi Türkçe formatları float'a çevirir."""
    if not t_str:
        return 0.0
    t = str(t_str).replace("₺", "").replace("TL", "").replace(" ", "").strip()
    if not t:
        return 0.0
    t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def format_tl(deger):
    """125000.0 → '125.000,00 ₺'"""
    try:
        val = float(deger or 0.0)
        s = f"{val:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{s} ₺"
    except (TypeError, ValueError):
        return "0,00 ₺"


def para_yazisi_olustur(sayi):
    """Toplam tutarı yazıyla ifade eder."""
    try:
        sayi = float(sayi or 0.0)
        if sayi <= 0:
            return "SIFIR TÜRK LİRASI"
        tam = int(sayi)
        kurus = int(round((sayi - tam) * 100))
        yazi = UcretsizEvrakOkuyucu.sayiyi_turkce_okunusa_cevir(tam) + " TÜRK LİRASI"
        if kurus > 0:
            yazi += f" VE {UcretsizEvrakOkuyucu.sayiyi_turkce_okunusa_cevir(kurus)} KURUŞ"
        return yazi
    except (TypeError, ValueError):
        return ""

class BuyukHarfDogrulayici(QValidator):
    """Yazılan metni anında büyük harfe çevirir.

    textChanged + setText yerine doğrulayıcı kullanılıyor; setText imleci
    metnin sonuna atladığı için ortadaki bir harfi düzeltmek imkânsız hale
    geliyordu.
    """
    def validate(self, metin, konum):
        return (QValidator.State.Acceptable, buyuk_harf(metin), konum)

class BuyukHarfKutusu(QLineEdit):
    """Yazılan her harfi anında büyük harfe çeviren metin kutusu."""
    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setValidator(BuyukHarfDogrulayici(self))

class SadeceSayiKutusu(QLineEdit):
    """Harf kabul etmez; silik 0.00 placeholder; 1000 → 1.000; fiyat modunda 1 → 1,0."""
    def __init__(self, fiyat_modu=False, parent=None):
        super().__init__(parent)
        self.fiyat_modu = fiyat_modu
        self.setPlaceholderText("0.00")
        regex = QRegularExpression(r"^[0-9.,]*$")
        self.setValidator(QRegularExpressionValidator(regex, self))
        self.textEdited.connect(self.otomatik_formatla)
        self.editingFinished.connect(self.tamamla_format)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        QTimer.singleShot(0, self.selectAll)

    def otomatik_formatla(self, text):
        temiz = text.replace(" ", "").replace("₺", "").strip()
        if not temiz:
            return

        ayrac = "," if "," in temiz else ("." if "." in temiz and temiz.count(".") == 1 and len(temiz.split(".")[1]) <= 2 else None)

        if ayrac:
            tam_kisim = temiz.split(ayrac)[0].replace(".", "")
            ondalik = temiz.split(ayrac)[1]
        else:
            tam_kisim = temiz.replace(".", "")
            ondalik = None

        if tam_kisim.isdigit():
            formatli_tam = f"{int(tam_kisim):,}".replace(",", ".")
            yeni = f"{formatli_tam},{ondalik}" if ondalik is not None else formatli_tam
            cp = self.cursorPosition()
            self.blockSignals(True)
            self.setText(yeni)
            self.setCursorPosition(min(len(yeni), cp + 1))
            self.blockSignals(False)

    def tamamla_format(self):
        val = self.sayi_al()
        if val > 0:
            if self.fiyat_modu:
                txt = self.text().strip()
                if "," not in txt and "." not in txt:
                    self.setText(f"{txt},0")
            self.otomatik_formatla(self.text())

    def sayi_al(self):
        txt = self.text().replace(".", "").replace(",", ".").replace("₺", "").replace(" ", "").strip()
        try:
            return float(txt) if txt else 0.0
        except ValueError:
            return 0.0

def turkce_toleransli_metin(text):
    """ı/i, ş/s, ğ/g, ü/u, ö/o, ç/c farklarını silerek esnek arama yapar"""
    if not text:
        return ""
    text = text.replace("İ", "i").replace("I", "ı").lower()
    harf_haritasi = {"ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c"}
    for tr, en in harf_haritasi.items():
        text = text.replace(tr, en)
    return text

class SilikSifirliBinlikKutu(QLineEdit):
    """0'ı silik placeholder yapar; basılan rakamları anında 1.000, 10.000 yapar"""
    def __init__(self, birim_eki="", parent=None):
        super().__init__(parent)
        self.birim_eki = birim_eki
        self.setPlaceholderText(f"0 {birim_eki}".strip())
        self.textEdited.connect(self.formatla)

    def formatla(self, metin):
        temiz = metin.replace(".", "").replace(self.birim_eki, "").replace(" ", "").strip()
        if not temiz:
            self.setText("")
            return

        if "," in temiz:
            tam, ondalik = temiz.split(",", 1)
            ondalik = "".join(ch for ch in ondalik if ch.isdigit())
        else:
            tam, ondalik = temiz, None

        if not tam:
            tam = "0"
        if tam.isdigit():
            sayi = int(tam)
            formatli = f"{sayi:,}".replace(",", ".")
            sonuc = f"{formatli},{ondalik}" if ondalik is not None else formatli
            if self.birim_eki:
                sonuc += f" {self.birim_eki}"
            pos = self.cursorPosition()
            self.setText(sonuc)
            self.setCursorPosition(min(len(sonuc), pos + 1))

    def deger_yaz(self, sayi):
        if not sayi:
            self.setText("")
            return
        if abs(sayi - int(sayi)) < 1e-9:
            yazi = f"{int(round(sayi)):,}".replace(",", ".")
        else:
            tam, ondalik = f"{sayi:.2f}".split(".")
            yazi = f"{int(tam):,}".replace(",", ".") + "," + ondalik
        if self.birim_eki:
            yazi += f" {self.birim_eki}"
        self.blockSignals(True)
        self.setText(yazi)
        self.blockSignals(False)

    def sayisal_deger(self):
        t = self.text().replace(".", "").replace(self.birim_eki, "").replace(",", ".").replace(" ", "").strip()
        try:
            return float(t) if t else 0.0
        except ValueError:
            return 0.0

class AkilliSayiKutusu(QDoubleSpinBox):
    def __init__(self, decimals=2, max_val=10000000.0, birlesik=False):
        super().__init__()
        self.setRange(0.0, max_val)
        self.setDecimals(decimals)
        # Ok butonlarını görünür yapar (Fotoğraftaki yukarı/aşağı oklar)
        self.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.UpDownArrows)
        self.lineEdit().setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setStyleSheet(self._stil(birlesik))

    def _stil(self, birlesik):
        """macOS'un okları yok etmesini engelleyen net yukarı/aşağı ok stili.

        birlesik=True olduğunda kutu, solundaki para birimi kutusuna yapışır
        (sol kenarlık yok, sadece sağ köşeler yuvarlak).
        """
        if birlesik:
            kenar = """
                border: 1px solid #ced4da;
                border-left: none;
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
            """
        else:
            kenar = """
                border: 1px solid #ced4da;
                border-radius: 4px;
            """
        return ("""
            QDoubleSpinBox {
                KENARLIK
                padding-left: 6px;
                padding-right: 20px;
                font-size: 13px;
                background-color: #ffffff;
            }
            QDoubleSpinBox::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 18px;
                border-left: 1px solid #ced4da;
                border-bottom: 1px solid #ced4da;
                background-color: #f8fafc;
            }
            QDoubleSpinBox::up-arrow {
                image: url(YUKARI_OK);
                width: 8px;
                height: 5px;
            }
            QDoubleSpinBox::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 18px;
                border-left: 1px solid #ced4da;
                background-color: #f8fafc;
            }
            QDoubleSpinBox::down-arrow {
                image: url(ASAGI_OK);
                width: 8px;
                height: 5px;
            }
            QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover {
                background-color: #e2e8f0;
            }
        """.replace("KENARLIK", kenar.strip())
           .replace("YUKARI_OK", ok_ikonu("yukari", "#555555"))
           .replace("ASAGI_OK", ok_ikonu("asagi", "#555555")))

    def focusInEvent(self, event):
        super().focusInEvent(event)
        QTimer.singleShot(0, self.lineEdit().selectAll)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        QTimer.singleShot(0, self.lineEdit().selectAll)

POPUP_MIN_SATIR_YUKSEKLIGI = 36

# macOS'ta açılır kutunun satırlarının sıkışmasını önler.
# Yükseklik sabit değil alt sınırdır; içerik daha fazlasını gerektirirse büyüyebilir.
class GenisPopupDelegesi(QStyledItemDelegate):
    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        size.setHeight(max(size.height(), POPUP_MIN_SATIR_YUKSEKLIGI))
        return size

# (etiket, kod, simge) - açılır kutu ve tablo aynı listeyi kullanır
CURRENCIES = [
    ("TRY ₺", "TRY", "₺"),
    ("USD $", "USD", "$"),
    ("EUR €", "EUR", "€"),
    ("GBP £", "GBP", "£"),
    ("AZN ₼", "AZN", "₼"),
    ("JPY ¥", "JPY", "¥"),
    ("CHF", "CHF", "CHF"),
    ("CAD $", "CAD", "$"),
    ("RUB ₽", "RUB", "₽"),
    ("CNY ¥", "CNY", "¥"),
]
CURRENCY_SYMBOLS = {code: symbol for _, code, symbol in CURRENCIES}
DEFAULT_CURRENCY = "TRY"

# (etiket, kod) - 10 popüler ölçüm birimi
UNITS = [
    ("KG", "KG"),
    ("TON", "TON"),
    ("LİTRE", "LT"),
    ("METRE", "M"),
    ("ADET", "AD"),
    ("RULO", "RL"),
    ("TORBA", "TRB"),
    ("BIGBAG", "BB"),
    ("BALYA", "BLY"),
    ("M² (Kare)", "M2"),
]
DEFAULT_UNIT = "KG"

def para_birimi_kutusu():
    """En çok kullanılan 10 para birimini açılır kutu yapar.

    Fiyat kutusunun soluna yapıştığı için sadece sol köşeleri yuvarlak.
    """
    combo = QComboBox()
    combo.setFixedHeight(38)
    combo.setFixedWidth(90)
    for label, code, _ in CURRENCIES:
        combo.addItem(label, code)

    combo.setItemDelegate(GenisPopupDelegesi(combo))
    combo.setStyleSheet("""
        QComboBox {
            border: 1px solid #cbd5e1;
            border-top-left-radius: 5px;
            border-bottom-left-radius: 5px;
            border-top-right-radius: 0px;
            border-bottom-right-radius: 0px;
            background-color: #f8fafc;
            font-weight: bold;
            font-size: 15px;
            padding-left: 6px;
            color: #1e293b;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 22px;
            border-left: 1px solid #cbd5e1;
        }
        QComboBox::down-arrow {
            image: url(OK_IKONU);
            width: 8px;
            height: 5px;
        }
        QComboBox QAbstractItemView {
            border: 1px solid #cbd5e1;
            background-color: white;
            selection-background-color: #0284c7;
            selection-color: white;
            font-size: 15px;
            font-weight: bold;
            min-width: 120px;
            padding: 4px;
        }
    """.replace("OK_IKONU", ok_ikonu()))
    return combo

def combo_secimi_ayarla(combo, value):
    idx = combo.findData(value)
    combo.setCurrentIndex(idx if idx >= 0 else 0)

DB_NAME = "pos_data.db"


def guncel_fiyati_getir(urun_adi):
    """Son 4 aydaki en son satış fiyatını döner; yoksa None."""
    if not urun_adi:
        return None
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""
            SELECT price, created_at
            FROM sales_history
            WHERE product_name = ? COLLATE NOCASE AND price > 0
            ORDER BY id DESC
            LIMIT 80
        """, (urun_adi,))
        satirlar = c.fetchall()
        conn.close()
    except Exception:
        return None
    sinir = datetime.now() - timedelta(days=120)
    for fiyat, tarih in satirlar:
        ham = str(tarih or "").strip().split(" ")[0]
        dt = None
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d.%m.%y"):
            try:
                dt = datetime.strptime(ham, fmt)
                break
            except ValueError:
                continue
        if dt is None:
            continue
        if dt >= sinir:
            return float(fiyat)
    return None


def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS product_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            stock_kg REAL DEFAULT 0.0,
            buy_price REAL DEFAULT 0.0,
            sell_price REAL DEFAULT 0.0,
            buy_currency TEXT DEFAULT 'TRY',
            sell_currency TEXT DEFAULT 'TRY',
            unit TEXT DEFAULT 'KG',
            FOREIGN KEY (group_id) REFERENCES product_groups(id) ON DELETE CASCADE
        )
    """)

    c.execute("PRAGMA table_info(products)")
    mevcut_kolonlar = [row[1] for row in c.fetchall()]
    for kolon, varsayilan in (("buy_currency", "TRY"), ("sell_currency", "TRY"), ("unit", "KG")):
        if kolon not in mevcut_kolonlar:
            c.execute(f"ALTER TABLE products ADD COLUMN {kolon} TEXT DEFAULT '{varsayilan}'")

    c.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            shopping_count INTEGER DEFAULT 0,
            debt REAL DEFAULT 0.0,
            payment REAL DEFAULT 0.0,
            remaining_debt REAL DEFAULT 0.0,
            last_payment_date TEXT,
            detail TEXT,
            address TEXT,
            phone TEXT,
            tax_office TEXT,
            tax_number TEXT,
            term_days INTEGER DEFAULT 0,
            credit_limit REAL DEFAULT 0.0,
            note TEXT,
            latitude TEXT,
            longitude TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS sales_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            personnel_name TEXT DEFAULT 'Genel',
            customer_id INTEGER,
            customer_name TEXT,
            product_name TEXT,
            qty REAL DEFAULT 0.0,
            price REAL DEFAULT 0.0,
            total_amount REAL,
            payment_type TEXT,
            payment_received REAL DEFAULT 0.0,
            note TEXT,
            items_json TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.execute("PRAGMA table_info(sales_history)")
    satis_kolonlar = [row[1] for row in c.fetchall()]
    for kolon, tanim in (
        ("personnel_name", "TEXT DEFAULT 'Genel'"),
        ("product_name", "TEXT"),
        ("qty", "REAL DEFAULT 0.0"),
        ("price", "REAL DEFAULT 0.0"),
        ("payment_received", "REAL DEFAULT 0.0"),
        ("note", "TEXT"),
        ("items_json", "TEXT"),
    ):
        if kolon not in satis_kolonlar:
            c.execute(f"ALTER TABLE sales_history ADD COLUMN {kolon} {tanim}")
    c.execute("""
        CREATE TABLE IF NOT EXISTS wp_kaynak_olaylar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kaynak_anahtar TEXT NOT NULL UNIQUE,
            tarih TEXT,
            personel TEXT,
            ham_govde TEXT,
            ham_musteri_aday TEXT,
            belirsizlik TEXT,
            kalem_no INTEGER DEFAULT 0,
            sales_history_id INTEGER,
            tutar REAL,
            tahsilat REAL,
            durum TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS finans_evraklar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            evrak_tipi TEXT,
            vade_tarihi DATE,
            tutar REAL,
            sahibi_kesideci TEXT,
            banka_adi TEXT,
            sehir TEXT,
            foto_yolu TEXT,
            durum TEXT DEFAULT 'BEKLİYOR',
            eklenme_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.execute("PRAGMA table_info(finans_evraklar)")
    evrak_kolonlar = [row[1] for row in c.fetchall()]
    if "customer_id" not in evrak_kolonlar:
        c.execute("ALTER TABLE finans_evraklar ADD COLUMN customer_id INTEGER")
    c.execute('''
        CREATE TABLE IF NOT EXISTS email_ayarlari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            smtp_sunucu TEXT,
            smtp_port INTEGER,
            gonderen_email TEXT,
            gonderen_sifre TEXT,
            alici_1 TEXT,
            alici_2 TEXT,
            alici_3 TEXT,
            alici_4 TEXT,
            alici_5 TEXT
        )
    ''')
    c.execute("""
        CREATE TABLE IF NOT EXISTS personnel (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            full_name TEXT,
            phone TEXT,
            role TEXT DEFAULT 'Satış Temsilcisi',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.execute("PRAGMA table_info(customers)")
    musteri_kolonlar = [row[1] for row in c.fetchall()]
    for kolon, tanim in (
        ("shopping_count", "INTEGER DEFAULT 0"),
        ("debt", "REAL DEFAULT 0.0"),
        ("payment", "REAL DEFAULT 0.0"),
        ("remaining_debt", "REAL DEFAULT 0.0"),
        ("last_payment_date", "TEXT"),
        ("detail", "TEXT"),
        ("address", "TEXT"),
        ("phone", "TEXT"),
        ("tax_office", "TEXT"),
        ("tax_number", "TEXT"),
        ("term_days", "INTEGER DEFAULT 0"),
        ("credit_limit", "REAL DEFAULT 0.0"),
        ("note", "TEXT"),
        ("latitude", "TEXT"),
        ("longitude", "TEXT"),
    ):
        if kolon not in musteri_kolonlar:
            c.execute(f"ALTER TABLE customers ADD COLUMN {kolon} {tanim}")
    if "balance" in musteri_kolonlar:
        c.execute("""
            UPDATE customers
            SET remaining_debt = COALESCE(balance, 0)
            WHERE COALESCE(remaining_debt, 0) = 0 AND COALESCE(balance, 0) != 0
        """)

    conn.commit()

    varsayilan_gruplar = ["GRUPSUZLAR", "POM", "ABS", "HDPE", "PP MOBLEN", "PA6", "PA66", "ANTİŞOK", "PVC"]
    for g in varsayilan_gruplar:
        c.execute("INSERT OR IGNORE INTO product_groups (name) VALUES (?)", (g,))

    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        c.execute("SELECT id FROM product_groups WHERE name = 'POM' LIMIT 1")
        pom_row = c.fetchone()
        pom_id = pom_row[0] if pom_row else None
        if pom_id:
            c.execute("""
                INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price)
                VALUES (?, 'BEYAZ POM', 12500.0, 27.00, 37.00)
            """, (pom_id,))
            c.execute("""
                INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price)
                VALUES (?, 'SİYAH POM', 8200.0, 25.00, 35.00)
            """, (pom_id,))
        conn.commit()
    conn.close()

class TahminliMusteriKutusu(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.musteri_verileri = []  # [(id, gorunen_ad), ...]

        self.model_liste = QStringListModel()
        self.completer_obj = QCompleter(self.model_liste, self)
        self.completer_obj.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.completer_obj.setFilterMode(Qt.MatchFlag.MatchContains)
        self.setCompleter(self.completer_obj)
        self.lineEdit().setPlaceholderText("🔍 Müşteri adı veya firma yazın...")
        self.lineEdit().textEdited.connect(self.akilli_filtrele)

    def veri_yukle(self, musteriler):
        """Müşteri listesini yükler [(id, name), ...]"""
        self.musteri_verileri = musteriler
        self.blockSignals(True)
        self.clear()

        isimler = [f"👤 {m[1]}" for m in musteriler]
        self.model_liste.setStringList(isimler)

        self.addItem("Lütfen Müşteri Seçiniz veya Yazınız...", None)
        for mid, name in musteriler:
            self.addItem(f"👤 {name}", mid)
        self.blockSignals(False)

    def akilli_filtrele(self, girilen_metin):
        temiz = girilen_metin.replace("👤", "").strip().lower()
        if not temiz or len(temiz) < 2:
            return
        eslesenler = []
        for mid, name in self.musteri_verileri:
            if temiz in name.lower():
                eslesenler.append(f"👤 {name}")
        if len(eslesenler) < 3:
            tum_isimler = [m[1] for m in self.musteri_verileri]
            yakin_tahminler = difflib.get_close_matches(buyuk_harf(temiz), tum_isimler, n=5, cutoff=0.45)
            for yk in yakin_tahminler:
                fmt = f"👤 {yk}"
                if fmt not in eslesenler:
                    eslesenler.append(fmt)
        if eslesenler:
            self.model_liste.setStringList(eslesenler)
            self.completer_obj.complete()


# --- 5. ve 7. FOTOĞRAFTAKİ ÜRÜN EKLEME & DÜZENLEME PENCERESİ ---
class UrunDuzenleDialog(QDialog):
    def __init__(self, parent=None, product_id=None):
        super().__init__(parent)
        self.product_id = product_id
        self.setWindowTitle("ÜRÜN BİLGİLERİ" if not product_id else "ÜRÜN DÜZENLE")
        self.setFixedWidth(660)
        self.setStyleSheet("background-color: #ffffff;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(14)

        # Başlık
        baslik_text = "YENİ ÜRÜN EKLE" if not self.product_id else "ÜRÜN GÜNCELLE"
        lbl_baslik = QLabel(baslik_text)
        lbl_baslik.setStyleSheet("font-size: 17px; font-weight: bold; color: #1e293b; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px;")
        layout.addWidget(lbl_baslik)

        # 1. Ürün Grubu (Büyük Yazılı)
        layout.addWidget(QLabel("ÜRÜN GRUBU:", styleSheet="font-size: 13px; font-weight: bold; color: #334155;"))
        self.combo_grp = QComboBox()
        self.combo_grp.setFixedHeight(38)
        self.combo_grp.setItemDelegate(GenisPopupDelegesi(self.combo_grp))
        self.combo_grp.setStyleSheet("""
            QComboBox {
                border: 1px solid #cbd5e1; border-radius: 5px; padding: 4px 10px;
                font-size: 15px; font-weight: bold; background-color: #f8fafc; color: #1e293b;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 26px;
                border-left: 1px solid #cbd5e1;
            }
            QComboBox::down-arrow {
                image: url(OK_IKONU);
                width: 10px;
                height: 6px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #cbd5e1; background-color: #ffffff;
                selection-background-color: #0284c7; selection-color: #ffffff;
                font-size: 15px; font-weight: bold; min-width: 250px; padding: 4px;
            }
        """.replace("OK_IKONU", ok_ikonu()))
        # Grupları doldur
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name FROM product_groups ORDER BY name")
        for g_id, g_name in c.fetchall():
            self.combo_grp.addItem(g_name, g_id)
        conn.close()
        layout.addWidget(self.combo_grp)

        # 2. Hammadde / Çeşit Adı (Büyük Harf)
        layout.addWidget(QLabel("HAMMADDE / ÇEŞİT ADI:", styleSheet="font-size: 13px; font-weight: bold; color: #334155;"))
        self.txt_ad = BuyukHarfKutusu("ÖRN: BEYAZ POM, ŞEFFAF PP")
        self.txt_ad.setFixedHeight(38)
        self.txt_ad.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 5px; padding: 0 10px; font-size: 14px; font-weight: bold;")
        layout.addWidget(self.txt_ad)

        # 3. Birim
        layout.addWidget(QLabel("ÖLÇÜ BİRİMİ:", styleSheet="font-size: 13px; font-weight: bold; color: #334155;"))
        self.combo_birim = QComboBox()
        self.combo_birim.setFixedHeight(38)
        self.combo_birim.setItemDelegate(GenisPopupDelegesi(self.combo_birim))
        self.combo_birim.setStyleSheet("""
            QComboBox {
                border: 1px solid #cbd5e1; border-radius: 5px; padding: 4px 10px;
                font-size: 15px; font-weight: bold; background-color: #ffffff; color: #1e293b;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 26px;
                border-left: 1px solid #cbd5e1;
            }
            QComboBox::down-arrow {
                image: url(OK_IKONU);
                width: 10px;
                height: 6px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #cbd5e1; background-color: #ffffff;
                selection-background-color: #0284c7; selection-color: #ffffff;
                font-size: 15px; font-weight: bold; min-width: 140px; padding: 4px;
            }
        """.replace("OK_IKONU", ok_ikonu()))
        for ad, kod in UNITS:
            self.combo_birim.addItem(ad, kod)
        layout.addWidget(self.combo_birim)

        # 4. Fiyatlar ve Stok Satırı (fiyatların solunda para birimi seçimi)
        fiyat_row = QHBoxLayout()
        fiyat_row.setSpacing(10)

        # Satış Fiyatı
        v_sat = QVBoxLayout()
        v_sat.addWidget(QLabel("SATIŞ FİYATI:", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        sat_box = QHBoxLayout()
        sat_box.setSpacing(0)
        self.combo_sat_kur = para_birimi_kutusu()
        sat_box.addWidget(self.combo_sat_kur)
        self.spin_sat = AkilliSayiKutusu(decimals=2, max_val=1000000.0, birlesik=True)
        self.spin_sat.setFixedHeight(38)
        self.spin_sat.setMinimumWidth(105)
        sat_box.addWidget(self.spin_sat)
        v_sat.addLayout(sat_box)
        fiyat_row.addLayout(v_sat)

        # Alış Fiyatı
        v_al = QVBoxLayout()
        v_al.addWidget(QLabel("ALIŞ FİYATI:", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        al_box = QHBoxLayout()
        al_box.setSpacing(0)
        self.combo_al_kur = para_birimi_kutusu()
        al_box.addWidget(self.combo_al_kur)
        self.spin_al = AkilliSayiKutusu(decimals=2, max_val=1000000.0, birlesik=True)
        self.spin_al.setFixedHeight(38)
        self.spin_al.setMinimumWidth(105)
        al_box.addWidget(self.spin_al)
        v_al.addLayout(al_box)
        fiyat_row.addLayout(v_al)

        # Stok Miktarı
        v_stk = QVBoxLayout()
        v_stk.addWidget(QLabel("STOK MİKTARI:", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        self.spin_stk = AkilliSayiKutusu(decimals=2, max_val=10000000.0)
        self.spin_stk.setFixedHeight(38)
        self.spin_stk.setMinimumWidth(130)
        v_stk.addWidget(self.spin_stk)
        fiyat_row.addLayout(v_stk)

        layout.addLayout(fiyat_row)

        layout.addSpacing(10)

        # 5. Alt Butonlar (İptal & Kaydet/Güncelle)
        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_iptal = QPushButton("İPTAL")
        btn_iptal.setFixedHeight(40)
        btn_iptal.setFixedWidth(100)
        btn_iptal.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_iptal.setStyleSheet("background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; border-radius: 5px; font-weight: bold;")
        btn_iptal.clicked.connect(self.reject)
        btn_box.addWidget(btn_iptal)

        btn_kaydet = QPushButton("GÜNCELLE" if self.product_id else "ÜRÜNÜ KAYDET")
        btn_kaydet.setFixedHeight(40)
        btn_kaydet.setFixedWidth(150)
        btn_kaydet.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_kaydet.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover { background-color: #059669; }
        """)
        btn_kaydet.clicked.connect(self.kaydet)
        btn_box.addWidget(btn_kaydet)

        layout.addLayout(btn_box)

        # Düzenleme modundaysa mevcut verileri doldur
        if self.product_id:
            self.verileri_yukle()

    def verileri_yukle(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""
            SELECT group_id, name, stock_kg, buy_price, sell_price,
                   buy_currency, sell_currency, unit
            FROM products WHERE id = ?
        """, (self.product_id,))
        prod = c.fetchone()
        conn.close()
        if prod:
            g_id, name, stock, buy, sell, buy_kur, sell_kur, unit = prod
            combo_secimi_ayarla(self.combo_grp, g_id)
            self.txt_ad.setText(name)
            self.spin_stk.setValue(stock)
            self.spin_al.setValue(buy)
            self.spin_sat.setValue(sell)
            combo_secimi_ayarla(self.combo_al_kur, buy_kur)
            combo_secimi_ayarla(self.combo_sat_kur, sell_kur)
            combo_secimi_ayarla(self.combo_birim, unit)

    def kaydet(self):
        ad = buyuk_harf(self.txt_ad.text().strip())
        if not ad:
            QMessageBox.warning(self, "Uyarı", "Lütfen hammadde / ürün adını giriniz!")
            return

        grp_id = self.combo_grp.currentData()
        if grp_id is None:
            QMessageBox.warning(self, "Uyarı", "Önce en az bir ürün grubu eklemelisiniz!")
            return

        stk = self.spin_stk.value()
        al = self.spin_al.value()
        sat = self.spin_sat.value()
        al_kur = self.combo_al_kur.currentData()
        sat_kur = self.combo_sat_kur.currentData()
        birim = self.combo_birim.currentData()

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        if self.product_id:
            c.execute("""
                UPDATE products 
                SET group_id = ?, name = ?, stock_kg = ?, buy_price = ?, sell_price = ?,
                    buy_currency = ?, sell_currency = ?, unit = ?
                WHERE id = ?
            """, (grp_id, ad, stk, al, sat, al_kur, sat_kur, birim, self.product_id))
        else:
            c.execute("""
                INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price,
                                      buy_currency, sell_currency, unit)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (grp_id, ad, stk, al, sat, al_kur, sat_kur, birim))
        conn.commit()
        conn.close()
        self.accept()

# --- MÜŞTERİ EKLEME / GÜNCELLEME PENCERESİ ---
class MusteriEkleDialog(QDialog):
    """Hem satış hem müşteriler sayfasından çağrılan dinamik müşteri kartı."""
    def __init__(self, parent=None, customer_id=None):
        super().__init__(parent)
        self.customer_id = customer_id
        self.setWindowTitle("Müşteri Kart Bilgileri" if not customer_id else "Müşteri Güncelle")
        self.setFixedSize(600, 520)
        self.setStyleSheet("background-color: #ffffff;")
        self.init_ui()
        if self.customer_id:
            self.mevcut_yukle()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        lbl_baslik = QLabel("MÜŞTERİ KART BİLGİLERİ")
        lbl_baslik.setFixedHeight(38)
        lbl_baslik.setStyleSheet("background-color: #0284c7; color: white; font-size: 14px; font-weight: bold; padding-left: 12px; border-radius: 4px;")
        layout.addWidget(lbl_baslik)

        h_tur = QHBoxLayout()
        lbl_tur = QLabel("Müşteri Türü:")
        lbl_tur.setStyleSheet("font-size: 13px; font-weight: bold; color: #1e293b;")
        h_tur.addWidget(lbl_tur)

        self.rb_bireysel = QRadioButton("Bireysel")
        self.rb_kurumsal = QRadioButton("Kurumsal")
        self.rb_bireysel.setChecked(True)
        self.rb_bireysel.setStyleSheet("font-size: 13px; font-weight: bold; color: #334155;")
        self.rb_kurumsal.setStyleSheet("font-size: 13px; font-weight: bold; color: #334155;")
        self.btn_grp_tur = QButtonGroup(self)
        self.btn_grp_tur.addButton(self.rb_bireysel)
        self.btn_grp_tur.addButton(self.rb_kurumsal)
        self.rb_bireysel.toggled.connect(self.tur_degisti)
        h_tur.addWidget(self.rb_bireysel)
        h_tur.addWidget(self.rb_kurumsal)
        h_tur.addStretch()
        layout.addLayout(h_tur)

        self.stack_isim = QStackedLayout()

        w_bireysel = QWidget()
        l_bireysel = QHBoxLayout(w_bireysel)
        l_bireysel.setContentsMargins(0, 0, 0, 0)
        l_bireysel.setSpacing(10)

        v_ad = QVBoxLayout()
        v_ad.addWidget(QLabel("Müşteri Adı (*):", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_ad = BuyukHarfKutusu("Müşteri Adı")
        self.txt_ad.setFixedHeight(36)
        self.txt_ad.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_ad.addWidget(self.txt_ad)
        l_bireysel.addLayout(v_ad)

        v_soyad = QVBoxLayout()
        v_soyad.addWidget(QLabel("Müşteri Soyadı (*):", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_soyad = BuyukHarfKutusu("Müşteri Soyadı")
        self.txt_soyad.setFixedHeight(36)
        self.txt_soyad.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_soyad.addWidget(self.txt_soyad)
        l_bireysel.addLayout(v_soyad)
        self.stack_isim.addWidget(w_bireysel)

        w_kurumsal = QWidget()
        l_kurumsal = QVBoxLayout(w_kurumsal)
        l_kurumsal.setContentsMargins(0, 0, 0, 0)
        l_kurumsal.setSpacing(4)
        l_kurumsal.addWidget(QLabel("Müşteri Ünvanı (*):", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_unvan = BuyukHarfKutusu("Müşteri / Firma Tam Ünvanı (Örn: ADANA CAN PLASTİK SAN. TİC. LTD. ŞTİ.)")
        self.txt_unvan.setFixedHeight(36)
        self.txt_unvan.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        l_kurumsal.addWidget(self.txt_unvan)
        self.stack_isim.addWidget(w_kurumsal)
        layout.addLayout(self.stack_isim)

        h_vergi = QHBoxLayout()
        v_vno = QVBoxLayout()
        v_vno.addWidget(QLabel("Vergi No / TCKN:", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_vergi_no = QLineEdit()
        self.txt_vergi_no.setPlaceholderText("Vergi No / TCKN")
        self.txt_vergi_no.setFixedHeight(36)
        self.txt_vergi_no.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_vno.addWidget(self.txt_vergi_no)
        h_vergi.addLayout(v_vno)

        v_vd = QVBoxLayout()
        v_vd.addWidget(QLabel("Vergi Dairesi:", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_vergi_dairesi = BuyukHarfKutusu("Vergi Dairesi")
        self.txt_vergi_dairesi.setFixedHeight(36)
        self.txt_vergi_dairesi.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_vd.addWidget(self.txt_vergi_dairesi)
        h_vergi.addLayout(v_vd)
        layout.addLayout(h_vergi)

        h_iletisim = QHBoxLayout()
        v_tel = QVBoxLayout()
        v_tel.addWidget(QLabel("Telefon:", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_telefon = QLineEdit()
        self.txt_telefon.setPlaceholderText("05XX XXX XX XX")
        self.txt_telefon.setFixedHeight(36)
        self.txt_telefon.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_tel.addWidget(self.txt_telefon)
        h_iletisim.addLayout(v_tel)

        v_lim = QVBoxLayout()
        v_lim.addWidget(QLabel("Açık Hesap Limiti (TL):", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_limit = SilikSifirliBinlikKutu()
        self.txt_limit.setFixedHeight(36)
        self.txt_limit.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        v_lim.addWidget(self.txt_limit)
        h_iletisim.addLayout(v_lim)
        layout.addLayout(h_iletisim)

        layout.addWidget(QLabel("Adres:", styleSheet="font-size: 12px; font-weight: bold; color: #475569;"))
        self.txt_adres = BuyukHarfKutusu("Firma veya teslimat adresi...")
        self.txt_adres.setFixedHeight(36)
        self.txt_adres.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        layout.addWidget(self.txt_adres)

        h_btns = QHBoxLayout()
        h_btns.addStretch()
        btn_iptal = QPushButton("İptal")
        btn_iptal.setFixedHeight(38)
        btn_iptal.setFixedWidth(90)
        btn_iptal.setStyleSheet("background-color: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold;")
        btn_iptal.clicked.connect(self.reject)
        h_btns.addWidget(btn_iptal)

        self.btn_kaydet = QPushButton("  ✓ Yeni Müşteri Oluştur  " if not self.customer_id else "  ✓ Müşteriyi Güncelle  ")
        self.btn_kaydet.setFixedHeight(38)
        self.btn_kaydet.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_kaydet.setStyleSheet("background-color: #10b981; color: white; border: none; border-radius: 4px; font-weight: bold; padding: 0 16px; font-size: 13px;")
        self.btn_kaydet.clicked.connect(self.kaydet)
        h_btns.addWidget(self.btn_kaydet)
        layout.addLayout(h_btns)

    def tur_degisti(self):
        if self.rb_bireysel.isChecked():
            self.stack_isim.setCurrentIndex(0)
        else:
            self.stack_isim.setCurrentIndex(1)

    def mevcut_yukle(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(
            "SELECT name, phone, tax_office, tax_number, address, credit_limit FROM customers WHERE id = ?",
            (self.customer_id,)
        )
        row = c.fetchone()
        conn.close()
        if not row:
            return
        name, phone, tax_office, tax_number, address, credit_limit = row
        self.rb_kurumsal.setChecked(True)
        self.txt_unvan.setText(name or "")
        self.txt_telefon.setText(phone or "")
        self.txt_vergi_dairesi.setText(tax_office or "")
        self.txt_vergi_no.setText(tax_number or "")
        self.txt_adres.setText(address or "")
        if credit_limit:
            self.txt_limit.setText(f"{int(credit_limit):,}".replace(",", ".") if float(credit_limit) == int(credit_limit) else f"{credit_limit:.2f}".replace(".", ","))

    def kaydet(self):
        if self.rb_bireysel.isChecked():
            ad = buyuk_harf(self.txt_ad.text().strip())
            soyad = buyuk_harf(self.txt_soyad.text().strip())
            if not ad:
                QMessageBox.warning(self, "Eksik Bilgi", "Lütfen müşteri adını giriniz!")
                return
            tam_isim = f"{ad} {soyad}".strip()
        else:
            unvan = buyuk_harf(self.txt_unvan.text().strip())
            if not unvan:
                QMessageBox.warning(self, "Eksik Bilgi", "Lütfen kurumsal müşteri ünvanını giriniz!")
                return
            tam_isim = unvan

        v_no = self.txt_vergi_no.text().strip()
        v_dairesi = buyuk_harf(self.txt_vergi_dairesi.text().strip())
        tel = self.txt_telefon.text().strip()
        adres = buyuk_harf(self.txt_adres.text().strip())
        limit_val = self.txt_limit.sayisal_deger() if hasattr(self.txt_limit, "sayisal_deger") else 0.0

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        if self.customer_id:
            c.execute("""
                UPDATE customers
                SET name = ?, phone = ?, tax_office = ?, tax_number = ?, address = ?, credit_limit = ?
                WHERE id = ?
            """, (tam_isim, tel, v_dairesi, v_no, adres, limit_val, self.customer_id))
        else:
            c.execute("""
                INSERT INTO customers (name, phone, tax_office, tax_number, address, credit_limit, remaining_debt)
                VALUES (?, ?, ?, ?, ?, ?, 0.0)
            """, (tam_isim, tel, v_dairesi, v_no, adres, limit_val))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Başarılı", f"'{tam_isim}' başarıyla kaydedildi!")
        self.accept()


class MusteriSecDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.secilen_musteri = None
        self.setWindowTitle("Müşteri Seç")
        self.setFixedSize(650, 480)
        self.setStyleSheet("background-color: #ffffff;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(12)

        lbl_baslik = QLabel("MÜŞTERİ SEÇ")
        lbl_baslik.setFixedHeight(40)
        lbl_baslik.setStyleSheet("background-color: #0284c7; color: white; font-size: 15px; font-weight: bold; padding-left: 12px; border-radius: 4px;")
        layout.addWidget(lbl_baslik)

        self.txt_search = BuyukHarfKutusu("Müşteri İsmini Giriniz...")
        self.txt_search.setFixedHeight(40)
        self.txt_search.setStyleSheet("border: 2px solid #cbd5e1; border-radius: 6px; padding: 0 12px; font-size: 14px; font-weight: bold;")
        self.txt_search.textChanged.connect(self.musterileri_ara)
        layout.addWidget(self.txt_search)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(4)
        self.tbl.setHorizontalHeaderLabels(["Müşteri", "Açık Hesap Limiti", "Kalan Borç", "Seç"])
        self.tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.tbl.setColumnWidth(1, 130)
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.tbl.setColumnWidth(2, 130)
        self.tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.tbl.setColumnWidth(3, 80)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl.setStyleSheet("""
            QTableWidget { border: 1px solid #e2e8f0; font-size: 13px; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 36px; border: none; border-bottom: 2px solid #cbd5e1; }
        """)
        layout.addWidget(self.tbl)

        b_row = QHBoxLayout()
        btn_yeni = QPushButton("  + Yeni Müşteri Oluştur  ")
        btn_yeni.setFixedHeight(38)
        btn_yeni.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yeni.setStyleSheet("background-color: #0284c7; color: white; font-size: 13px; font-weight: bold; border-radius: 5px; border: none;")
        btn_yeni.clicked.connect(self.yeni_musteri_ac)
        b_row.addWidget(btn_yeni)

        b_row.addStretch()

        btn_kapat = QPushButton("Kapat")
        btn_kapat.setFixedHeight(38)
        btn_kapat.setFixedWidth(90)
        btn_kapat.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_kapat.setStyleSheet("background-color: #0ea5e9; color: white; font-weight: bold; border-radius: 5px; border: none;")
        btn_kapat.clicked.connect(self.reject)
        b_row.addWidget(btn_kapat)

        layout.addLayout(b_row)
        self.musterileri_ara()

    def musterileri_ara(self):
        txt = buyuk_harf(self.txt_search.text().strip())
        self.tbl.setRowCount(0)
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        if txt:
            c.execute("SELECT id, name, credit_limit, remaining_debt FROM customers WHERE name LIKE ? ORDER BY name ASC LIMIT 50", (f"%{txt}%",))
        else:
            c.execute("SELECT id, name, credit_limit, remaining_debt FROM customers ORDER BY name ASC LIMIT 50")
        rows = c.fetchall()
        conn.close()

        self.tbl.setRowCount(len(rows))
        for r_idx, (cid, name, limit_val, rem_debt) in enumerate(rows):
            self.tbl.setRowHeight(r_idx, 38)
            it_name = QTableWidgetItem(name)
            it_name.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.tbl.setItem(r_idx, 0, it_name)

            it_lim = QTableWidgetItem(f"{(limit_val or 0.0):,.2f}")
            it_lim.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl.setItem(r_idx, 1, it_lim)

            it_rem = QTableWidgetItem(f"{(rem_debt or 0.0):,.2f}")
            it_rem.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl.setItem(r_idx, 2, it_rem)

            btn_sec = QPushButton("Seç")
            btn_sec.setFixedHeight(28)
            btn_sec.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_sec.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 3px; border: none;")
            btn_sec.clicked.connect(lambda _, c_id=cid, c_nm=name: self.secim_yap(c_id, c_nm))
            self.tbl.setCellWidget(r_idx, 3, btn_sec)

    def secim_yap(self, cid, name):
        self.secilen_musteri = {"id": cid, "name": name}
        self.accept()

    def yeni_musteri_ac(self):
        dlg = MusteriEkleDialog(self)
        if dlg.exec():
            self.musterileri_ara()


class MusteriKumeleyici:
    # 1. Müşteri Adından Kesinlikle Kazınacak Kelimeler (Malzeme, Katkı, İşlem Notları)
    YASAKLI_KELIMELER = {
        "numune", "numunesi", "kalsit", "kalsitli", "elyaf", "elyafli", "kirim", "kırım", "kırma", "kirma",
        "diger", "diğer", "musteri", "müşteri", "musterisi", "musteriler", "satis", "satislar", "satış",
        "hesaba", "hesap", "yazilacak", "yazildi", "yaz", "alindi", "verildi", "gitcek", "gidecek", "iade",
        "kazanc", "kazanç", "net", "toplam", "tl", "kg", "kğ", "ton", "bin", "netkar",
        "siyah", "beyaz", "gri", "krem", "mavi", "kirmizi", "kırmızı", "yesil", "yeşil", "sari", "sarı", "mor",
        "ntr", "nat", "naturel", "renkli", "seffaf", "şeffaf",
        "diyah", "siyak", "sşyah", "kermizi", "kermızı", "krimizi", "beyz", "beya",
        "granul", "granül", "granur", "granür", "grangr", "hurda", "orj", "orjinal", "orijinal",
        "capak", "çapak", "aliyo", "aliyor", "aldi", "verecek", "sisirme", "şişirme", "enjeksiyon",
        "cekme", "po", "pu", "mal", "hammadde", "fiyati", "fiyatı", "tonu",
        "ay", "ayi", "ayinda", "ayında", "haftasi", "haftası",
        "odeme", "ödeme", "odemesi", "ödemesi", "odemesini", "ödemesini", "ödemem", "odemem",
        "odemeler", "ödemeler", "alindi", "alındı", "aldim", "aldım", "verdi", "verdim",
        "geldi", "getirdi", "gonderdi", "gönderdi", "kalan", "kaldi", "kaldı", "bakiye", "bakiyesi",
        "den", "dan", "ten", "tan", "icin", "için", "tuttu", "yapildi", "yapıldı",
        "kalem", "adet", "kutu", "cuval", "çuval", "palet", "torba", "nakit",
        "takoz", "ince", "kalin", "kalın", "vakum", "buzbeyaz", "buz", "dolar", "euro", "birim", "kurus", "kuruş",
        "aksesuar", "kalemcilik", "etiket", "ari", "arı", "bey", "abi", "abla", "hanim", "hanım",
        "usta", "ortak", "ortağı", "kardes", "kardeş", "arti", "artı",
        "brabs", "beyap", "otj", "orji", "mobil", "len", "meh", "kdv", "dahil",
        "elotromel", "elastomer", "asa", "alti", "altı", "alis", "alış", "kamuran", "com", "elf"
    }
    for _ay in (
        "ocak", "subat", "şubat", "mart", "nisan", "mayis", "mayıs", "haziran",
        "temmuz", "agustos", "ağustos", "eylul", "eylül", "ekim", "kasim", "kasım", "aralik", "aralık",
    ):
        YASAKLI_KELIMELER.add(_ay)
        YASAKLI_KELIMELER.add(_ay + "in")
        YASAKLI_KELIMELER.add(_ay + "ın")
        YASAKLI_KELIMELER.add(_ay + "nin")
        YASAKLI_KELIMELER.add(_ay + "nın")
    # Kurumsal Ekler (Soyadı olamaz)
    KURUMSAL_EKLER = {
        "fermuar", "aksesuar", "kalip", "kalıp", "plastik", "metal",
        "tekstil", "ambalaj", "etiket", "kaucuk", "iplik", "sanayi",
        "ticaret", "ltd", "sti", "aş", "düğme", "dugme", "kardesler", "ogullari"
    }
    # Türkiye İlleri, İlçeleri ve Marmara Plastik Sanayi Bölgeleri Sözlüğü
    SANAYI_VE_SEHIR_BOLGELERI = {
        "ikitelli", "topcular", "topçular", "merter", "bayrampasa", "bayrampaşa",
        "hadimkoy", "hadımköy", "kirac", "kıraç", "beylikduzu", "beylikdüzü",
        "esenyurt", "zeytinburnu", "gungoren", "güngören", "dudullu", "imes",
        "tuzla", "umraniye", "ümraniye", "kartal", "pendik", "silivri", "catalca",
        "gebze", "dilovasi", "dilovası", "kocaeli", "bursa", "corlu", "çorlu",
        "cerkezkoy", "çerkezköy", "tekirdag", "tekirdağ", "izmir", "ankara",
        "konya", "gaziantep", "adana", "kayseri", "denizli", "eskisehir", "eskişehir",
        "levent"
    }

    @classmethod
    def tr_temizle_ve_koke_in(cls, kelime):
        """
        Türkçe klavye hatalarını (hürbey / hurbey) ve çoğul eklerini (-ler / -lar)
        ortak bir anahtara indirger.
        """
        if not kelime:
            return ""
        s = kelime.lower()
        tr_map = {'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c'}
        for tr, en in tr_map.items():
            s = s.replace(tr, en)

        if s.endswith("ler") or s.endswith("lar"):
            s = s[:-3]
        return s

    @classmethod
    def konumu_basa_al_ve_formatla(cls, isim_ham):
        """
        'Hürbey İkitelli' -> 'İKİTELLİ HÜRBEY'
        'Topçular Ahmet'  -> 'TOPÇULAR AHMET'
        yapar. Bölgeyi daima en başa sabitler.
        """
        if not isim_ham or len(isim_ham) < 3:
            return isim_ham
        kelimeler = isim_ham.split()
        if len(kelimeler) < 2:
            return isim_ham.upper()
        bulunan_bolgeler = []
        diger_kelimeler = []
        tr_map = {'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c'}

        def bolge_norm(kelime):
            s = (kelime or "").lower()
            for tr, en in tr_map.items():
                s = s.replace(tr, en)
            return s

        bolge_set = {bolge_norm(x) for x in cls.SANAYI_VE_SEHIR_BOLGELERI}
        for w in kelimeler:
            if bolge_norm(w) in bolge_set:
                bulunan_bolgeler.append(w.upper())
            else:
                diger_kelimeler.append(w.upper())
        if bulunan_bolgeler and diger_kelimeler:
            return " ".join(bulunan_bolgeler + diger_kelimeler)
        return isim_ham.upper()

    @classmethod
    def musteri_kok_olustur(cls, ham_isim):
        """
        'Numune İsmail', 'Kalsit Merter', 'Hürbeyler' metinlerinden
        tertemiz birleştirme anahtarı üretir.
        """
        if not ham_isim:
            return ""
        ham_isim = cls.konumu_basa_al_ve_formatla(ham_isim)

        ham_norm = ham_isim.lower()
        tr_map = {'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c'}
        for tr, en in tr_map.items():
            ham_norm = ham_norm.replace(tr, en)
        parcalar = re.findall(r'[a-z]+', ham_norm)

        anlamli = []
        for p in parcalar:
            p_kok = cls.tr_temizle_ve_koke_in(p)
            if p_kok not in cls.YASAKLI_KELIMELER and p not in cls.KURUMSAL_EKLER and len(p_kok) > 1:
                anlamli.append(p_kok)
        if not anlamli:
            anlamli = [cls.tr_temizle_ve_koke_in(p) for p in parcalar if p not in cls.YASAKLI_KELIMELER]
        return "_".join(anlamli)

    @classmethod
    def kok_benzerlik_skoru(cls, kok1, kok2):
        if not kok1 or not kok2:
            return 0.0
        if kok1 == kok2:
            return 1.0
        kelimeler1 = [k for k in kok1.split("_") if len(k) >= 4]
        kelimeler2 = [k for k in kok2.split("_") if len(k) >= 4]
        en_iyi = 0.0
        for k1 in kelimeler1:
            for k2 in kelimeler2:
                oran = difflib.SequenceMatcher(None, k1, k2).ratio()
                if oran > en_iyi:
                    en_iyi = oran
        genel = difflib.SequenceMatcher(None, kok1, kok2).ratio()
        return max(en_iyi, genel)

    @classmethod
    def kanonik_musteri_anahtari(cls, isim_ham):
        """
        'ALAATTİN MUSTAF', 'MUSTAFA ALAATTİ', 'ALATTİN MUSTAFA' metinlerinin
        hepsini tek bir ortak anahtara ('alaattin_mustafa') indirger.
        Kelime sırası (Mustafa Alaattin vs Alaattin Mustafa) farkını yok eder.
        """
        if not isim_ham or len(isim_ham.strip()) < 2:
            return ""
        s = isim_ham.lower()
        tr_map = {'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c'}
        for tr, en in tr_map.items():
            s = s.replace(tr, en)
        ham_kelimeler = re.findall(r'[a-z]+', s)

        anlamli_kelimeler = []
        for w in ham_kelimeler:
            w_kok = cls.tr_temizle_ve_koke_in(w)
            if (w_kok in cls.YASAKLI_KELIMELER or
                w in getattr(WhatsAppSatisAyristirici, 'HAMMADDE_TICARI_ISIMLER', set()) or
                len(w) <= 1):
                continue

            if w in ["mustaf", "mustafe"]: w = "mustafa"
            elif w in ["alaatti", "alattin", "alatin", "alatttin"]: w = "alaattin"
            elif w in ["ahme", "ahmet"]: w = "ahmet"
            elif w in ["ismai", "ismail"]: w = "ismail"
            elif w in ["hurbey", "hurbeyler"]: w = "hurbey"
            anlamli_kelimeler.append(w)
        if not anlamli_kelimeler:
            return "musterisiz"
        anlamli_kelimeler.sort()
        return "_".join(anlamli_kelimeler)


FIIL_EKLERI = ("MIS", "MIŞ", "MUS", "MÜŞ", "DIK", "TIK", "CEK", "CAK", "DIM", "DIMI", "YOR", "YORUM")
YASAKLI_TAM_KELIMELER = {
    "ADAM", "ADAMA", "ADAMIN", "HESAP", "HESABA", "YAZ", "YAZAR", "YAZARSIN", "YAZILDI",
    "CIKAR", "CIKARIM", "CIKACAK", "KULLAN", "KULLANMIS", "PATLAK", "PATLAMIS", "CUVAL",
    "CUVALLAR", "GELDI", "GITTI", "ALINDI", "VERILDI", "VER", "AL", "BEN", "SEN", "BIZ"
}


def gecersiz_cari_mi(metin):
    """Metin fiil, emir kipi veya sohbet cümlesi içeriyorsa True döner."""
    temiz = str(metin or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
    kelimeler = temiz.split()
    if not kelimeler or len(temiz) < 3:
        return True
    if any(k in YASAKLI_TAM_KELIMELER for k in kelimeler):
        return True
    isim_istisna = {
        "PLASTIK", "METAL", "KALIP", "SANAYI", "TICARET", "LIMITED", "AKSESUAR",
        "TAHSIN", "ERSIN", "RAMAZAN", "MUSTAFA", "ALAATTIN",
    }
    for k in kelimeler:
        if k in isim_istisna or k.endswith("PLASTIK"):
            continue
        if len(k) >= 5 and any(k.endswith(ek) for ek in FIIL_EKLERI):
            return True
    return False


class WhatsAppSatisAyristirici:
    POLIMER_DONUSUM = {
        "pom": "POM", "delrin": "POM", "derlin": "POM", "asetal": "POM",
        "abs": "ABS",
        "ant": "ANTİŞOK", "anti": "ANTİŞOK", "antisok": "ANTİŞOK", "hips": "ANTİŞOK",
        "pa6": "PA6", "n6": "PA6", "naylon6": "PA6", "ultramid": "PA6", "akromid": "PA6",
        "pa66": "PA66", "n66": "PA66", "naylon66": "PA66",
        "pp": "PP MOBLEN", "moblen": "PP MOBLEN",
        "hdpe": "HDPE", "elteks": "HDPE", "i20": "HDPE", "pe100": "HDPE",
        "ldpe": "LDPE", "naylon": "LDPE",
        "pvc": "PVC",
        "ps": "KRİSTAL", "kristal": "KRİSTAL", "gpps": "KRİSTAL",
        "pc": "PC", "polikarbon": "PC", "lexan": "PC",
        "pmma": "PMMA", "pleksi": "PMMA", "akrilik": "PMMA",
        "pet": "PET", "polyester": "PET", "petg": "PET",
        "asa": "ASA", "pbt": "PBT", "tpu": "TPU", "tpe": "TPE", "eva": "EVA"
    }
    HAMMADDE_TICARI_ISIMLER = {
        "delrin", "derlin", "pom", "asetal", "polyacetal", "kocetal", "kepital",
        "elteks", "eltek", "i20", "pe100", "hdpe", "ldpe", "lldpe", "pe",
        "moblen", "pp", "polipropilen", "homopolimer", "kopolimer", "random",
        "abs", "akrilik", "pleksi", "pmma", "san", "kristal", "krsital", "krsistal", "gpps", "hips", "antisok", "antişok", "ant",
        "pa6", "pa66", "naylon", "naylon6", "akromid", "ultramid",
        "pvc", "pet", "polikarbon", "pc", "tpe", "tpu", "eva", "kaucuk", "kauçuk", "valoks", "valox", "pbt"
    }
    TR_ISIM_DUZELTME = {
        "GOKHAN": "GÖKHAN", "BULENT": "BÜLENT", "OZGUR": "ÖZGÜR", "SUKRU": "ŞÜKRÜ",
        "HUSEYIN": "HÜSEYİN", "ISMAIL": "İSMAİL", "IBRAHIM": "İBRAHİM", "OMER": "ÖMER",
        "CAGLAR": "ÇAĞLAR", "UGUR": "UĞUR", "YASAR": "YAŞAR", "ALATTIN": "ALAATTİN",
        "HURBEY": "HÜRBEY", "HURBEYLER": "HÜRBEYLER", "SADI": "ŞADİ", "COSAR": "COŞAR"
    }
    YAYGIN_TURKCE_ISIMLER = [
        "MEHMET","AHMET","MUSTAFA","ALİ","HÜSEYİN","HASAN","İBRAHİM","MURAT","ÖMER",
        "OSMAN","SÜLEYMAN","İSMAİL","MAHMUT","KEMAL","CENGİZ","ERDAL","ERKAN","SERKAN",
        "SERDAR","VOLKAN","EMRE","ONUR","OKAN","OĞUZ","ORHAN","ÖZGÜR","ÖZKAN","BARIŞ",
        "BURAK","DENİZ","DOĞAN","EREN","FATİH","GÖKHAN","HAKAN","İLKER","KAAN","LEVENT",
        "MERT","METİN","RECEP","SAMİ","SEDAT","SELÇUK","UĞUR","ÜMİT","YAŞAR","YAVUZ",
        "YILMAZ","YUNUS","ZAFER","ADEM","ADNAN","ALPER","ARDA","AYDIN","BÜLENT","CEMAL",
        "COŞKUN","ÇAĞLAR","EYÜP","FARUK","GÜVEN","KENAN","KEREM","SALİH","SİNAN","SONER",
        "TANER","YAKUP","ZEKİ","AYŞE","FATMA","EMİNE","HATİCE","ZEYNEP","ELİF","MERYEM",
        "SEVGİ","SULTAN","GÜL","NUR","NURCAN","NAZLI","ÖZLEM","PINAR","SEMA","SEVDA",
        "SİBEL","TUBA","YASEMİN","YILDIZ","ZEHRA","ASLI","AYLA","BEYZA","BURCU","CANAN",
        "DERYA","DİLEK","DUYGU","EBRU","ECE","EDA","ESRA","FİLİZ","FUNDA","GAMZE","LEYLA",
        "MELEK","MERVE","NİHAL","OYA","ÖZGE","SELMA","SENA","SİNEM"
    ]
    @classmethod
    def en_yakin_turkce_ismi_tahmin_et(cls, kisaltma):
        kisaltma_norm = turkce_toleransli_metin(kisaltma)
        if not (3 <= len(kisaltma_norm) <= 5):
            return None
        en_iyi_ad, en_iyi_skor = None, 0.0
        for aday in cls.YAYGIN_TURKCE_ISIMLER:
            aday_norm = turkce_toleransli_metin(aday)
            if aday_norm == kisaltma_norm:
                return aday
            if len(aday_norm) <= len(kisaltma_norm) or aday_norm[0] != kisaltma_norm[0]:
                continue
            konum, hepsi_eslesti = 0, True
            for harf in kisaltma_norm:
                bulundu = aday_norm.find(harf, konum)
                if bulundu == -1:
                    hepsi_eslesti = False
                    break
                konum = bulundu + 1
            if not hepsi_eslesti:
                continue
            oran = len(kisaltma_norm) / len(aday_norm)
            if oran < 0.5:
                continue
            benzerlik = difflib.SequenceMatcher(None, kisaltma_norm, aday_norm).ratio()
            skor = oran * 0.6 + benzerlik * 0.4
            if skor > en_iyi_skor:
                en_iyi_skor, en_iyi_ad = skor, aday
        return en_iyi_ad if en_iyi_skor >= 0.55 else None
    MEDYA_VE_COP_DESENLER = [
        "<medya dahil edilmedi>", "medya dahil edilmedi", "goruntu dahil edilmedi",
        "image omitted", "audio omitted", "video omitted", "sticker omitted",
        "cevapsiz sesli arama", "cevapsiz goruntulu arama", "mesaj silindi",
        "bu mesaj silindi", "guvenlik kodu degisti", "gruba katildi", "gruptan ayrildi",
        "eklendi", "cikarildi", "degistirildi", ".jpg", ".png", ".webp", ".mp4", ".opus", ".pdf"
    ]

    RE_KILO = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:kg|kğ|kilo|ton)', re.IGNORECASE)
    RE_FIYAT = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:tl|lira|₺)', re.IGNORECASE)
    RE_ODEME = re.compile(r'(\d+(?:[.,]\d+)?)\s*(?:bin)?\s*(?:₺|tl|usd|\$)?\s*(?:odeme|pesin|nakit)', re.IGNORECASE)
    RE_KAZANC = re.compile(r'net\s*kazan[cç]\s*(\d+(?:[.,]\d+)?)\s*(?:bin)?', re.IGNORECASE)
    # iOS/Mac: [8/17/20, 1:47:32 PM] Bro Ahmet: ...
    # iOS/Mac (saat önce): [11:40 AM, 10.09.2026] Bro Ahmet: ...
    # Android: 17/08/2020, 13:47 - Bro Ahmet: ...
    RE_TAM_BASLIK = re.compile(
        r'^[\u200e\ufeff\s]*'
        r'(?:'
        r'\[(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}),?\s+(\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?)\]'
        r'|'
        r'\[(\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?),?\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\]'
        r'|'
        r'(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}),?\s+(\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?)\s*-\s*'
        r')\s*'
        r'([^:]+):\s*(.*)$',
        re.DOTALL
    )

    @classmethod
    def baslik_ve_mesaj_coz(cls, satir):
        """Satırdan Tarih, Saat, Personel ve Mesaj Gövdesini hatasız çeker."""
        m = cls.RE_TAM_BASLIK.match(satir.strip())
        if not m:
            return None, None, "DÜKKAN", satir.strip()
        ham_tarih = m.group(1) or m.group(4) or m.group(5)
        saat = m.group(2) or m.group(3) or m.group(6)
        gonderen_ham = m.group(7).strip()
        mesaj_govdesi = (m.group(8) or "").strip()
        tarih_parcalar = re.split(r'[./-]', ham_tarih)
        if len(tarih_parcalar) == 3:
            p1, p2, p3 = tarih_parcalar
            if len(p3) == 2:
                p3 = "20" + p3
            if int(p1) <= 12 and int(p2) > 12:
                gun, ay, yil = p2.zfill(2), p1.zfill(2), p3
            else:
                gun, ay, yil = p1.zfill(2), p2.zfill(2), p3
            duzgun_tarih = f"{gun}.{ay}.{yil}"
        else:
            duzgun_tarih = ham_tarih
        g_up = gonderen_ham.upper()
        if "AHMET" in g_up:
            personel = "AHMET"
        elif "BABA" in g_up:
            personel = "BABA"
        elif "SİZ" in g_up or "SIZ" in g_up:
            personel = "MEHMET"
        elif "MUSTAFA" in g_up:
            personel = "MUSTAFA"
        elif "BURAK" in g_up:
            personel = "BURAK"
        elif "HARUN" in g_up:
            personel = "HARUN"
        else:
            personel = re.sub(r'[^a-zA-ZğüşıöçĞÜŞİÖÇ\s]', '', gonderen_ham).strip().upper()
            if not personel or len(personel) < 2:
                personel = "DÜKKAN"
        return duzgun_tarih, saat, personel, mesaj_govdesi

    @classmethod
    def cop_satir_mi(cls, satir_norm):
        """Görsel, medya, ses kaydı veya WhatsApp sistem bildirimi mi kontrol eder"""
        if not satir_norm or len(satir_norm) < 3:
            return True
        for cop in cls.MEDYA_VE_COP_DESENLER:
            if cop in satir_norm:
                return True
        return False

    @classmethod
    def sohbet_ve_rapor_filtresi(cls, satir):
        """Sadece kilo içermeyen ciro/rapor cümlelerini eler; tahsilatlı satışları atmaz."""
        if not satir:
            return True
        s = satir.lower().replace("ı", "i").replace("ş", "s").replace("ğ", "g").replace("ç", "c").replace("ö", "o").replace("ü", "u")
        if cls.RE_KILO.search(s):
            return False
        if re.search(r'\bnet\s+satis\b', s) or "ilk haftasi" in s or "ciro" in s:
            return True
        if re.search(r'\b\d+(?:[.,]\d+)?\s*(?:bin)?\s*(?:tl|lira)\s+net\s+satis\b', s):
            return True
        return False

    @classmethod
    def tl_tutari_oku(cls, metin):
        """12.500 TL → 12500; 12,50 → 12.5. Bin+büyük sayı çelişkisini ayırır."""
        s = cls.temizle_norm(metin)
        bin_m = re.search(r'(\d+(?:[.,]\d+)?)\s*bin\b', s)
        if bin_m:
            val = float(bin_m.group(1).replace(",", "."))
            if val >= 100:
                return None, "Tutar çelişkisi: sayı ve 'bin' birlikte."
            return val * 1000.0, ""
        m = re.search(r'(\d{1,3}(?:\.\d{3})+)(?:[.,](\d{1,2}))?\s*(?:tl|lira|₺)', s)
        if m:
            tam = float(m.group(1).replace(".", ""))
            if m.group(2):
                tam += float("0." + m.group(2))
            return tam, ""
        m = re.search(r'(\d+)(?:[.,](\d{1,2}))?\s*(?:tl|lira|₺)', s)
        if m:
            tam = float(m.group(1))
            if m.group(2):
                tam += float("0." + m.group(2))
            return tam, ""
        return 0.0, ""

    @classmethod
    def kg_coz(cls, metin):
        """'iki ton 800 kg' → 2800; düz kg/ton kalır."""
        s = cls.temizle_norm(metin)
        yazi = {
            "bir": 1, "iki": 2, "uc": 3, "dort": 4, "bes": 5,
            "alti": 6, "yedi": 7, "sekiz": 8, "dokuz": 9, "on": 10,
        }
        m = re.search(
            r'(?:(' + "|".join(yazi.keys()) + r')|(\d+(?:[.,]\d+)?))\s*ton'
            r'(?:\s+(\d+(?:[.,]\d+)?)\s*(?:kg|kğ|kilo))?',
            s,
        )
        if m:
            if m.group(1):
                ton = float(yazi[m.group(1)])
            else:
                ton = float((m.group(2) or "0").replace(",", "."))
            extra = float((m.group(3) or "0").replace(",", "."))
            return ton * 1000.0 + extra
        kg_m = cls.RE_KILO.search(s)
        if not kg_m:
            return 0.0
        v = float(kg_m.group(1).replace(",", "."))
        if "ton" in kg_m.group(0).lower() and v <= 50:
            v *= 1000.0
        return v

    @staticmethod
    def temizle_norm(metin):
        return turkce_toleransli_metin(metin)

    @staticmethod
    def tur_tespit_et(isim):
        norm = WhatsAppSatisAyristirici.temizle_norm(isim)
        kurumsal_anahtarlar = [
            "plastik", "kalip", "metal", "fermuar", "sanayi", "ltd", "sti",
            "as", "kalemcilik", "geri donusum", "kaucuk", "ambalaj", "tekstil", "oto"
        ]
        if any(k in norm for k in kurumsal_anahtarlar):
            return "KURUMSAL"
        return "BİREYSEL"

    @classmethod
    def fiyat_ile_form_tahmin_et(cls, polimer, renk, girilen_fiyat, mevcut_form, tum_urunler):
        """
        Mesajda 'granül' veya 'çapak' açıkça yazmıyorsa, fiyata bakarak
        malzemenin granül mü çapak mı olduğunu çözer.
        (Örn: Siyah POM 27 TL ise çapak, 45 TL yazılmışsa granül kabul eder)
        """
        # Eğer mesajda zaten 'GRANÜL', 'ÇAPAK' veya 'ORJİNAL' açıkça yazılmışsa dokunma
        if mevcut_form:
            return mevcut_form
        if not girilen_fiyat or girilen_fiyat <= 0:
            return ""
        # Veritabanında aynı polimer ve renkteki kayıtlı baz fiyatları tara
        kayitli_fiyatlar = []
        pol_norm = cls.temizle_norm(polimer)
        renk_norm = cls.temizle_norm(renk) if renk else ""
        for p_id, p_ad, p_fiyat in tum_urunler:
            p_ad_norm = cls.temizle_norm(p_ad)
            if pol_norm in p_ad_norm and (not renk_norm or renk_norm in p_ad_norm):
                kayitli_fiyatlar.append(float(p_fiyat or 0.0))
        # Eğer referans ürün varsa onun fiyatını baz al, yoksa genel sektör eşiklerini kullan
        baz_fiyat = min(kayitli_fiyatlar) if kayitli_fiyatlar else 0.0
        # POLİMER BAZLI PİYASA EŞİKLERİ (Çapak vs Granül Ayrım Noktaları)
        ESIKLER = {
            "POM": {"capak_max": 34.0, "granul_min": 38.0},
            "PP": {"capak_max": 28.0, "granul_min": 34.0},
            "HDPE": {"capak_max": 30.0, "granul_min": 36.0},
            "ABS": {"capak_max": 42.0, "granul_min": 52.0},
            "PA6": {"capak_max": 55.0, "granul_min": 70.0},
            "PA66": {"capak_max": 70.0, "granul_min": 90.0},
        }
        esik = ESIKLER.get(polimer, {"capak_max": 30.0, "granul_min": 40.0})
        # Eğer kayıtlı baz fiyat varsa ve girilen fiyat baz fiyatın %35+ üzerindeyse -> Granül
        if baz_fiyat > 0:
            if girilen_fiyat >= baz_fiyat * 1.35 or girilen_fiyat >= esik["granul_min"]:
                return "GRANÜL"
            elif girilen_fiyat <= esik["capak_max"]:
                return "ÇAPAK"
        else:
            if girilen_fiyat >= esik["granul_min"]:
                return "GRANÜL"
            elif girilen_fiyat <= esik["capak_max"]:
                return "ÇAPAK"
        return ""

    @classmethod
    def musteri_adini_ayikla(cls, satir_ham, tum_musteriler):
        """
        'Valoks Baran', 'PC Gülsan Pompa', 'Diyah Alaattin', 'Gitcek Şeffaf PC Seçim' gibi
        satırlardan çöpleri söküp atar; geriye sadece saf müşteri adını bırakır.
        """
        norm_s = cls.temizle_norm(satir_ham)

        # 1. Saat, AM/PM ve Personel Kalıntılarını Sil
        norm_s = re.sub(r'\b(am|pm)\b', ' ', norm_s)
        norm_s = re.sub(r'\[?\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?\]?', ' ', norm_s)

        cop_personel = ["bro ahmet", "baba1", "baba", "siz", "admin", "kole"]
        for p in cop_personel:
            norm_s = re.sub(r'\b' + re.escape(p) + r'\b', ' ', norm_s)

        # 2. HAMMADDE VE TİCARİ İSİMLERİ İSİMDEN SÖKÜP AT (Valoks, PC, Akrilik vb.)
        for hammadde in cls.HAMMADDE_TICARI_ISIMLER:
            norm_s = re.sub(r'\b' + re.escape(hammadde) + r'\b', ' ', norm_s)

        # 3. YASAKLI ÇÖP KELİMELERİ VE RENK HATALARINI TEMİZLE (diyah, kermızı, gitcek, iade)
        isim_korunan = {"usta", "bey", "abi", "abla", "hanim", "hanım", "ahmet", "yavuz"}
        for yasak in MusteriKumeleyici.YASAKLI_KELIMELER:
            if yasak in isim_korunan:
                continue
            norm_s = re.sub(r'\b' + re.escape(yasak) + r'\b', ' ', norm_s)

        # 4. Sayı, Para, Kilo
        norm_s = re.sub(r'\d+(?:[.,]\d+)?\s*(?:kg|kğ|kilo|ton|tl|lira|₺|\$|usd|bin)', ' ', norm_s)
        norm_s = re.sub(r'[\d.,₺$:]', ' ', norm_s)

        # 5. Kalan Harfleri Parçala
        kelimeler = norm_s.split()
        temiz_kelimeler = []
        for w in kelimeler:
            w_harf = re.sub(r'[^a-zA-ZğüşıöçĞÜŞİÖÇ]', '', w)
            w_harf_buyuk = w_harf.upper()
            if w_harf_buyuk in ISIM_DUZELTME_HARITASI:
                w_harf_buyuk = ISIM_DUZELTME_HARITASI[w_harf_buyuk]
            if w_harf_buyuk in ERKEK_ISIMLERI or w_harf_buyuk in SOYISIMLER:
                temiz_kelimeler.append(w_harf_buyuk)
                continue
            if len(w_harf) > 1 and w_harf not in cls.HAMMADDE_TICARI_ISIMLER and (
                w_harf not in MusteriKumeleyici.YASAKLI_KELIMELER or w_harf in isim_korunan
            ):
                temiz_kelimeler.append(w_harf.upper())

        if not temiz_kelimeler:
            return None, "MÜŞTERİSİZ SATIŞ"

        saf_isim = " ".join(temiz_kelimeler).strip()

        if "MESAJ" in saf_isim or "GORUNTU" in saf_isim or "GÖRÜNTÜ" in saf_isim or "DAHIL" in saf_isim or "DAHİL" in saf_isim:
            return None, "MÜŞTERİSİZ SATIŞ"

        if len(saf_isim.split()) > 6:
            return None, "MÜŞTERİSİZ SATIŞ"

        if hasattr(MusteriKumeleyici, 'konumu_basa_al_ve_formatla'):
            saf_isim = MusteriKumeleyici.konumu_basa_al_ve_formatla(saf_isim)

        if saf_isim in ["ADAM", "ÇOCUK", "ABİ", "YOK", "BEN", "DEN", "DAN", "VERDİM"] or len(saf_isim) < 3:
            return None, "MÜŞTERİSİZ SATIŞ"

        tahsilat_kelimeleri = ["ODEME", "ÖDEME", "BAKİYE", "NAKİT", "KDV", "HAVALE", "İADE", "ÇEK", "DELDİM"]
        for t_kelime in tahsilat_kelimeleri:
            if t_kelime in saf_isim:
                return None, "MÜŞTERİSİZ SATIŞ"

        # --- SÜPER AGRESİF KÖK YAKALAYICI ---
        isim_ust = saf_isim.upper().replace("İ", "I")

        # MUTASAN KURALI (İçinde mutasan geçen her şeyi tek isme sabitler)
        if "MUTASAN" in isim_ust:
            saf_isim = "MUTASAN"

        # Orhan Bahçe Kuralı (atıcakmış vb. ekleri yutar)
        elif "ORHAN" in isim_ust and "BAH" in isim_ust:
            saf_isim = "ORHAN BAHÇE"

        # Ramazan Kuralları
        elif "RAMAZAN" in isim_ust:
            if "EFOR" in isim_ust:
                saf_isim = "EFOR PLASTİK RAMAZAN"
            elif "ELEKT" in isim_ust:
                saf_isim = "RAMAZAN ELEKTRİKÇİ"
            elif "IKITELLI" in isim_ust:
                saf_isim = "İKİTELLİ RAMAZAN"
            else:
                saf_isim = "RAMAZAN"

        elif "ALAT" in isim_ust or "ALAAT" in isim_ust or "ALLAT" in isim_ust:
            saf_isim = "ALAATTİN"
        elif "BAYP" in isim_ust or "RAYP" in isim_ust:
            saf_isim = "BAYPEN"
        elif "CEYH" in isim_ust:
            saf_isim = "CEYHAN"

        if gecersiz_cari_mi(saf_isim):
            return None, "MÜŞTERİSİZ SATIŞ"

        try:
            esleyici = MusteriEsleyici(tum_musteriler or [])
            es = esleyici.coz(saf_isim)
            if es.kesin:
                return es.musteri_id, es.musteri_ad
        except Exception:
            pass

        if hasattr(MusteriKumeleyici, 'kanonik_musteri_anahtari'):
            saf_kok = MusteriKumeleyici.kanonik_musteri_anahtari(saf_isim)
            for m_id, m_ad in tum_musteriler:
                m_kok = MusteriKumeleyici.kanonik_musteri_anahtari(m_ad)
                if saf_kok and m_kok and saf_kok == m_kok:
                    return m_id, m_ad

        return None, saf_isim

    @classmethod
    def tek_satir_cift_satis_ayirici(cls, ham_metin):
        """
        '258 kğ ntr pom 65 tl 250 kğ siyah pom 50 tl dgs metal' gibi mesajları
        alt alta iki ayrı siparişe böler ve müşteri adını ikisine de yapıştırır.
        """
        if not ham_metin:
            return ""
        yeni_satirlar = []
        for satir in ham_metin.split('\n'):
            s_kucuk = satir.lower()
            if s_kucuk.count("tl") > 1 and (s_kucuk.count("kg") + s_kucuk.count("kğ") + s_kucuk.count("ton")) > 1:
                bolunmus = re.sub(r'(?i)(tl|lira)\s+(?=\d{1,5}\s*(kg|kğ|ton))', r'\1 | ', satir)
                parcalar = bolunmus.split(" | ")
                son_parca = parcalar[-1]
                musteri_kismi = re.split(r'(?i)(?:tl|lira)[\s\.]+', son_parca)[-1]
                for i in range(len(parcalar) - 1):
                    yeni_satirlar.append(parcalar[i] + " " + musteri_kismi)
                yeni_satirlar.append(parcalar[-1])
            else:
                yeni_satirlar.append(satir)
        return '\n'.join(yeni_satirlar)

    @classmethod
    def bloklari_ayristir(cls, ham_metin, tum_musteriler, tum_urunler):
        ham_metin = cls.tek_satir_cift_satis_ayirici(ham_metin)
        mesaj_listesi = []
        aktif = None
        arama_pos = 0
        for s in ham_metin.split('\n'):
            start_idx = ham_metin.find(s, arama_pos) if s else arama_pos
            if start_idx == -1:
                start_idx = arama_pos
            arama_pos = start_idx + len(s) + 1
            tarih, saat, personel, govde = cls.baslik_ve_mesaj_coz(s)
            if tarih is not None:
                if aktif:
                    mesaj_listesi.append(aktif)
                aktif = {
                    "personel": personel,
                    "tarih": f"{tarih} {saat}".strip() if saat else tarih,
                    "govde": govde,
                    "char_start": start_idx,
                    "char_end": start_idx + len(s)
                }
            elif aktif:
                aktif["govde"] = (aktif["govde"] + "\n" + s.strip()).strip()
                aktif["char_end"] = start_idx + len(s)
            elif s.strip():
                aktif = {
                    "personel": "DÜKKAN",
                    "tarih": "",
                    "govde": s.strip(),
                    "char_start": start_idx,
                    "char_end": start_idx + len(s)
                }
        if aktif:
            mesaj_listesi.append(aktif)

        birlesik = []
        for m_item in mesaj_listesi:
            govde = (m_item.get("govde") or "").strip()
            kg_var = bool(cls.RE_KILO.search(cls.temizle_norm(govde))) or ("ton" in cls.temizle_norm(govde))
            yazar_duzeltme = "yazar" in cls.temizle_norm(govde)
            norm_g = cls.temizle_norm(govde)
            odeme_devam = any(k in norm_g for k in ["odeme", "alindi", "aldim"])
            polimer_var = any(p in norm_g for p in ["pom", "delrin", "derlin", "abs", "pp", "moblen", "hdpe", "pa6", "kristal", "pvc"])
            if birlesik and govde and not cls.cop_satir_mi(norm_g) and (
                not kg_var or yazar_duzeltme or (odeme_devam and not polimer_var)
            ):
                prev = birlesik[-1]
                prev["govde"] = (prev["govde"] + "\n" + govde).strip()
                prev["char_end"] = m_item["char_end"]
            else:
                birlesik.append(m_item)
        mesaj_listesi = birlesik

        sonuclar = []
        for m_item in mesaj_listesi:
            personel = m_item["personel"]
            tarih_saat = m_item["tarih"]
            icerik = m_item["govde"]
            c_start = m_item["char_start"]
            c_end = m_item["char_end"]
            tum_norm = cls.temizle_norm(icerik)
            if cls.cop_satir_mi(tum_norm) or cls.sohbet_ve_rapor_filtresi(icerik):
                continue
            blok_net_kazanc = 0.0
            kazanc_m = cls.RE_KAZANC.search(tum_norm)
            if kazanc_m:
                k_val = float(kazanc_m.group(1).replace(",", "."))
                if "bin" in kazanc_m.group(0) and k_val < 1000:
                    k_val *= 1000.0
                blok_net_kazanc = k_val
            blok_odeme_alindi = 0.0
            odeme_belirsiz = ""
            tam_odeme_ifadesi = bool(re.search(r'odemesi?\s+alindi', tum_norm))
            odeme_m = None if tam_odeme_ifadesi else cls.RE_ODEME.search(tum_norm)
            if odeme_m:
                val = float(odeme_m.group(1).replace(",", "."))
                if "bin" in odeme_m.group(0) and val >= 100:
                    odeme_belirsiz = "Tutar çelişkisi: sayı ve 'bin' birlikte."
                elif "bin" in odeme_m.group(0):
                    val *= 1000.0
                    blok_odeme_alindi = val
                else:
                    blok_odeme_alindi = val
            odeme_usd = 0.0
            odeme_eur = 0.0
            if any(k in tum_norm for k in ["aldim", "aldi", "odeme"]):
                usd_m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:\$|usd|dolar)', tum_norm, flags=re.IGNORECASE)
                eur_m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:€|eur|euro)', tum_norm, flags=re.IGNORECASE)
                if usd_m:
                    odeme_usd = float(usd_m.group(1).replace(",", "."))
                if eur_m:
                    odeme_eur = float(eur_m.group(1).replace(",", "."))
            satirlar = [l.strip() for l in icerik.split('\n') if l.strip()]
            kilo_gecen_satirlar = [s for s in satirlar if cls.kg_coz(s) > 0]
            odeme_kg_satir = [
                s for s in kilo_gecen_satirlar
                if any(k in cls.temizle_norm(s) for k in ["odeme", "alindi"])
                and "yazar" not in cls.temizle_norm(s)
                and not any(p in cls.temizle_norm(s) for p in ["pom", "delrin", "derlin", "abs", "pp", "moblen", "hdpe", "pa6"])
            ]
            satis_kg_satirlar = [s for s in kilo_gecen_satirlar if s not in odeme_kg_satir]
            genel_m_id, genel_m_ad = None, ""
            for s in satirlar:
                mid, mad = cls.musteri_adini_ayikla(s, tum_musteriler)
                if mid and mad and mad not in ["MÜŞTERİSİZ SATIŞ", "GENEL POLİMER"]:
                    genel_m_id, genel_m_ad = mid, mad
                    break
            if not genel_m_ad:
                for s in satirlar:
                    mid, mad = cls.musteri_adini_ayikla(s, tum_musteriler)
                    if mad and mad not in ["MÜŞTERİSİZ SATIŞ", "GENEL POLİMER"]:
                        genel_m_id, genel_m_ad = mid, mad
                        break
            duzeltme = [s for s in satis_kg_satirlar if "yazar" in cls.temizle_norm(s)]
            if duzeltme:
                islem_satirlari = [" ".join(satirlar)]
                satis_kg_satirlar = duzeltme[-1:]
            elif len(satis_kg_satirlar) <= 1:
                islem_satirlari = [" ".join([s for s in satirlar if s not in odeme_kg_satir])]
            else:
                islem_satirlari = satis_kg_satirlar
            blok_sonuclar = []
            for satir in islem_satirlari:
                norm_s = cls.temizle_norm(satir)
                if not any(c.isdigit() for c in norm_s) and not any(p in norm_s for p in ["pom", "abs", "pp", "hdpe", "pa6", "hurda"]):
                    continue
                miktar = 0.0
                kg_hedef = duzeltme[-1] if duzeltme else satir
                miktar = cls.kg_coz(kg_hedef)
                fiyat = None
                fiyat_arama = cls.RE_KAZANC.sub('', norm_s)
                tahsilat_ifadesi = any(k in tum_norm for k in ["aldim", "odeme al", "odeme verdi", "odeme geldi"])
                dolar_m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:\$|usd|dolar)', fiyat_arama, flags=re.IGNORECASE)
                if dolar_m and not tahsilat_ifadesi:
                    fiyat = float(dolar_m.group(1).replace(",", "."))
                else:
                    f_m = cls.RE_FIYAT.search(fiyat_arama)
                    if f_m:
                        fiyat = float(f_m.group(1).replace(",", "."))
                renk = ""
                if "siyah" in norm_s or "syh" in norm_s:
                    renk = "SİYAH"
                elif "beyaz" in norm_s or "byz" in norm_s:
                    renk = "BEYAZ"
                elif "gri" in norm_s:
                    renk = "GRİ"
                elif "krem" in norm_s:
                    renk = "KREM"
                elif any(k in norm_s for k in ["ntr", "nat", "naturel"]):
                    renk = "NATUREL"
                elif any(k in norm_s for k in ["mavi", "turuncu", "kirmizi", "lacivert", "yesil", "sari", "renkli"]):
                    renk = "RENKLİ"
                polimer = ""
                for halk_adi, kimyasal in cls.POLIMER_DONUSUM.items():
                    if re.search(r'\b' + re.escape(halk_adi) + r'\b', norm_s):
                        polimer = kimyasal
                        break
                if not polimer:
                    polimer = "PP" if "moblen" in norm_s else "GENEL POLİMER"
                mevcut_form = ""
                if any(k in norm_s for k in ["granur", "granür", "granul", "granül"]):
                    mevcut_form = "GRANÜL"
                elif any(k in norm_s for k in ["capak", "çapak"]):
                    mevcut_form = "ÇAPAK"
                elif any(k in norm_s for k in ["orj", "orjinal"]):
                    mevcut_form = "ORJİNAL"
                nihai_form = cls.fiyat_ile_form_tahmin_et(polimer, renk, fiyat, mevcut_form, tum_urunler)
                if polimer != "POM" and nihai_form == "GRANÜL":
                    nihai_form = mevcut_form if mevcut_form == "GRANÜL" else ""
                katki = ""
                if any(k in norm_s for k in ["elyafli", "elyaf", "cam elyaf"]):
                    katki = "ELYAFLI"
                elif any(k in norm_s for k in ["kalsitli", "kalsit", "dolgu"]):
                    katki = "KALSİTLİ"
                elif any(k in norm_s for k in ["antisok", "antishock"]):
                    katki = "ANTİŞOK"
                bilesenler = [p for p in [renk, nihai_form, katki, polimer] if p]
                tam_urun_adi = " ".join(bilesenler).strip()
                if miktar <= 0:
                    continue
                satir_id, satir_ad = cls.musteri_adini_ayikla(satir, tum_musteriler)
                if satir_ad and satir_ad not in ["MÜŞTERİSİZ SATIŞ", "GENEL POLİMER"]:
                    m_id, m_ad = satir_id, satir_ad
                else:
                    m_id, m_ad = genel_m_id, genel_m_ad
                bulunan_urun = None
                for p_id, p_ad, p_fiyat in tum_urunler:
                    if cls.temizle_norm(p_ad) == cls.temizle_norm(tam_urun_adi):
                        bulunan_urun = (p_id, p_ad, p_fiyat)
                        break
                nihai_fiyat = float(fiyat) if fiyat is not None else 0.0
                toplam_tutar = round(miktar * nihai_fiyat, 2) if nihai_fiyat else 0.0
                durum = "TAMAM" if miktar > 0 and nihai_fiyat > 0 else ("EKSİK FİYAT" if miktar > 0 else "EKSİK KİLO")
                if odeme_belirsiz:
                    durum = "INCELE"
                blok_sonuclar.append({
                    "tarih": tarih_saat,
                    "personel": personel,
                    "musteri_id": m_id,
                    "musteri_ad": m_ad if m_ad else "MÜŞTERİSİZ SATIŞ",
                    "urun_id": bulunan_urun[0] if bulunan_urun else None,
                    "urun_ad": tam_urun_adi,
                    "miktar": miktar,
                    "fiyat": nihai_fiyat,
                    "tutar": toplam_tutar,
                    "kar": blok_net_kazanc if blok_net_kazanc > 0 else 0.0,
                    "odeme_alindi": 0.0,
                    "odeme_usd": odeme_usd,
                    "odeme_eur": odeme_eur,
                    "durum": durum,
                    "char_start": c_start,
                    "char_end": c_end,
                    "ham_govde": icerik,
                    "ham_musteri_aday": m_ad or "",
                    "belirsizlik": odeme_belirsiz,
                })
            if blok_sonuclar:
                if any(k in tum_norm for k in ["odemesi alindi", "odeme alindi", "odemesi alındı"]) and blok_odeme_alindi <= 0:
                    ilk = blok_sonuclar[0]
                    if ilk["tutar"] > 0 and not odeme_kg_satir:
                        blok_odeme_alindi = ilk["tutar"]
                        blok_sonuclar[0]["belirsizlik"] = (blok_sonuclar[0].get("belirsizlik") or "") + " Tam ödeme: kg x birim."
                if odeme_kg_satir:
                    kaplanan = sum(cls.kg_coz(s) for s in odeme_kg_satir)
                    blok_sonuclar[0]["belirsizlik"] = (
                        (blok_sonuclar[0].get("belirsizlik") or "")
                        + f" Ödeme kapsamı {kaplanan:g} kg; TL tahsilat uydurulmadı."
                    ).strip()
                blok_sonuclar[0]["odeme_alindi"] = blok_odeme_alindi
                for i, kayit in enumerate(blok_sonuclar):
                    kayit["kalem_no"] = i
                    kayit["olay_tipi"] = "SATIS"
                sonuclar.extend(blok_sonuclar)
            else:
                saf_odeme = any(k in tum_norm for k in ["odeme geldi", "odeme verdi", "odeme alindi", "pesin"])
                if saf_odeme or odeme_belirsiz:
                    tutar_tl, celişki = cls.tl_tutari_oku(icerik)
                    if celişki:
                        odeme_belirsiz = celişki
                        tutar_tl = 0.0
                    elif tutar_tl <= 0 and blok_odeme_alindi > 0 and not odeme_belirsiz:
                        tutar_tl = blok_odeme_alindi
                    ad_kaynak = re.sub(
                        r'(?i)\b(odeme|ödeme|geldi|verdi|alindi|alındı|tl|lira|bin|pesin|nakit)\b',
                        ' ', icerik
                    )
                    ad_kaynak = re.sub(r'(?i)plastikten', 'plastik', ad_kaynak)
                    mid, mad = cls.musteri_adini_ayikla(ad_kaynak, tum_musteriler)
                    if not mad or mad in ["MÜŞTERİSİZ SATIŞ", "GENEL POLİMER"]:
                        mid, mad = genel_m_id, genel_m_ad
                    sonuclar.append({
                        "tarih": tarih_saat,
                        "personel": personel,
                        "musteri_id": mid,
                        "musteri_ad": mad if mad and mad != "GENEL POLİMER" else "MÜŞTERİSİZ SATIŞ",
                        "urun_id": None,
                        "urun_ad": "TAHSİLAT",
                        "miktar": 0.0,
                        "fiyat": 0.0,
                        "tutar": 0.0,
                        "kar": 0.0,
                        "odeme_alindi": 0.0 if odeme_belirsiz else tutar_tl,
                        "odeme_usd": odeme_usd,
                        "odeme_eur": odeme_eur,
                        "durum": "INCELE" if odeme_belirsiz else "TAMAM",
                        "char_start": c_start,
                        "char_end": c_end,
                        "ham_govde": icerik,
                        "ham_musteri_aday": mad or "",
                        "belirsizlik": odeme_belirsiz,
                        "kalem_no": 0,
                        "olay_tipi": "TAHSILAT",
                    })
        gorulen = {}
        for r in sonuclar:
            ana = (str(r.get("musteri_ad") or "").upper(), str(r.get("urun_ad") or "").upper(), float(r.get("miktar") or 0), r.get("olay_tipi") or "SATIS")
            gorulen.setdefault(ana, []).append(r)
        for ana, grup in gorulen.items():
            tarihler = {str(x.get("tarih") or "") for x in grup}
            if len(grup) > 1 and len(tarihler) > 1:
                for x in grup:
                    x["olasi_mukerrer"] = True
                    x["durum"] = "INCELE"
                    x["belirsizlik"] = ((x.get("belirsizlik") or "") + " Olası mükerrer/fiyat tamamlama; silinmedi.").strip()
        return sonuclar


def duzenleme_mesafesi(s1, s2):
    if len(s1) < len(s2):
        return duzenleme_mesafesi(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        cur = [i + 1]
        for j, c2 in enumerate(s2):
            cur.append(min(prev[j + 1] + 1, cur[j] + 1, prev[j] + (c1 != c2)))
        prev = cur
    return prev[-1]


def kelimeler_benzer_mi(k1, k2):
    m = max(len(k1), len(k2))
    if m < 4:
        return k1 == k2
    f = duzenleme_mesafesi(k1, k2)
    if m >= 8:
        return f <= 3
    if m >= 6:
        return f <= 2
    return f <= 1


def mevcut_musterilerle_esle(ham_ad, tum_musteriler):
    """Kesin eşleşmede cari id döner. INCELE'de yeni hesap açmaz."""
    musteri_siz = MUSTERISIZ
    ham = str(ham_ad or "").strip() or musteri_siz
    if ham.upper().replace("İ", "I") in ("MUSTERISIZ SATIS", "MÜŞTERİSİZ SATIŞ"):
        for row in tum_musteriler:
            if str(row[1] or "").strip() == musteri_siz:
                rem = float(row[2] or 0.0) if len(row) > 2 else 0.0
                return row[0], musteri_siz, rem
        return None, musteri_siz, 0.0
    try:
        esleyici = MusteriEsleyici(tum_musteriler or [])
        es = esleyici.coz(ham)
    except Exception:
        return None, ham, 0.0
    if es.kesin:
        return es.musteri_id, es.musteri_ad, esleyici.devir(es.musteri_id)
    if es.durum in ("INCELE", "BELIRSIZ", "REFERANS_EKSIK"):
        return None, musteri_siz, 0.0
    return None, ham, 0.0


def wp_musteri_borcunu_guncelle(c, m_id, tutar, tahsilat, mevcut_devir_borcu, onceki_wp_toplami=None, satis_mi=True):
    p = float(onceki_wp_toplami or 0.0)
    t = float(tutar or 0.0)
    d = float(mevcut_devir_borcu or 0.0)
    eklenecek_borc = wp_ek_borc(d, p, t)
    c.execute(
        """
        UPDATE customers
        SET remaining_debt = remaining_debt + ?,
            debt = debt + ?,
            payment = payment + ?,
            shopping_count = shopping_count + ?
        WHERE id = ?
        """,
        (eklenecek_borc, eklenecek_borc, tahsilat, 1 if satis_mi else 0, m_id),
    )
    return p + t


def wp_kaynak_anahtari_uret(tarih, personel, ham_govde, kalem_no):
    ham = f"{tarih or ''}|{personel or ''}|{ham_govde or ''}|{int(kalem_no or 0)}"
    return hashlib.sha256(ham.encode("utf-8")).hexdigest()


def wp_musterisiz_id_al(c):
    c.execute("SELECT id FROM customers WHERE name = ? LIMIT 1", (MUSTERISIZ,))
    row = c.fetchone()
    if row:
        return row[0]
    c.execute("INSERT INTO customers (name, debt, remaining_debt, shopping_count, payment) VALUES (?, 0, 0, 0, 0)", (MUSTERISIZ,))
    return c.lastrowid


def wp_satis_kayit_servisi(c, conn, kayitlar, musteri_map, urun_map, grup_id_bulucu):
    """toplu_kaydet ve ekran kaydı ortak yol. Aynı kaynak ikinci kez yazılmaz."""
    satis_gecmisi_ekle = []
    ozet = {"kaydedilen": 0, "mevcut": 0, "inceleme": 0, "hata": 0}
    c.execute("SELECT id, name, remaining_debt FROM customers")
    tum_musteriler = list(c.fetchall())
    wp_toplam = {}
    devir_d = {}
    onceki_urun_map = dict(urun_map)
    try:
        for d in kayitlar:
            olay = d.get("olay_tipi") or "SATIS"
            m_ad_ham = str(d.get("musteri_ad") or MUSTERISIZ)
            u_ad = str(d.get("urun_ad") or "")
            miktar = float(d.get("miktar", 0.0) or 0.0)
            if olay != "TAHSILAT" and (miktar <= 0 or not u_ad):
                ozet["inceleme"] += 1
                continue
            fiyat = float(d.get("fiyat", 0.0) or 0.0)
            tutar = round(miktar * fiyat, 2) if fiyat else float(d.get("tutar", 0.0) or 0.0)
            tahsilat = float(d.get("odeme_alindi", 0.0) or 0.0)
            tarih = d.get("tarih") or ""
            personel = d.get("personel", "Genel")
            ham_govde = d.get("ham_govde") or u_ad
            kalem_no = d.get("kalem_no", 0)
            anahtar = wp_kaynak_anahtari_uret(tarih, personel, ham_govde, kalem_no)
            c.execute("SELECT id FROM wp_kaynak_olaylar WHERE kaynak_anahtar = ? LIMIT 1", (anahtar,))
            if c.fetchone():
                ozet["mevcut"] += 1
                continue

            parser_id = d.get("musteri_id")
            if parser_id:
                m_id = parser_id
                m_ad = m_ad_ham
                c.execute("SELECT remaining_debt FROM customers WHERE id = ?", (m_id,))
                row = c.fetchone()
                mevcut_devir = float(row[0] or 0.0) if row else 0.0
            else:
                m_id, m_ad, mevcut_devir = mevcut_musterilerle_esle(m_ad_ham, tum_musteriler)
            if not m_id:
                if m_ad and m_ad != MUSTERISIZ and d.get("durum") != "INCELE":
                    c.execute(
                        """INSERT INTO customers (name, debt, payment, remaining_debt, shopping_count)
                           VALUES (?, 0, 0, 0, 0)""",
                        (m_ad,),
                    )
                    m_id = c.lastrowid
                    mevcut_devir = 0.0
                    tum_musteriler.append((m_id, m_ad, 0.0))
                else:
                    m_id = wp_musterisiz_id_al(c)
                    m_ad = MUSTERISIZ
                    mevcut_devir = 0.0
                    tum_musteriler.append((m_id, m_ad, 0.0))

            if m_id not in devir_d:
                devir_d[m_id] = mevcut_devir
            if olay != "TAHSILAT":
                try:
                    _, u_ad = standart_urun_ve_grup_belirle(u_ad, "")
                except Exception:
                    pass
                if not u_ad:
                    ozet["inceleme"] += 1
                    continue
                if u_ad not in urun_map:
                    dogru_gid = grup_id_bulucu(c, u_ad)
                    c.execute("""
                        INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price)
                        VALUES (?, ?, 0.0, 0.0, ?)
                    """, (dogru_gid, u_ad, fiyat or 0.0))
                    urun_map[u_ad] = (c.lastrowid, fiyat)
                u_id = urun_map[u_ad][0]
                c.execute("UPDATE products SET stock_kg = MAX(0, stock_kg - ?) WHERE id = ?", (miktar, u_id))
            wp_toplam[m_id] = wp_musteri_borcunu_guncelle(
                c, m_id, tutar if olay != "TAHSILAT" else 0.0, tahsilat, devir_d[m_id], wp_toplam.get(m_id, 0.0),
                satis_mi=(olay != "TAHSILAT"),
            )
            not_metin = d.get("belirsizlik") or ""
            if d.get("odeme_usd") or d.get("odeme_eur"):
                not_metin = (not_metin + f" USD:{d.get('odeme_usd') or 0} EUR:{d.get('odeme_eur') or 0}").strip()
            c.execute("""
                INSERT INTO sales_history (personnel_name, customer_id, customer_name, product_name, qty, price, total_amount, payment_type, payment_received, created_at, note)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                personel, m_id, m_ad, u_ad, miktar, fiyat, tutar,
                "NAKİT" if tahsilat > 0 else "AÇIK HESAP", tahsilat, tarih, not_metin
            ))
            sid = c.lastrowid
            c.execute("""
                INSERT INTO wp_kaynak_olaylar (kaynak_anahtar, tarih, personel, ham_govde, ham_musteri_aday, belirsizlik, kalem_no, sales_history_id, tutar, tahsilat, durum)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                anahtar, tarih, personel, ham_govde, d.get("ham_musteri_aday") or m_ad_ham,
                d.get("belirsizlik") or "", kalem_no, sid, tutar, tahsilat, d.get("durum") or "TAMAM"
            ))
            ozet["kaydedilen"] += 1
        conn.commit()
    except Exception:
        conn.rollback()
        urun_map.clear()
        urun_map.update(onceki_urun_map)
        ozet["hata"] += 1
        raise
    return ozet


class WhatsAppDevAktarimThread(QThread):
    ilerleme_sinyali = pyqtSignal(int)
    durum_sinyali = pyqtSignal(str)
    tamamlandi_sinyali = pyqtSignal(dict)

    def __init__(self, dosya_yolu, parent=None):
        super().__init__(parent)
        self.dosya_yolu = dosya_yolu
        self.iptal_edildi = False

    def run(self):
        if not os.path.exists(self.dosya_yolu):
            self.durum_sinyali.emit("Hata: Dosya bulunamadı!")
            return
        toplam_boyut = os.path.getsize(self.dosya_yolu)
        okunan_boyut = 0
        toplam_islenen_satir = 0
        toplam_eklenen_satis = 0
        conn = sqlite3.connect(DB_NAME, timeout=60)
        c = conn.cursor()
        try:
            c.execute("SELECT id, name FROM customers")
            tum_musteriler = c.fetchall()
            c.execute("SELECT id, name, sell_price FROM products")
            tum_urunler = c.fetchall()
            musteri_map = {m[1].upper(): m[0] for m in tum_musteriler}
            urun_map = {p[1].upper(): (p[0], p[2]) for p in tum_urunler}
            buffer_blok = []
            metin_havuzu = []
            batch_db_kayit = []
            RE_WP_BASLIK = re.compile(
                r'^[\u200e\ufeff\s]*'
                r'(?:'
                r'\[\d{1,2}[./-]\d{1,2}[./-]\d{2,4},?\s+\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?\]'
                r'|'
                r'\[\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?[,\s]+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\]'
                r'|'
                r'\d{1,2}[./-]\d{1,2}[./-]\d{2,4},?\s+\d{1,2}[:.]\d{2}(?:[:.]\d{2})?(?:\s*[AaPp][Mm])?\s*-\s*'
                r')\s*[^:]+:',
                re.MULTILINE
            )
            with open(self.dosya_yolu, "r", encoding="utf-8", errors="ignore") as f:
                for satir in f:
                    if self.iptal_edildi:
                        break
                    okunan_boyut += len(satir.encode("utf-8"))
                    toplam_islenen_satir += 1
                    temiz_satir = satir.lstrip('\u200e\ufeff')
                    if RE_WP_BASLIK.match(temiz_satir):
                        if buffer_blok:
                            metin_havuzu.append("".join(buffer_blok))
                            buffer_blok.clear()
                        buffer_blok.append(satir)
                        if len(metin_havuzu) >= 200:
                            ayrisan = WhatsAppSatisAyristirici.bloklari_ayristir("".join(metin_havuzu), tum_musteriler, tum_urunler)
                            for it in ayrisan:
                                batch_db_kayit.append(it)
                                toplam_eklenen_satis += 1
                            metin_havuzu.clear()
                    else:
                        buffer_blok.append(satir)
                    if len(batch_db_kayit) >= 5000:
                        self.toplu_kaydet(c, conn, batch_db_kayit, musteri_map, urun_map)
                        batch_db_kayit.clear()
                        gc.collect()
                    if toplam_islenen_satir % 2000 == 0:
                        yuzde = int((okunan_boyut / toplam_boyut) * 100) if toplam_boyut > 0 else 0
                        self.ilerleme_sinyali.emit(min(yuzde, 99))
                        self.durum_sinyali.emit(f"⚡ {toplam_islenen_satir:,} satır tarandı | {toplam_eklenen_satis:,} satış bulundu...")
                if buffer_blok:
                    metin_havuzu.append("".join(buffer_blok))
                if metin_havuzu:
                    ayrisan = WhatsAppSatisAyristirici.bloklari_ayristir("".join(metin_havuzu), tum_musteriler, tum_urunler)
                    for it in ayrisan:
                        batch_db_kayit.append(it)
                        toplam_eklenen_satis += 1
                if batch_db_kayit and not self.iptal_edildi:
                    self.toplu_kaydet(c, conn, batch_db_kayit, musteri_map, urun_map)
                    batch_db_kayit.clear()
            if self.iptal_edildi:
                conn.rollback()
                self.durum_sinyali.emit("Aktarım iptal edildi.")
                return
            conn.commit()
            self.ilerleme_sinyali.emit(100)
            self.tamamlandi_sinyali.emit({
                "toplam_satir": toplam_islenen_satir,
                "toplam_satis": toplam_eklenen_satis
            })
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _csv_satir_sozluk(self, row, basliklar=None):
        if not row:
            return None
        harita = {}
        if basliklar:
            for i, h in enumerate(basliklar):
                harita[str(h or "").strip().lower()] = i

        def al(*adlar, varsayilan=""):
            for ad in adlar:
                if ad in harita and harita[ad] < len(row):
                    return row[harita[ad]]
            return varsayilan

        if harita:
            urun = al("urun", "ürün", "product", "product_name", "malzeme")
            musteri = al("musteri", "müşteri", "customer", "customer_name", "cari")
            qty_h = al("miktar", "qty", "kg", "kilo")
            fiyat_h = al("fiyat", "price", "birim")
            tah_h = al("tahsilat", "odeme", "ödeme", "payment")
            tarih = al("tarih", "date", "created_at")
            personel = al("personel", "personnel", varsayilan="Genel")
        else:
            if len(row) < 4 or not row[2]:
                return None
            tarih, personel, urun = row[0], row[1] or "Genel", row[2]
            qty_h = row[3]
            fiyat_h = row[4] if len(row) > 4 else ""
            tah_h = row[6] if len(row) > 6 else ""
            musteri = ""
        try:
            qty = float(str(qty_h or 0).replace(",", ".") or 0)
        except ValueError:
            return None
        if qty <= 0 or not urun:
            return None
        try:
            pr = float(str(fiyat_h or 0).replace(",", ".") or 0) if fiyat_h else 0.0
        except ValueError:
            pr = 0.0
        try:
            tah = float(str(tah_h or 0).replace(",", ".") or 0) if tah_h else 0.0
        except ValueError:
            tah = 0.0
        return {
            "tarih": str(tarih or ""),
            "personel": str(personel or "Genel"),
            "musteri_ad": str(musteri or "").strip() or MUSTERISIZ,
            "urun_ad": str(urun or ""),
            "miktar": qty,
            "fiyat": pr,
            "odeme_alindi": tah,
        }

    def urun_icin_dogru_grup_id_bul(self, c, u_ad, grup_id_map):
        u_up = u_ad.upper()
        if "POM" in u_up: hedef = "POM"
        elif "ABS" in u_up: hedef = "ABS"
        elif any(k in u_up for k in ["HDPE", "ELTEKS", "I20", "PE100"]): hedef = "HDPE"
        elif any(k in u_up for k in ["PP", "MOBLEN"]): hedef = "PP MOBLEN"
        elif "PA66" in u_up or "N66" in u_up: hedef = "PA66"
        elif "PA6" in u_up or "N6" in u_up: hedef = "PA6"
        elif "PVC" in u_up: hedef = "PVC"
        elif any(k in u_up for k in ["ANTİŞOK", "HIPS", "ANT"]): hedef = "ANTİŞOK"
        else: hedef = "GRUPSUZLAR"
        if hedef not in grup_id_map:
            c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (hedef,))
            row = c.fetchone()
            if row:
                grup_id_map[hedef] = row[0]
            else:
                c.execute("INSERT INTO product_groups (name) VALUES (?)", (hedef,))
                grup_id_map[hedef] = c.lastrowid
        return grup_id_map[hedef]

    def toplu_kaydet(self, c, conn, kayitlar, musteri_map, urun_map):
        if not hasattr(self, 'grup_id_map'):
            self.grup_id_map = {}

        def grup_id_bulucu(cur, u_ad):
            return self.urun_icin_dogru_grup_id_bul(cur, u_ad, self.grup_id_map)

        return wp_satis_kayit_servisi(c, conn, kayitlar, musteri_map, urun_map, grup_id_bulucu)


class PersonelEkleDialog(QDialog):
    def __init__(self, parent=None, p_id=None):
        super().__init__(parent)
        self.p_id = p_id
        self.setWindowTitle("Personel Kartı" if not p_id else "Personel Düzenle")
        self.setFixedSize(450, 360)
        self.setStyleSheet("background-color: #ffffff;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        lbl_baslik = QLabel("YENİ PERSONEL EKLE" if not self.p_id else "PERSONELİ DÜZENLE")
        lbl_baslik.setFixedHeight(38)
        lbl_baslik.setStyleSheet("background-color: #0284c7; color: white; font-size: 14px; font-weight: bold; padding-left: 12px; border-radius: 4px;")
        layout.addWidget(lbl_baslik)

        layout.addWidget(QLabel("Personel Adı Soyadı (*):", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        self.txt_ad_soyad = BuyukHarfKutusu("Örn: AHMET YILMAZ veya BRO AHMET")
        self.txt_ad_soyad.setFixedHeight(36)
        self.txt_ad_soyad.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        layout.addWidget(self.txt_ad_soyad)

        layout.addWidget(QLabel("Görevi / Ünvanı:", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        self.txt_gorev = BuyukHarfKutusu("Örn: SATIŞ DANIŞMANI, DEPO SORUMLUSU")
        self.txt_gorev.setFixedHeight(36)
        self.txt_gorev.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        layout.addWidget(self.txt_gorev)

        layout.addWidget(QLabel("Telefon Numarası:", styleSheet="font-size: 12px; font-weight: bold; color: #334155;"))
        self.txt_telefon = QLineEdit()
        self.txt_telefon.setPlaceholderText("05XX XXX XX XX")
        self.txt_telefon.setFixedHeight(36)
        self.txt_telefon.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 13px;")
        layout.addWidget(self.txt_telefon)
        layout.addStretch()

        h_btns = QHBoxLayout()
        h_btns.addStretch()
        btn_iptal = QPushButton("İptal")
        btn_iptal.setFixedHeight(36)
        btn_iptal.setFixedWidth(80)
        btn_iptal.setStyleSheet("background-color: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold;")
        btn_iptal.clicked.connect(self.reject)
        h_btns.addWidget(btn_iptal)

        btn_kaydet = QPushButton("  ✓ Personeli Kaydet  ")
        btn_kaydet.setFixedHeight(36)
        btn_kaydet.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_kaydet.setStyleSheet("background-color: #10b981; color: white; border: none; border-radius: 4px; font-weight: bold; padding: 0 16px; font-size: 13px;")
        btn_kaydet.clicked.connect(self.kaydet)
        h_btns.addWidget(btn_kaydet)
        layout.addLayout(h_btns)
        if self.p_id:
            self.verileri_yukle()

    def verileri_yukle(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT name, role, phone FROM personnel WHERE id = ?", (self.p_id,))
        row = c.fetchone()
        conn.close()
        if row:
            self.txt_ad_soyad.setText(row[0] or "")
            self.txt_gorev.setText(row[1] or "")
            self.txt_telefon.setText(row[2] or "")

    def kaydet(self):
        ad_soyad = buyuk_harf(self.txt_ad_soyad.text().strip())
        gorev = buyuk_harf(self.txt_gorev.text().strip()) or "SATIŞ SORUMLUSU"
        tel = self.txt_telefon.text().strip()
        if not ad_soyad:
            QMessageBox.warning(self, "Eksik Bilgi", "Lütfen personel adını ve soyadını giriniz!")
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        try:
            if not self.p_id:
                c.execute("INSERT INTO personnel (name, full_name, role, phone) VALUES (?, ?, ?, ?)", (ad_soyad, ad_soyad, gorev, tel))
            else:
                c.execute("UPDATE personnel SET name = ?, full_name = ?, role = ?, phone = ? WHERE id = ?", (ad_soyad, ad_soyad, gorev, tel, self.p_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Başarılı", f"'{ad_soyad}' başarıyla kaydedildi!")
            self.accept()
        except sqlite3.IntegrityError:
            conn.close()
            QMessageBox.warning(self, "Zaten Kayıtlı", f"'{ad_soyad}' isimli personel sistemde zaten mevcut!")


class MusteriIletisimDialog(QDialog):
    def __init__(self, musteri_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Cari Kart: {musteri_data.get('name', '')}")
        self.setFixedSize(420, 320)
        self.setStyleSheet("background-color: #ffffff;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        lbl_b = QLabel("MÜŞTERİ KÜNYESİ")
        lbl_b.setStyleSheet("font-size: 14px; font-weight: bold; color: #0284c7; border-bottom: 2px solid #0284c7; padding-bottom: 4px;")
        layout.addWidget(lbl_b)

        bilgiler = [
            ("Müşteri / Ünvan:", musteri_data.get("name", "-")),
            ("Telefon:", musteri_data.get("phone", "-")),
            ("Vergi Dairesi:", musteri_data.get("tax_office", "-")),
            ("Vergi / TCKN No:", musteri_data.get("tax_number", "-")),
            ("Açık Hesap Limiti:", f"{float(musteri_data.get('credit_limit') or 0.0):,.2f} ₺"),
            ("Kayıtlı Adres:", musteri_data.get("address", "-"))
        ]

        for baslik, deger in bilgiler:
            h = QHBoxLayout()
            lbl_k = QLabel(baslik)
            lbl_k.setStyleSheet("font-weight: bold; color: #475569; width: 130px;")
            lbl_v = QLabel(str(deger) if deger else "-")
            lbl_v.setStyleSheet("color: #0f172a;")
            lbl_v.setWordWrap(True)
            h.addWidget(lbl_k, stretch=2)
            h.addWidget(lbl_v, stretch=3)
            layout.addLayout(h)

        layout.addStretch()
        btn_kapat = QPushButton("Kapat")
        btn_kapat.setFixedHeight(34)
        btn_kapat.setStyleSheet("background: #f1f5f9; font-weight: bold; border-radius: 4px; border: 1px solid #cbd5e1;")
        btn_kapat.clicked.connect(self.accept)
        layout.addWidget(btn_kapat)


class ModernBarChartWidget(QWidget):
    """macOS uyumlu, yuvarlatılmış çubuk (bar) grafiği bileşeni"""
    def __init__(self, baslik="Satış Grafiği", birim="₺", parent=None):
        super().__init__(parent)
        self.baslik = baslik
        self.birim = birim
        self.veriler = []
        self.setMinimumHeight(240)
        self.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 8px;")

    def veri_guncelle(self, veri_listesi):
        self.veriler = veri_listesi
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        painter.setFont(QFont("SF Pro Display", 11, QFont.Weight.Bold))
        painter.setPen(QColor("#1e293b"))
        painter.drawText(16, 26, self.baslik)
        if not self.veriler:
            painter.setFont(QFont("SF Pro Text", 10))
            painter.setPen(QColor("#94a3b8"))
            painter.drawText(QRectF(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, "Bu aralıkta gösterilecek veri bulunamadı.")
            return
        max_deger = max([v[1] for v in self.veriler]) if self.veriler else 1.0
        if max_deger <= 0:
            max_deger = 1.0
        alt_bosluk = 40
        ust_bosluk = 50
        sol_bosluk = 20
        sag_bosluk = 20
        cizim_w = w - sol_bosluk - sag_bosluk
        cizim_h = h - ust_bosluk - alt_bosluk
        n = len(self.veriler)
        bar_w = min(54.0, (cizim_w / max(n, 1)) * 0.65)
        aralik = cizim_w / max(n, 1)
        for i, (etiket, deger, renk_hex) in enumerate(self.veriler):
            oran = float(deger) / max_deger
            bar_h = max(6.0, cizim_h * oran)
            x = sol_bosluk + (i * aralik) + (aralik - bar_w) / 2.0
            y = ust_bosluk + (cizim_h - bar_h)
            painter.setBrush(QBrush(QColor(renk_hex)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(QRectF(x, y, bar_w, bar_h), 5, 5)
            painter.setFont(QFont("SF Pro Text", 8, QFont.Weight.DemiBold))
            painter.setPen(QColor("#334155"))
            val_txt = f"{deger:,.0f} {self.birim}" if deger < 100000 else f"{deger/1000:,.1f}K {self.birim}"
            painter.drawText(QRectF(x - 25, y - 20, bar_w + 50, 18), Qt.AlignmentFlag.AlignCenter, val_txt)
            painter.setFont(QFont("SF Pro Text", 9))
            painter.setPen(QColor("#64748b"))
            painter.drawText(QRectF(x - 20, h - alt_bosluk + 8, bar_w + 40, 24), Qt.AlignmentFlag.AlignHCenter, str(etiket)[:9])


class ModernDonutChartWidget(QWidget):
    """Ödeme tipleri ve grup payları için modern halka grafiği"""
    def __init__(self, baslik="Dağılım", parent=None):
        super().__init__(parent)
        self.baslik = baslik
        self.dilimler = []
        self.setMinimumHeight(240)
        self.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 8px;")

    def veri_guncelle(self, dilim_listesi):
        self.dilimler = dilim_listesi
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        painter.setFont(QFont("SF Pro Display", 11, QFont.Weight.Bold))
        painter.setPen(QColor("#1e293b"))
        painter.drawText(16, 26, self.baslik)
        toplam = sum([v[1] for v in self.dilimler])
        if toplam <= 0 or not self.dilimler:
            painter.setFont(QFont("SF Pro Text", 10))
            painter.setPen(QColor("#94a3b8"))
            painter.drawText(QRectF(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, "Veri yok.")
            return
        cap = min(w * 0.42, h * 0.65)
        cx, cy = (w * 0.32), (h * 0.55)
        donut_rect = QRectF(cx - cap/2, cy - cap/2, cap, cap)
        baslangic_acisi = 90.0 * 16
        for etiket, deger, renk in self.dilimler:
            if deger <= 0:
                continue
            span = (deger / toplam) * 360.0 * 16
            painter.setBrush(QBrush(QColor(renk)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawPie(donut_rect, int(baslangic_acisi), int(-span))
            baslangic_acisi -= span
        delik_cap = cap * 0.58
        painter.setBrush(QBrush(QColor("white")))
        painter.drawEllipse(QRectF(cx - delik_cap/2, cy - delik_cap/2, delik_cap, delik_cap))
        lx = w * 0.58
        ly = 60
        painter.setFont(QFont("SF Pro Text", 9, QFont.Weight.Medium))
        for etiket, deger, renk in self.dilimler:
            if deger <= 0:
                continue
            yuzde = (deger / toplam) * 100.0
            painter.setBrush(QBrush(QColor(renk)))
            painter.drawRoundedRect(QRectF(lx, ly, 10, 10), 2, 2)
            painter.setPen(QColor("#334155"))
            painter.drawText(int(lx + 18), int(ly + 9), f"{etiket} (%{yuzde:.1f})")
            ly += 24


_GRAFIK_RENKLER = [
    "#4285F4", "#EA4335", "#FBBC04", "#34A853", "#FF6D01",
    "#46BDC6", "#7BAAF7", "#F07B72", "#FCD04F", "#71C287",
]


class _DinamikGrafikTuval(QWidget):
    def __init__(self, motor):
        super().__init__()
        self.motor = motor
        self.setMinimumHeight(280)
        self.setStyleSheet("background: white;")

    def paintEvent(self, event):
        self.motor._tuval_ciz(QPainter(self), self.width(), self.height())


class DinamikRaporGrafigi(QWidget):
    """Google Data Studio paleti, K/M eksen, satır kırık etiket — PyQt6 (matplotlib yok)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.satirlar = []
        self.col_map = {"urun": 3, "per": 1, "musteri": 2, "kg": 4, "tutar": 5, "odeme": 6}
        self._etiketler = []
        self._degerler = []
        self._baslik = ""
        self.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 8px;")
        yerlesim = QVBoxLayout(self)
        yerlesim.setContentsMargins(12, 12, 12, 12)

        panel = QHBoxLayout()
        panel.setSpacing(10)
        cmb_stil = "padding: 4px; border: 1px solid #cbd5e1; border-radius: 4px; font-weight: bold;"

        self.cmb_analiz = QComboBox()
        self.cmb_analiz.addItems(["Ürün Dağılımı", "Hammadde Grupları", "Müşteri Dağılımı", "Personel Dağılımı", "Ödeme Tipleri"])
        self.cmb_analiz.setStyleSheet(cmb_stil)

        self.cmb_metrik = QComboBox()
        self.cmb_metrik.addItems(["Ciro (₺)", "Tonaj (KG)", "İşlem Adedi"])
        self.cmb_metrik.setStyleSheet(cmb_stil)

        self.cmb_grafik = QComboBox()
        self.cmb_grafik.addItems([
            "1. Dikey Çubuk (Bar)", "2. Yatay Çubuk (Barh)", "3. Pasta Grafiği (Pie)",
            "4. Halka Grafiği (Donut)", "5. Çizgi Grafiği (Line)", "6. Alan Grafiği (Area)",
            "7. Lolipop Grafiği (Stem)", "8. Basamak Grafiği (Step)", "9. Nokta Saçılım (Scatter)",
            "10. Yüzde Çubuk (Percent Bar)",
        ])
        self.cmb_grafik.setStyleSheet(cmb_stil)

        panel.addWidget(QLabel("<b>Eksen:</b>"))
        panel.addWidget(self.cmb_analiz)
        panel.addWidget(QLabel("<b>Değer:</b>"))
        panel.addWidget(self.cmb_metrik)
        panel.addWidget(QLabel("<b>Görünüm:</b>"))
        panel.addWidget(self.cmb_grafik)
        panel.addStretch()
        yerlesim.addLayout(panel)

        self.tuval = _DinamikGrafikTuval(self)
        yerlesim.addWidget(self.tuval)

        self.cmb_analiz.currentIndexChanged.connect(self.grafik_ciz)
        self.cmb_metrik.currentIndexChanged.connect(self.grafik_ciz)
        self.cmb_grafik.currentIndexChanged.connect(self.grafik_ciz)

    def veri_yukle(self, satirlar, col_map=None):
        self.satirlar = satirlar or []
        if col_map:
            self.col_map = col_map
        self.grafik_ciz()

    @staticmethod
    def format_sayi(x, pos=None):
        if x >= 1e6:
            return f"{x * 1e-6:.1f}M"
        if x >= 1e3:
            return f"{x * 1e-3:.1f}K"
        return f"{x:.0f}"

    @staticmethod
    def _kirik_etiket(metin):
        return "\n".join(textwrap.wrap(str(metin or ""), width=12) or [""])

    def _hammadde_grup(self, urun):
        u = (urun or "").upper()
        if "POM" in u:
            return "POM"
        if "ABS" in u:
            return "ABS"
        if any(k in u for k in ("HDPE", "ELTEKS", "I20")):
            return "HDPE"
        if any(k in u for k in ("PP", "MOBLEN")):
            return "PP MOBLEN"
        if "PA" in u:
            return "POLİAMİD (PA)"
        if "PVC" in u:
            return "PVC"
        if any(k in u for k in ("ANTİŞOK", "HIPS", "ANT")):
            return "ANTİŞOK"
        return "DİĞER"

    def grafik_ciz(self):
        analiz = self.cmb_analiz.currentText()
        metrik = self.cmb_metrik.currentText()
        veri = {}
        for s in self.satirlar:
            try:
                urun = str(s[self.col_map["urun"]] or "Bilinmeyen")
                per = str(s[self.col_map["per"]] or "Dükkan")
                mus = str(s[self.col_map["musteri"]] or "Müşterisiz")
                kg = float(s[self.col_map["kg"]] or 0.0)
                tutar = float(s[self.col_map["tutar"]] or 0.0)
                odeme = str(s[self.col_map.get("odeme", 6)] or "Bilinmeyen")
            except (IndexError, TypeError, ValueError):
                continue
            if analiz == "Ürün Dağılımı":
                anahtar = urun
            elif analiz == "Hammadde Grupları":
                anahtar = self._hammadde_grup(urun)
            elif analiz == "Müşteri Dağılımı":
                anahtar = mus
            elif analiz == "Personel Dağılımı":
                anahtar = per
            else:
                anahtar = odeme
            if anahtar not in veri:
                veri[anahtar] = {"tutar": 0.0, "kg": 0.0, "adet": 0}
            veri[anahtar]["tutar"] += tutar
            veri[anahtar]["kg"] += kg
            veri[anahtar]["adet"] += 1
        kriter = "tutar" if "Ciro" in metrik else "kg" if "Tonaj" in metrik else "adet"
        sirali = sorted(veri.items(), key=lambda x: x[1][kriter], reverse=True)[:15]
        self._etiketler = [k for k, _ in sirali]
        self._degerler = [v[kriter] for _, v in sirali]
        self._baslik = f"İlk 15: {analiz} - {metrik}"
        self.tuval.update()

    def _tuval_ciz(self, painter, w, h):
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(0, 0, w, h, QColor("#ffffff"))
        painter.setFont(QFont("SF Pro Display", 12, QFont.Weight.Bold))
        painter.setPen(QColor("#1e293b"))
        painter.drawText(12, 22, self._baslik)
        if not self.satirlar or not self._etiketler or sum(self._degerler) == 0:
            painter.setFont(QFont("SF Pro Text", 11))
            painter.setPen(QColor("#64748b"))
            painter.drawText(QRectF(0, 0, w, h), Qt.AlignmentFlag.AlignCenter,
                             "Bu tarih aralığında gösterilecek veri yok.")
            return
        tur = self.cmb_grafik.currentText()
        etiketler, degerler = self._etiketler, self._degerler
        n = len(etiketler)
        renkler = [_GRAFIK_RENKLER[i % len(_GRAFIK_RENKLER)] for i in range(n)]
        if "3." in tur or "4." in tur:
            self._ciz_pasta(painter, w, h, etiketler, degerler, renkler, donut="4." in tur)
            return
        if "2." in tur:
            self._ciz_yatay(painter, w, h, etiketler, degerler, renkler)
            return
        cizilecek = degerler
        yuzde_mod = "10." in tur
        if yuzde_mod:
            toplam = sum(degerler) or 1.0
            cizilecek = [(d / toplam) * 100.0 for d in degerler]
        self._ciz_eksenli(painter, w, h, etiketler, cizilecek, renkler, tur, yuzde_mod)

    def _ciz_kirik_yazi(self, painter, rect, metin, hiza):
        painter.drawText(rect, hiza | Qt.TextFlag.TextWordWrap, self._kirik_etiket(metin))

    def _ciz_pasta(self, painter, w, h, etiketler, degerler, renkler, donut=False):
        toplam = sum(degerler)
        if toplam <= 0:
            return
        cap = min(w * 0.40, (h - 36) * 0.70)
        cx, cy = w * 0.30, h * 0.56
        dilim = QRectF(cx - cap / 2, cy - cap / 2, cap, cap)
        aci = 90.0 * 16
        for deger, renk in zip(degerler, renkler):
            if deger <= 0:
                continue
            span = (deger / toplam) * 360.0 * 16
            painter.setBrush(QBrush(QColor(renk)))
            painter.setPen(QPen(QColor("white"), 1.5))
            painter.drawPie(dilim, int(aci), int(-span))
            aci -= span
        if donut:
            delik = cap * 0.58
            painter.setBrush(QBrush(QColor("white")))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(QRectF(cx - delik / 2, cy - delik / 2, delik, delik))
        lx, ly = w * 0.56, 40
        painter.setFont(QFont("SF Pro Text", 8, QFont.Weight.Medium))
        for etiket, deger, renk in zip(etiketler, degerler, renkler):
            if deger <= 0:
                continue
            painter.setBrush(QBrush(QColor(renk)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(QRectF(lx, ly, 9, 9), 2, 2)
            painter.setPen(QColor("#334155"))
            kirik = self._kirik_etiket(etiket).replace("\n", " ")
            painter.drawText(int(lx + 14), int(ly + 9), f"{kirik[:18]}  {self.format_sayi(deger)}  (%{(deger / toplam) * 100:.1f})")
            ly += 20

    def _ciz_yatay(self, painter, w, h, etiketler, degerler, renkler):
        max_d = max(degerler) or 1.0
        ust, alt, sol, sag = 40, 18, 108, 52
        cw, ch = w - sol - sag, h - ust - alt
        n = len(etiketler)
        satir_h = ch / max(n, 1)
        cubuk_h = min(18.0, satir_h * 0.55)
        painter.setPen(QPen(QColor("#cbd5e1"), 1))
        for k in range(4):
            gx = sol + cw * k / 3.0
            painter.drawLine(QPointF(gx, ust), QPointF(gx, ust + ch))
            painter.setFont(QFont("SF Pro Text", 8))
            painter.setPen(QColor("#475569"))
            painter.drawText(QRectF(gx - 24, ust + ch + 2, 48, 14), Qt.AlignmentFlag.AlignHCenter,
                             self.format_sayi(max_d * k / 3.0))
            painter.setPen(QPen(QColor("#cbd5e1"), 1))
        painter.setFont(QFont("SF Pro Text", 8))
        for i, (etiket, deger, renk) in enumerate(zip(etiketler, degerler, renkler)):
            y = ust + i * satir_h + (satir_h - cubuk_h) / 2
            bw = max(4.0, (deger / max_d) * cw * 0.92)
            painter.setPen(QColor("#475569"))
            self._ciz_kirik_yazi(
                painter, QRectF(4, y - 6, sol - 10, cubuk_h + 14), etiket,
                Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight,
            )
            painter.setBrush(QBrush(QColor(renk)))
            painter.setPen(QPen(QColor("white"), 1))
            painter.drawRoundedRect(QRectF(sol, y, bw, cubuk_h), 4, 4)
            painter.setPen(QColor("#334155"))
            painter.setFont(QFont("SF Pro Text", 7, QFont.Weight.Bold))
            painter.drawText(QRectF(sol + bw + 4, y, 48, cubuk_h), Qt.AlignmentFlag.AlignVCenter, self.format_sayi(deger))

    def _ciz_eksenli(self, painter, w, h, etiketler, degerler, renkler, tur, yuzde_mod):
        max_d = max(degerler) or 1.0
        ust, alt, sol, sag = 44, 58, 46, 14
        cw, ch = w - sol - sag, h - ust - alt
        n = len(etiketler)
        aralik = cw / max(n, 1)
        noktalar = []
        for i, deger in enumerate(degerler):
            x = sol + i * aralik + aralik / 2
            y = ust + ch - (deger / max_d) * (ch * 0.92)
            noktalar.append(QPointF(x, y))
        painter.setPen(QPen(QColor("#cbd5e1"), 1))
        painter.setFont(QFont("SF Pro Text", 8))
        for k in range(4):
            gy = ust + ch * k / 3.0
            painter.setPen(QPen(QColor("#cbd5e1"), 1))
            painter.drawLine(QPointF(sol, gy), QPointF(sol + cw, gy))
            val = max_d * (1 - k / 3.0)
            painter.setPen(QColor("#475569"))
            yazi = f"{val:.0f}%" if yuzde_mod else self.format_sayi(val)
            painter.drawText(QRectF(2, gy - 8, sol - 6, 16), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, yazi)
        if "1." in tur or "10." in tur:
            bar_w = min(40.0, aralik * 0.60)
            bar_renk = "#46BDC6" if "10." in tur else None
            for i, (p, renk, deger) in enumerate(zip(noktalar, renkler, degerler)):
                bh = max(4.0, ust + ch - p.y())
                painter.setBrush(QBrush(QColor(bar_renk or renk)))
                painter.setPen(QPen(QColor("white"), 1))
                painter.drawRoundedRect(QRectF(p.x() - bar_w / 2, p.y(), bar_w, bh), 4, 4)
                painter.setPen(QColor("#334155"))
                painter.setFont(QFont("SF Pro Text", 7, QFont.Weight.Bold))
                etiket_d = f"{deger:.1f}%" if yuzde_mod else self.format_sayi(deger)
                painter.drawText(QRectF(p.x() - 28, p.y() - 16, 56, 14), Qt.AlignmentFlag.AlignHCenter, etiket_d)
        elif "5." in tur or "6." in tur:
            cizgi = QColor("#34A853" if "6." in tur else "#4285F4")
            if "6." in tur and noktalar:
                yol = QPainterPath()
                yol.moveTo(noktalar[0].x(), ust + ch)
                for p in noktalar:
                    yol.lineTo(p)
                yol.lineTo(noktalar[-1].x(), ust + ch)
                yol.closeSubpath()
                painter.setBrush(QBrush(QColor(52, 168, 83, 50)))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPath(yol)
            painter.setPen(QPen(cizgi, 2.5))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            for i in range(len(noktalar) - 1):
                painter.drawLine(noktalar[i], noktalar[i + 1])
            painter.setBrush(QBrush(cizgi))
            for p in noktalar:
                painter.setPen(QPen(QColor("white"), 1.5))
                painter.drawEllipse(p, 4, 4)
        elif "7." in tur:
            for p in noktalar:
                painter.setPen(QPen(QColor("#EA4335"), 2))
                painter.drawLine(QPointF(p.x(), ust + ch), p)
                painter.setBrush(QBrush(QColor("#EA4335")))
                painter.setPen(QPen(QColor("white"), 1.5))
                painter.drawEllipse(p, 5, 5)
        elif "8." in tur:
            yol = QPainterPath()
            dolgu = QPainterPath()
            if noktalar:
                dolgu.moveTo(noktalar[0].x() - aralik / 2, ust + ch)
                for i, p in enumerate(noktalar):
                    x0, x1 = p.x() - aralik / 2, p.x() + aralik / 2
                    if i == 0:
                        yol.moveTo(x0, p.y())
                    else:
                        yol.lineTo(x0, p.y())
                    yol.lineTo(x1, p.y())
                    dolgu.lineTo(x0, p.y())
                    dolgu.lineTo(x1, p.y())
                dolgu.lineTo(noktalar[-1].x() + aralik / 2, ust + ch)
                dolgu.closeSubpath()
                painter.setBrush(QBrush(QColor(251, 188, 4, 50)))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawPath(dolgu)
                painter.setPen(QPen(QColor("#FBBC04"), 2.5))
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawPath(yol)
        elif "9." in tur:
            for p, deger, renk in zip(noktalar, degerler, renkler):
                r = max(5.0, (deger / max_d) * 14.0)
                painter.setBrush(QBrush(QColor(renk)))
                painter.setPen(QPen(QColor("white"), 1.5))
                painter.drawEllipse(p, r, r)
        painter.setFont(QFont("SF Pro Text", 8))
        painter.setPen(QColor("#475569"))
        for i, etiket in enumerate(etiketler):
            x = sol + i * aralik
            self._ciz_kirik_yazi(
                painter, QRectF(x, h - alt + 2, aralik, alt - 4), etiket, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
            )


class GoogleStilGrafik(QWidget):
    """Google Sheets/Docs grafiklerine benzeyen, sade, düz renkli bar/çizgi/pasta grafik.
    Harici kütüphane gerektirmez, PyQt'nin kendi QPainter'ıyla çizilir."""

    PALET = ["#4285F4", "#EA4335", "#FBBC04", "#34A853", "#9C27B0", "#00ACC1", "#FF7043", "#7E57C2"]

    def __init__(self, tur="bar", parent=None):
        super().__init__(parent)
        self.tur = tur
        self.etiketler = []
        self.degerler = []
        self.baslik = ""
        self.setMinimumHeight(260)

    def veri_ayarla(self, etiketler, degerler, baslik=""):
        self.etiketler = etiketler
        self.degerler = degerler
        self.baslik = baslik
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        p.fillRect(0, 0, w, h, QColor("white"))

        if not self.degerler or all(v == 0 for v in self.degerler):
            p.setPen(QColor("#94a3b8"))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Bu aralıkta veri yok")
            p.end()
            return

        ust_pay = 34 if self.baslik else 12
        if self.baslik:
            p.setPen(QColor("#334155"))
            p.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            p.drawText(12, 22, self.baslik)

        if self.tur == "pasta":
            self._pasta_ciz(p, w, h, ust_pay)
        elif self.tur == "cizgi":
            self._cizgi_ciz(p, w, h, ust_pay)
        else:
            self._bar_ciz(p, w, h, ust_pay)
        p.end()

    def _bar_ciz(self, p, w, h, ust_pay):
        sol_pay, alt_pay, sag_pay = 55, 34, 20
        cizim_genislik = w - sol_pay - sag_pay
        cizim_yukseklik = h - ust_pay - alt_pay
        maks_deger = max(self.degerler) * 1.15 or 1

        p.setPen(QColor("#e2e8f0"))
        for i in range(5):
            y = ust_pay + cizim_yukseklik - (cizim_yukseklik * i / 4)
            p.drawLine(sol_pay, int(y), w - sag_pay, int(y))
            p.setPen(QColor("#94a3b8"))
            p.setFont(QFont("Segoe UI", 8))
            deger = maks_deger * i / 4
            p.drawText(4, int(y) + 4, f"{deger:,.0f}")
            p.setPen(QColor("#e2e8f0"))

        n = len(self.degerler)
        bar_genislik = cizim_genislik / n * 0.55
        bosluk = cizim_genislik / n
        for i, (etiket, deger) in enumerate(zip(self.etiketler, self.degerler)):
            x = sol_pay + i * bosluk + (bosluk - bar_genislik) / 2
            bar_yukseklik = (deger / maks_deger) * cizim_yukseklik if maks_deger else 0
            y = ust_pay + cizim_yukseklik - bar_yukseklik
            renk = QColor(self.PALET[i % len(self.PALET)])
            p.setBrush(renk)
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(int(x), int(y), int(bar_genislik), int(bar_yukseklik), 3, 3)
            p.setPen(QColor("#334155"))
            p.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
            p.drawText(int(x - 10), int(y) - 4, int(bar_genislik + 20), 14,
                       Qt.AlignmentFlag.AlignCenter, f"{deger:,.0f}")
            p.setPen(QColor("#64748b"))
            p.setFont(QFont("Segoe UI", 8))
            p.drawText(int(x - 20), h - alt_pay + 6, int(bar_genislik + 40), 20,
                       Qt.AlignmentFlag.AlignCenter, str(etiket))

    def _cizgi_ciz(self, p, w, h, ust_pay):
        sol_pay, alt_pay, sag_pay = 55, 34, 20
        cizim_genislik = w - sol_pay - sag_pay
        cizim_yukseklik = h - ust_pay - alt_pay
        maks_deger = max(self.degerler) * 1.15 or 1
        n = len(self.degerler)

        p.setPen(QColor("#e2e8f0"))
        for i in range(5):
            y = ust_pay + cizim_yukseklik - (cizim_yukseklik * i / 4)
            p.drawLine(sol_pay, int(y), w - sag_pay, int(y))
            p.setPen(QColor("#94a3b8"))
            p.setFont(QFont("Segoe UI", 8))
            p.drawText(4, int(y) + 4, f"{(maks_deger * i / 4):,.0f}")
            p.setPen(QColor("#e2e8f0"))

        noktalar = []
        for i, deger in enumerate(self.degerler):
            x = sol_pay + (cizim_genislik * i / max(1, n - 1))
            y = ust_pay + cizim_yukseklik - (deger / maks_deger) * cizim_yukseklik
            noktalar.append((x, y))

        dolgu = QPolygon()
        dolgu.append(QPoint(int(noktalar[0][0]), int(ust_pay + cizim_yukseklik)))
        for x, y in noktalar:
            dolgu.append(QPoint(int(x), int(y)))
        dolgu.append(QPoint(int(noktalar[-1][0]), int(ust_pay + cizim_yukseklik)))
        p.setBrush(QColor(66, 133, 244, 40))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawPolygon(dolgu)

        p.setPen(QColor("#4285F4"))
        pen = p.pen()
        pen.setWidth(3)
        p.setPen(pen)
        for i in range(len(noktalar) - 1):
            p.drawLine(int(noktalar[i][0]), int(noktalar[i][1]), int(noktalar[i + 1][0]), int(noktalar[i + 1][1]))
        p.setBrush(QColor("#4285F4"))
        p.setPen(Qt.PenStyle.NoPen)
        for x, y in noktalar:
            p.drawEllipse(QPoint(int(x), int(y)), 4, 4)

        p.setPen(QColor("#64748b"))
        p.setFont(QFont("Segoe UI", 8))
        gosterilecek = self.etiketler if n <= 8 else self.etiketler[::max(1, n // 8)]
        adim = max(1, n // max(1, len(gosterilecek)))
        for i in range(0, n, adim):
            x, _ = noktalar[i]
            p.drawText(int(x - 25), h - alt_pay + 6, 50, 20, Qt.AlignmentFlag.AlignCenter, str(self.etiketler[i]))

    def _pasta_ciz(self, p, w, h, ust_pay):
        toplam = sum(self.degerler) or 1
        cap = min(w - 160, h - ust_pay - 20)
        cap = max(cap, 60)
        cx, cy = 20 + cap / 2, ust_pay + (h - ust_pay - 20) / 2
        aci_baslangic = 90 * 16
        for i, deger in enumerate(self.degerler):
            aci_pay = int(360 * 16 * deger / toplam)
            p.setBrush(QColor(self.PALET[i % len(self.PALET)]))
            p.setPen(QColor("white"))
            p.drawPie(int(cx - cap / 2), int(cy - cap / 2), int(cap), int(cap), aci_baslangic, -aci_pay)
            aci_baslangic -= aci_pay

        ly = int(ust_pay + 8)
        for i, (etiket, deger) in enumerate(zip(self.etiketler, self.degerler)):
            renk = QColor(self.PALET[i % len(self.PALET)])
            p.setBrush(renk)
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(int(cap + 40), ly, 12, 12, 2, 2)
            p.setPen(QColor("#334155"))
            p.setFont(QFont("Segoe UI", 9))
            yuzde = (deger / toplam) * 100
            p.drawText(int(cap + 58), ly + 11, f"{etiket}  (%{yuzde:.0f})")
            ly += 22


class BenimPOSPlastik(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BenimPOS - Plastik Geri Dönüşüm Hammadde Yönetimi")
        self.resize(1220, 800)
        self.setStyleSheet("background-color: #f4f6f9;")

        # Sayfalama (Pagination) değişkenleri (Her sayfada 10 ürün/grup)
        self.sayfa_boyutu = 10
        self.urun_aktif_sayfa = 1
        self.grup_aktif_sayfa = 1
        self.grup_filtre_modu = "TUMU"
        self.musteri_aktif_sayfa = 1
        self.detay_aktif_sayfa = 1

        self.init_ui()
        self.load_products()
        try:
            vade_hatirlatici_mailleri_gonder()
        except:
            pass

    def evrak_fotografi_sec_ve_oku(self):
        dosya_yolu, _ = QFileDialog.getOpenFileName(
            self, "Çek/Senet Fotoğrafı Seç", "", "Resim Dosyaları (*.png *.jpg *.jpeg)"
        )
        if not dosya_yolu:
            return None
        QMessageBox.information(
            self, "Tarama Başladı",
            "Çek analiz ediliyor ve matematiksel çapraz doğrulama yapılıyor...",
        )
        analiz = UcretsizEvrakOkuyucu.evrak_analiz_et(dosya_yolu)
        if not analiz["guvenli_mi"]:
            QMessageBox.critical(
                self, "GÜVENLİK UYARISI: EŞLEŞME HATASI!",
                f"Sistem çekte bir hata tespit etti:\n\n"
                f"Okunan Rakam: {analiz['tutar_rakam']:,.2f} ₺\n"
                f"Sistemin Çekte Aradığı Yazı: {analiz['beklenen_yazi']}\n\n"
                f"Neden: {analiz['hata_mesaji']}\n\n"
                f"Lütfen evrakı elinize alıp tutarı manuel olarak kontrol edip sisteme girin!",
            )
        else:
            QMessageBox.information(
                self, "Çapraz Doğrulama Başarılı",
                f"✅ Çek %100 oranında doğrulandı.\n\n"
                f"Tutar: {analiz['tutar_rakam']:,.2f} ₺\n"
                f"Tarih: {analiz['vade_tarihi']}\n"
                f"Banka: {analiz['banka_adi']}",
            )
        return analiz

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        root_layout = QHBoxLayout(main_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ================= 1. SOL MENÜ (AÇILIR ALT MENÜLÜ & BÜYÜK HARFLİ) =================
        sidebar = QFrame()
        sidebar.setFixedWidth(230)
        sidebar.setStyleSheet("background-color: #ffffff; border-right: 1px solid #dee2e6;")
        self._golge_ekle(sidebar, bulaniklik=25, y_offset=0, opaklik=18)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 15)
        sidebar_layout.setSpacing(2)

        brand_banner = QFrame()
        brand_banner.setFixedHeight(64)
        brand_banner.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #0369a1, stop:1 #0284c7);
                border: none;
            }
        """)
        brand_layout = QHBoxLayout(brand_banner)
        brand_layout.setContentsMargins(16, 0, 12, 0)
        brand_layout.setSpacing(10)

        rozet = QLabel("♻")
        rozet.setFixedSize(34, 34)
        rozet.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rozet.setStyleSheet("background-color: rgba(255,255,255,0.18); border-radius: 17px; font-size: 17px;")
        brand_layout.addWidget(rozet)

        marka_metin = QVBoxLayout()
        marka_metin.setSpacing(0)
        ad_lbl = QLabel("PlastikPOS")
        ad_lbl.setStyleSheet("color: white; font-size: 15px; font-weight: bold;")
        alt_lbl = QLabel("PRO SÜRÜM")
        alt_lbl.setStyleSheet("color: rgba(255,255,255,0.75); font-size: 10px; font-weight: bold; letter-spacing: 1px;")
        marka_metin.addWidget(ad_lbl)
        marka_metin.addWidget(alt_lbl)
        brand_layout.addLayout(marka_metin)
        brand_layout.addStretch()
        sidebar_layout.addWidget(brand_banner)

        btn_home = QPushButton("  🏠 ANASAYFA")
        self.btn_nav_sales = QPushButton("  🛒 SATIŞ YAP")
        for b in (btn_home, self.btn_nav_sales):
            b.setFixedHeight(38)
            b.setStyleSheet("background: transparent; color: #555; text-align: left; padding-left: 18px; border: none; font-size: 13px; font-weight: 500;")
            sidebar_layout.addWidget(b)

        # ================= ÜRÜNLER AÇILIR MENÜSÜ =================
        self.products_menu_container = QWidget()
        prod_container_layout = QVBoxLayout(self.products_menu_container)
        prod_container_layout.setContentsMargins(0, 0, 0, 0)
        prod_container_layout.setSpacing(0)

        # ÜRÜNLER ANA BAŞLIĞI
        self.btn_main_products = QPushButton("  📦 ÜRÜNLER ▾")
        self.btn_main_products.setFixedHeight(38)
        self.btn_main_products.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_main_products.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #0088cc;
                font-weight: bold;
                font-size: 13px;
                text-align: left;
                padding-left: 18px;
                border: none;
            }
            QPushButton:hover { background-color: #f0f7ff; }
        """)
        self.btn_main_products.clicked.connect(self.toggle_products_submenu)
        prod_container_layout.addWidget(self.btn_main_products)

        # AÇILIR ALT MENÜ KUTUSU
        self.submenu_products = QFrame()
        self.submenu_products.setStyleSheet("background-color: #f8fafc; border-left: 3px solid #0088cc;")
        submenu_layout = QVBoxLayout(self.submenu_products)
        submenu_layout.setContentsMargins(0, 4, 0, 4)
        submenu_layout.setSpacing(2)

        # Alt Buton 1: ÜRÜN EKLE VE GÜNCELLE
        self.btn_sub_add = QPushButton("– ÜRÜN EKLE VE GÜNCELLE")
        self.btn_sub_add.setFixedHeight(34)
        self.btn_sub_add.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sub_add.setStyleSheet("background-color: #eef6ff; color: #0088cc; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
        self.btn_sub_add.clicked.connect(lambda: self.switch_page(0))
        submenu_layout.addWidget(self.btn_sub_add)

        # Alt Buton 2: ÜRÜN GRUBU EKLE
        self.btn_sub_groups = QPushButton("– ÜRÜN GRUBU EKLE")
        self.btn_sub_groups.setFixedHeight(34)
        self.btn_sub_groups.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sub_groups.setStyleSheet("background-color: transparent; color: #555555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
        self.btn_sub_groups.clicked.connect(lambda: self.switch_page(1))
        submenu_layout.addWidget(self.btn_sub_groups)

        prod_container_layout.addWidget(self.submenu_products)
        sidebar_layout.addWidget(self.products_menu_container)
        self.products_menu_container.installEventFilter(self)

        # ================= MÜŞTERİLER AÇILIR MENÜSÜ =================
        self.customers_menu_container = QWidget()
        cust_container_layout = QVBoxLayout(self.customers_menu_container)
        cust_container_layout.setContentsMargins(0, 0, 0, 0)
        cust_container_layout.setSpacing(0)

        # Müşteriler Ana Butonu
        self.btn_main_customers = QPushButton("  👤 MÜŞTERİLER ▾")
        self.btn_main_customers.setFixedHeight(38)
        self.btn_main_customers.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_main_customers.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #555555;
                font-weight: 500;
                font-size: 13px;
                text-align: left;
                padding-left: 18px;
                border: none;
            }
            QPushButton:hover {
                background-color: #f0f7ff;
                color: #0088cc;
            }
        """)
        self.btn_main_customers.clicked.connect(self.toggle_customers_submenu)
        cust_container_layout.addWidget(self.btn_main_customers)

        # Müşteriler Alt Paneli (– MÜŞTERİLER ve – MÜŞTERİ DETAY)
        self.submenu_customers = QFrame()
        self.submenu_customers.setStyleSheet("background-color: #f8fafc; border-left: 3px solid #0088cc;")
        self.submenu_customers.hide()  # Başlangıçta kapalı
        cust_sub_layout = QVBoxLayout(self.submenu_customers)
        cust_sub_layout.setContentsMargins(0, 4, 0, 4)
        cust_sub_layout.setSpacing(2)

        # 1. Alt Buton: – MÜŞTERİLER
        self.btn_sub_cust_list = QPushButton("  – MÜŞTERİLER")
        self.btn_sub_cust_list.setFixedHeight(34)
        self.btn_sub_cust_list.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sub_cust_list.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #0088cc;
                font-weight: bold;
                text-align: left;
                padding-left: 20px;
                border: none;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #e2efff; }
        """)
        self.btn_sub_cust_list.clicked.connect(lambda: self.on_customer_menu_click("LIST"))
        cust_sub_layout.addWidget(self.btn_sub_cust_list)

        # 2. Alt Buton: – MÜŞTERİ DETAY
        self.btn_sub_cust_detail = QPushButton("  – MÜŞTERİ DETAY")
        self.btn_sub_cust_detail.setFixedHeight(34)
        self.btn_sub_cust_detail.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sub_cust_detail.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #555555;
                font-weight: bold;
                text-align: left;
                padding-left: 20px;
                border: none;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #e2efff; color: #0088cc; }
        """)
        self.btn_sub_cust_detail.clicked.connect(lambda: self.on_customer_menu_click("DETAIL"))
        cust_sub_layout.addWidget(self.btn_sub_cust_detail)

        cust_container_layout.addWidget(self.submenu_customers)
        sidebar_layout.addWidget(self.customers_menu_container)
        self.customers_menu_container.installEventFilter(self)

        self.reports_menu_container = QWidget()
        reports_container_layout = QVBoxLayout(self.reports_menu_container)
        reports_container_layout.setContentsMargins(0, 0, 0, 0)
        reports_container_layout.setSpacing(0)
        self.btn_main_reports = QPushButton("  📊 RAPORLAR ▾")
        self.btn_main_reports.setFixedHeight(38)
        self.btn_main_reports.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_main_reports.setStyleSheet("""
            QPushButton {
                background-color: transparent; color: #555555; font-weight: 500;
                font-size: 13px; text-align: left; padding-left: 18px; border: none;
            }
            QPushButton:hover { background-color: #f0f7ff; color: #0088cc; }
        """)
        self.btn_main_reports.clicked.connect(self.toggle_reports_submenu)
        reports_container_layout.addWidget(self.btn_main_reports)
        self.submenu_reports = QFrame()
        self.submenu_reports.setStyleSheet("background-color: #f8fafc; border-left: 3px solid #0088cc;")
        self.submenu_reports.hide()
        reports_sub_layout = QVBoxLayout(self.submenu_reports)
        reports_sub_layout.setContentsMargins(0, 4, 0, 4)
        reports_sub_layout.setSpacing(2)
        rapor_altlari = [
            ("– GÜNLÜK RAPOR", 0),
            ("– TARİHSEL RAPOR", 1),
            ("– ÜRÜNSEL RAPOR", 2),
            ("– GRUPSAL RAPOR", 3),
            ("– STOK HAREKET", 4),
            ("– PERSONEL HAREKET", 5),
        ]
        for etiket, idx in rapor_altlari:
            btn_r = QPushButton(f"  {etiket}")
            btn_r.setFixedHeight(34)
            btn_r.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_r.setStyleSheet("background: transparent; color: #555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 12px;")
            btn_r.clicked.connect(lambda _, i=idx: self.rapor_sayfasini_ac(i))
            reports_sub_layout.addWidget(btn_r)
        self.btn_nav_reports = QPushButton("  – PERSONEL KARTLARI")
        self.btn_nav_reports.setFixedHeight(34)
        self.btn_nav_reports.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_nav_reports.setStyleSheet("background: transparent; color: #555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 12px;")
        reports_sub_layout.addWidget(self.btn_nav_reports)
        self.btn_rapor_supheli = QPushButton("  – ŞÜPHELİ İŞLEMLER")
        self.btn_rapor_supheli.setFixedHeight(34)
        self.btn_rapor_supheli.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_rapor_supheli.setStyleSheet("""
            QPushButton {
                text-align: left; padding-left: 20px; height: 32px;
                color: #ef4444; font-weight: bold; border: none; background: transparent; font-size: 12px;
            }
            QPushButton:hover { background-color: #fee2e2; border-radius: 6px; }
        """)
        self.btn_rapor_supheli.clicked.connect(self.supheli_islemler_sayfasini_ac)
        reports_sub_layout.addWidget(self.btn_rapor_supheli)
        reports_container_layout.addWidget(self.submenu_reports)
        sidebar_layout.addWidget(self.reports_menu_container)
        self.reports_menu_container.installEventFilter(self)

        extra_menus = ["💼 ÇEK & SENET", "💬 WHATSAPP SATIŞ", "⚙ AYARLAR"]
        for m in extra_menus:
            btn = QPushButton(f"  {m}")
            btn.setFixedHeight(38)
            btn.setStyleSheet("background: transparent; color: #555; text-align: left; padding-left: 18px; border: none; font-size: 13px; font-weight: 500;")
            if "WHATSAPP SATIŞ" in m:
                self.btn_nav_whatsapp = btn
            elif "ÇEK" in m:
                self.btn_nav_finans = btn
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()
        root_layout.addWidget(sidebar)

        # ================= 2. SAĞ İÇERİK (SAYFALAR) =================
        self.stack = QStackedWidget()
        root_layout.addWidget(self.stack)

        self.page_products = self.create_products_page()
        self.stack.addWidget(self.page_products)

        self.page_groups = self.create_groups_page()
        self.stack.addWidget(self.page_groups)

        self.page_customers = self.create_customers_page()
        self.stack.addWidget(self.page_customers)

        self.page_customer_details = self.create_customer_details_page()
        self.stack.addWidget(self.page_customer_details)

        self.sales_page = self.create_sales_page()
        self.stack.addWidget(self.sales_page)
        self.btn_nav_sales.clicked.connect(lambda: self.stack.setCurrentWidget(self.sales_page))

        self.page_whatsapp = self.create_whatsapp_sales_page()
        self.stack.addWidget(self.page_whatsapp)
        self.btn_nav_whatsapp.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_whatsapp))

        self.page_personnel = self.create_personnel_page()
        self.stack.addWidget(self.page_personnel)
        self.btn_nav_reports.clicked.connect(self.personel_kartlari_ac)

        self.page_reports = self.create_reports_hub_page()
        self.stack.addWidget(self.page_reports)

        self.sayfa_supheli_islemler = self.create_supheli_islemler_page()
        self.stack.addWidget(self.sayfa_supheli_islemler)

        self.page_finance = self.create_finance_page()
        self.stack.addWidget(self.page_finance)
        self.btn_nav_finans.clicked.connect(self.finans_sayfasini_ac)

    def personel_kartlari_ac(self):
        self.stack.setCurrentWidget(self.page_personnel)
        self.load_personnel_data()

    def personel_rapor_ac(self):
        self.rapor_sayfasini_ac(0)

    def rapor_sayfasini_ac(self, idx):
        self.stack.setCurrentWidget(self.page_reports)
        self.report_stack.setCurrentIndex(idx)
        yenile = (
            self.gunluk_rapor_yenile, self.tarihsel_rapor_yenile,
            self.urunsel_rapor_yenile, self.grupsal_rapor_yenile,
            self.stok_hareket_rapor_yenile, self.personel_hareket_rapor_yenile,
        )
        if 0 <= idx < len(yenile):
            yenile[idx]()

    def finans_sayfasini_ac(self):
        self.stack.setCurrentWidget(self.page_finance)
        self.finans_tablosunu_guncelle()

    def toggle_reports_submenu(self):
        if self.submenu_reports.isVisible():
            self.submenu_reports.hide()
            self.btn_main_reports.setText("  📊 RAPORLAR ▸")
        else:
            self.submenu_reports.show()
            self.btn_main_reports.setText("  📊 RAPORLAR ▾")

    def eventFilter(self, obj, event):
        # Ürünler menüsü hover açılışı
        if obj == getattr(self, "products_menu_container", None) and event.type() == QEvent.Type.Enter:
            self.submenu_products.show()
            self.btn_main_products.setText("  📦 ÜRÜNLER ▾")
        # Müşteriler menüsü hover açılışı
        elif obj == getattr(self, "customers_menu_container", None) and event.type() == QEvent.Type.Enter:
            self.submenu_customers.show()
            self.btn_main_customers.setText("  👤 MÜŞTERİLER ▾")
        elif obj == getattr(self, "reports_menu_container", None) and event.type() == QEvent.Type.Enter:
            self.submenu_reports.show()
            self.btn_main_reports.setText("  📊 RAPORLAR ▾")
        return super().eventFilter(obj, event)

    def toggle_products_submenu(self):
        if self.submenu_products.isVisible():
            self.submenu_products.hide()
            self.btn_main_products.setText("  📦 ÜRÜNLER ▸")
        else:
            self.submenu_products.show()
            self.btn_main_products.setText("  📦 ÜRÜNLER ▾")

    def toggle_customers_submenu(self):
        """Müşteriler başlığına tıklanınca açılıp kapanmasını sağlar"""
        if self.submenu_customers.isVisible():
            self.submenu_customers.hide()
            self.btn_main_customers.setText("  👤 MÜŞTERİLER ▸")
        else:
            self.submenu_customers.show()
            self.btn_main_customers.setText("  👤 MÜŞTERİLER ▾")

    def on_customer_menu_click(self, target):
        """Müşteri alt butonlarına tıklandığında vurgu ve sayfa yönlendirmesi"""
        aktif = "background-color: #e0f2fe; color: #0088cc; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 12px;"
        pasif = "background-color: transparent; color: #555555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 12px;"
        if target == "LIST":
            self.btn_sub_cust_list.setStyleSheet(aktif)
            self.btn_sub_cust_detail.setStyleSheet(pasif)
            self.stack.setCurrentIndex(2)
            self.load_customers()
        else:
            self.btn_sub_cust_detail.setStyleSheet(aktif)
            self.btn_sub_cust_list.setStyleSheet(pasif)
            self.stack.setCurrentIndex(3)
            self.load_customer_details()

    # ================= SAYFA: SATIŞ YAP (BENİMPOS BİREBİR) =================
    def create_sales_page(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(15, 12, 15, 12)
        page_layout.setSpacing(10)

        self.sepet = []
        self.secili_musteri = None

        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        v_left = QVBoxLayout()
        h_search_bar = QHBoxLayout()
        h_search_bar.setSpacing(6)

        self.cmb_fiyat_tipi = QComboBox()
        self.cmb_fiyat_tipi.addItems(["Fiyat 1", "Fiyat 2"])
        self.cmb_fiyat_tipi.setFixedHeight(38)
        self.cmb_fiyat_tipi.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-weight: bold; background: white;")
        h_search_bar.addWidget(self.cmb_fiyat_tipi)

        self.txt_barkod_ara = BuyukHarfKutusu("Ürün barkodunu okutunuz veya çeşit yazınız (örn: beyaz pom)...")
        self.txt_barkod_ara.setFixedHeight(38)
        self.txt_barkod_ara.setStyleSheet("border: 2px solid #0ea5e9; border-radius: 4px; padding: 0 10px; font-size: 13.5px; font-weight: bold; background: white;")
        self.txt_barkod_ara.textChanged.connect(self.akilli_urun_onerileri_goster)
        self.txt_barkod_ara.returnPressed.connect(self.akilli_arama_sepete_ekle)
        h_search_bar.addWidget(self.txt_barkod_ara)

        btn_ara = QPushButton("🔍 Ara")
        btn_ara.setFixedSize(70, 38)
        btn_ara.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ara.setStyleSheet("background-color: #0ea5e9; color: white; font-weight: bold; border-radius: 4px; border: none; font-size: 13px;")
        btn_ara.clicked.connect(self.akilli_arama_sepete_ekle)
        h_search_bar.addWidget(btn_ara)

        btn_fiyat = QPushButton("Fiyat Gör")
        btn_fiyat.setFixedSize(80, 38)
        btn_fiyat.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_fiyat.setStyleSheet("background-color: #10b981; color: white; font-weight: bold; border-radius: 4px; border: none; font-size: 13px;")
        btn_fiyat.clicked.connect(self.fiyat_gor_popup)
        h_search_bar.addWidget(btn_fiyat)

        btn_yazdir = QPushButton("🖨 Yazdır ▾")
        btn_yazdir.setFixedSize(95, 38)
        btn_yazdir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yazdir.setStyleSheet("background-color: #f59e0b; color: white; font-weight: bold; border-radius: 4px; border: none; font-size: 13px;")
        btn_yazdir.clicked.connect(self.satisi_yazdir_diyalog)
        h_search_bar.addWidget(btn_yazdir)

        v_left.addLayout(h_search_bar)
        top_row.addLayout(v_left, stretch=6)

        h_totals = QHBoxLayout()
        h_totals.setSpacing(8)

        v_od = QVBoxLayout()
        v_od.addWidget(QLabel("Ödenen", styleSheet="color: #64748b; font-size: 11px; font-weight: bold;"))
        self.txt_odenen_tutar = SadeceSayiKutusu(fiyat_modu=False)
        self.txt_odenen_tutar.setFixedHeight(38)
        self.txt_odenen_tutar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.txt_odenen_tutar.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; font-size: 20px; font-weight: bold; color: #1e293b; background: white;")
        self.txt_odenen_tutar.textChanged.connect(self.para_ustu_hesapla)
        v_od.addWidget(self.txt_odenen_tutar)
        h_totals.addLayout(v_od)

        v_tut = QVBoxLayout()
        v_tut.addWidget(QLabel("Tutar", styleSheet="color: #64748b; font-size: 11px; font-weight: bold;"))
        self.lbl_toplam_tutar = QLabel("0.00")
        self.lbl_toplam_tutar.setFixedHeight(38)
        self.lbl_toplam_tutar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_toplam_tutar.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; font-size: 22px; font-weight: bold; color: #dc2626; background: white;")
        v_tut.addWidget(self.lbl_toplam_tutar)
        h_totals.addLayout(v_tut)

        v_pu = QVBoxLayout()
        v_pu.addWidget(QLabel("Para Üstü", styleSheet="color: #64748b; font-size: 11px; font-weight: bold;"))
        self.lbl_para_ustu = QLabel("0.00")
        self.lbl_para_ustu.setFixedHeight(38)
        self.lbl_para_ustu.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_para_ustu.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; font-size: 22px; font-weight: bold; color: #16a34a; background: white;")
        v_pu.addWidget(self.lbl_para_ustu)
        h_totals.addLayout(v_pu)

        top_row.addLayout(h_totals, stretch=4)
        page_layout.addLayout(top_row)

        mid_row = QHBoxLayout()
        mid_row.setSpacing(14)

        sol_panel = QVBoxLayout()

        h_m_bar = QHBoxLayout()
        self.btn_tab_m1 = QPushButton("Müşteri (0.00 ₺)")
        self.btn_tab_m1.setFixedHeight(32)
        self.btn_tab_m1.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 4px; border: none; padding: 0 16px;")
        h_m_bar.addWidget(self.btn_tab_m1)
        h_m_bar.addStretch()

        btn_temizle = QPushButton("🗑 Sepeti Temizle")
        btn_temizle.setFixedHeight(30)
        btn_temizle.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_temizle.setStyleSheet("color: #dc2626; border: none; font-weight: bold; font-size: 12px;")
        btn_temizle.clicked.connect(self.sepeti_temizle)
        h_m_bar.addWidget(btn_temizle)
        sol_panel.addLayout(h_m_bar)

        self.table_sepet = QTableWidget()
        self.table_sepet.setColumnCount(5)
        self.table_sepet.setHorizontalHeaderLabels(["Ürün", "Miktar & Birim", "Birim Fiyat", "Tutar", "Sil"])
        self.table_sepet.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_sepet.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.table_sepet.setColumnWidth(1, 230)
        self.table_sepet.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table_sepet.setColumnWidth(2, 130)
        self.table_sepet.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.table_sepet.setColumnWidth(3, 110)
        self.table_sepet.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table_sepet.setColumnWidth(4, 50)
        self.table_sepet.verticalHeader().setVisible(False)
        self.table_sepet.setStyleSheet("""
            QTableWidget { border: 1px solid #cbd5e1; background: white; font-size: 13.5px; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 36px; border-bottom: 2px solid #cbd5e1; }
        """)
        sol_panel.addWidget(self.table_sepet)
        mid_row.addLayout(sol_panel, stretch=6)

        sag_panel = QVBoxLayout()
        sag_panel.setSpacing(8)

        h_cust_bar = QHBoxLayout()
        self.lbl_secili_musteri = QLabel("Müşteri Seçilmedi")
        self.lbl_secili_musteri.setFixedHeight(36)
        self.lbl_secili_musteri.setStyleSheet("background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 10px; font-weight: bold; color: #334155;")
        h_cust_bar.addWidget(self.lbl_secili_musteri)

        btn_sec = QPushButton("  + Seç  ")
        btn_sec.setFixedHeight(36)
        btn_sec.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_sec.setStyleSheet("background-color: #0ea5e9; color: white; font-weight: bold; border-radius: 4px; border: none; padding: 0 14px;")
        btn_sec.clicked.connect(self.popup_musteri_sec_ac)
        h_cust_bar.addWidget(btn_sec)
        sag_panel.addLayout(h_cust_bar)

        h_nakit_sayilar = QHBoxLayout()
        h_nakit_sayilar.setSpacing(4)
        for val in ["20", "50", "100", "200", "+20", "-20"]:
            btn_n = QPushButton(val)
            btn_n.setFixedHeight(30)
            btn_n.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_n.setStyleSheet("background-color: #f0f9ff; color: #0284c7; border: 1px solid #bae6fd; border-radius: 4px; font-weight: bold; font-size: 11.5px;")
            btn_n.clicked.connect(lambda _, v=val: self.hizli_nakit_ekle(v))
            h_nakit_sayilar.addWidget(btn_n)
        sag_panel.addLayout(h_nakit_sayilar)

        h_pay_buttons = QHBoxLayout()
        h_pay_buttons.setSpacing(6)

        self.btn_pay_cash = QPushButton("₺ (F8)\nNAKİT")
        self.btn_pay_cash.setFixedHeight(65)
        self.btn_pay_cash.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay_cash.setStyleSheet("""
            QPushButton { background-color: #22c55e; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background-color: #16a34a; }
        """)
        self.btn_pay_cash.clicked.connect(lambda: self.satisi_tamamla("NAKİT"))
        h_pay_buttons.addWidget(self.btn_pay_cash)

        self.btn_pay_pos = QPushButton("💳 (F9)\nPOS")
        self.btn_pay_pos.setFixedHeight(65)
        self.btn_pay_pos.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay_pos.setStyleSheet("""
            QPushButton { background-color: #06b6d4; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background-color: #0891b2; }
        """)
        self.btn_pay_pos.clicked.connect(lambda: self.satisi_tamamla("POS / KREDİ KARTI"))
        h_pay_buttons.addWidget(self.btn_pay_pos)

        self.btn_pay_veresiye = QPushButton("📑 (F10)\nAÇIK HESAP")
        self.btn_pay_veresiye.setFixedHeight(65)
        self.btn_pay_veresiye.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay_veresiye.setStyleSheet("""
            QPushButton { background-color: #f59e0b; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background-color: #d97706; }
        """)
        self.btn_pay_veresiye.clicked.connect(lambda: self.satisi_tamamla("AÇIK HESAP (VERESİYE)"))
        h_pay_buttons.addWidget(self.btn_pay_veresiye)

        self.btn_pay_parcali = QPushButton("🔀\nPARÇALI")
        self.btn_pay_parcali.setFixedHeight(65)
        self.btn_pay_parcali.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay_parcali.setStyleSheet("""
            QPushButton { background-color: #2563eb; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background-color: #1d4ed8; }
        """)
        self.btn_pay_parcali.clicked.connect(lambda: self.satisi_tamamla("PARÇALI ÖDEME"))
        h_pay_buttons.addWidget(self.btn_pay_parcali)

        self.btn_pay_diger = QPushButton("+\nDİĞER")
        self.btn_pay_diger.setFixedHeight(65)
        self.btn_pay_diger.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay_diger.setStyleSheet("""
            QPushButton { background-color: #ef4444; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background-color: #dc2626; }
        """)
        self.btn_pay_diger.clicked.connect(lambda: self.satisi_tamamla("DİĞER"))
        h_pay_buttons.addWidget(self.btn_pay_diger)

        sag_panel.addLayout(h_pay_buttons)

        self.btn_iptal_son_islem = QPushButton("↩️ SON İŞLEMİ İPTAL ET (GERİ AL)")
        self.btn_iptal_son_islem.setFixedHeight(45)
        self.btn_iptal_son_islem.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_iptal_son_islem.setStyleSheet("""
            QPushButton {
                background-color: #ef4444;
                color: white;
                font-weight: bold;
                font-size: 13px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #dc2626;
            }
        """)
        self.btn_iptal_son_islem.clicked.connect(self.en_son_satis_islem_iptal_et)
        sag_panel.addWidget(self.btn_iptal_son_islem)

        sekme_scroll = QScrollArea()
        sekme_scroll.setMinimumHeight(48)
        sekme_scroll.setMaximumHeight(136)
        sekme_scroll.setWidgetResizable(True)
        sekme_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        sekme_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        sekme_scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        sekme_widget = QWidget()
        self.hizli_kategori_bar = QVBoxLayout(sekme_widget)
        self.hizli_kategori_bar.setContentsMargins(0, 4, 0, 4)
        self.hizli_kategori_bar.setSpacing(6)
        self.layout_grup_butonlari = self.hizli_kategori_bar
        sekme_scroll.setWidget(sekme_widget)
        sag_panel.addWidget(sekme_scroll)

        scroll_urunler = QScrollArea()
        scroll_urunler.setWidgetResizable(True)
        scroll_urunler.setStyleSheet("border: 1px solid #e2e8f0; background: #f8fafc;")
        self.grid_urun_alani = QWidget()
        self.grid_urun_layout = QGridLayout(self.grid_urun_alani)
        self.grid_urun_layout.setSpacing(8)
        self.grid_urun_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        scroll_urunler.setWidget(self.grid_urun_alani)
        sag_panel.addWidget(scroll_urunler)

        mid_row.addLayout(sag_panel, stretch=6)
        page_layout.addLayout(mid_row)

        self.load_product_groups()
        self.urunleri_kategoriye_gore_filtrele("ANA")
        return page

    def popup_musteri_sec_ac(self):
        dlg = MusteriSecDialog(self)
        if dlg.exec() and dlg.secilen_musteri:
            self.secili_musteri = dlg.secilen_musteri
            m_ad = self.secili_musteri["name"]
            self.lbl_secili_musteri.setText(f"👤 {m_ad}")
            toplam = sum(item["total"] for item in self.sepet)
            self.btn_tab_m1.setText(f"👤 {m_ad} ({toplam:,.2f} ₺)")

    def hizli_nakit_ekle(self, val):
        mevcut = self.txt_odenen_tutar.sayi_al()
        if val == "+20":
            mevcut += 20
        elif val == "-20":
            mevcut = max(0, mevcut - 20)
        else:
            mevcut = float(val)
        if mevcut <= 0:
            self.txt_odenen_tutar.clear()
        else:
            yazi = f"{int(mevcut):,}".replace(",", ".") if mevcut == int(mevcut) else f"{mevcut:.2f}".replace(".", ",")
            self.txt_odenen_tutar.otomatik_formatla(yazi)

    def para_ustu_hesapla(self):
        odenen = self.txt_odenen_tutar.sayi_al() if hasattr(self, "txt_odenen_tutar") else 0.0
        toplam = sum(item["total"] for item in self.sepet)
        para_ustu = max(0.0, odenen - toplam)
        self.lbl_para_ustu.setText(f"{para_ustu:,.2f}")

    def sepet_guncelle(self):
        self.table_sepet.setRowCount(len(self.sepet))
        toplam = 0.0

        for r_idx, item in enumerate(self.sepet):
            self.table_sepet.setRowHeight(r_idx, 44)

            it_name = QTableWidgetItem(item["name"])
            it_name.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_name.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_sepet.setItem(r_idx, 0, it_name)

            w_miktar_alani = QWidget()
            l_miktar = QHBoxLayout(w_miktar_alani)
            l_miktar.setContentsMargins(4, 2, 4, 2)
            l_miktar.setSpacing(6)

            txt_qty = SadeceSayiKutusu(fiyat_modu=False)
            txt_qty.setFixedHeight(34)
            txt_qty.setStyleSheet("""
                QLineEdit {
                    border: 1px solid #cbd5e1;
                    border-radius: 4px;
                    padding: 0 8px;
                    font-size: 13px;
                    font-weight: bold;
                    background: white;
                }
                QLineEdit:focus { border: 2px solid #0284c7; }
            """)
            if item["qty"] > 0:
                txt_qty.setText(f"{int(item['qty']):,}".replace(",", ".") if item["qty"] == int(item["qty"]) else f"{item['qty']:.2f}".replace(".", ","))
            l_miktar.addWidget(txt_qty, stretch=1)

            cmb_unit = QComboBox()
            cmb_unit.setFixedHeight(34)
            cmb_unit.setFixedWidth(95)
            cmb_unit.addItems([
                "KG", "TON", "BİG BAG", "METRE", "ADET",
                "RULO", "PAKET", "KOLİ", "LİTRE", "PALET"
            ])
            secili_birim = item.get("unit", "KG")
            if cmb_unit.findText(secili_birim) < 0:
                cmb_unit.addItem(secili_birim)
            cmb_unit.setCurrentText(secili_birim)
            cmb_unit.setStyleSheet("""
                QComboBox {
                    border: 1px solid #cbd5e1;
                    border-radius: 4px;
                    padding: 2px 6px;
                    font-size: 12px;
                    font-weight: bold;
                    background: #f8fafc;
                    color: #1e293b;
                }
                QComboBox::drop-down {
                    border: none;
                    width: 18px;
                }
                QComboBox QAbstractItemView {
                    border: 1px solid #cbd5e1;
                    background-color: #ffffff;
                    selection-background-color: #e0f2fe;
                    selection-color: #0369a1;
                    min-width: 100px;
                    padding: 4px;
                    font-size: 12px;
                    font-weight: bold;
                }
                QComboBox QAbstractItemView::item {
                    min-height: 26px;
                    padding-left: 8px;
                }
            """)
            cmb_unit.currentTextChanged.connect(lambda val, idx=r_idx: self.birim_degistir(idx, val))
            l_miktar.addWidget(cmb_unit)

            txt_qty.textChanged.connect(lambda _, idx=r_idx, w=txt_qty: self.sepet_satir_hesapla(idx, miktar_w=w))
            self.table_sepet.setCellWidget(r_idx, 1, w_miktar_alani)

            txt_price = SadeceSayiKutusu(fiyat_modu=True)
            txt_price.setFixedHeight(34)
            txt_price.setAlignment(Qt.AlignmentFlag.AlignRight)
            txt_price.setStyleSheet("""
                QLineEdit {
                    border: 1px solid #cbd5e1;
                    border-radius: 4px;
                    padding: 0 8px;
                    font-size: 13.5px;
                    font-weight: bold;
                    color: #0369a1;
                    background: #ffffff;
                }
                QLineEdit:focus {
                    border: 2px solid #0284c7;
                    background: #f0f9ff;
                }
            """)
            if item["price"] > 0:
                fiyat_str = f"{item['price']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                txt_price.setText(fiyat_str)
            txt_price.textChanged.connect(lambda _, idx=r_idx, w=txt_price: self.sepet_satir_hesapla(idx, fiyat_w=w))
            self.table_sepet.setCellWidget(r_idx, 2, txt_price)

            it_tot = QTableWidgetItem(f"{item['total']:,.2f} ₺")
            font_tutar = QFont()
            font_tutar.setPointSize(11)
            font_tutar.setWeight(QFont.Weight.Medium)
            it_tot.setFont(font_tutar)
            it_tot.setForeground(QColor("#1e293b"))
            it_tot.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it_tot.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_sepet.setItem(r_idx, 3, it_tot)

            btn_del = QPushButton("🗑")
            btn_del.setFixedSize(32, 32)
            btn_del.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_del.setToolTip("Sil")
            btn_del.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: #ef4444;
                    border: 1px solid transparent;
                    border-radius: 6px;
                    font-size: 16px;
                    padding: 0;
                }
                QPushButton:hover {
                    background-color: #fee2e2;
                    border: 1px solid #fca5a5;
                    color: #dc2626;
                }
            """)
            btn_del.clicked.connect(lambda _, idx=r_idx: self.sepetten_cikar(idx))

            w_del = QWidget()
            l_del = QHBoxLayout(w_del)
            l_del.setContentsMargins(0, 0, 0, 0)
            l_del.setAlignment(Qt.AlignmentFlag.AlignCenter)
            l_del.addWidget(btn_del)
            self.table_sepet.setCellWidget(r_idx, 4, w_del)

            toplam += item["total"]

        self.lbl_toplam_tutar.setText(f"{toplam:,.2f}")
        if self.secili_musteri:
            self.btn_tab_m1.setText(f"👤 {self.secili_musteri['name']} ({toplam:,.2f} ₺)")
        else:
            self.btn_tab_m1.setText(f"Müşteri ({toplam:,.2f} ₺)")
        self.para_ustu_hesapla()

    def birim_degistir(self, idx, yeni_birim):
        if idx < len(self.sepet):
            self.sepet[idx]["unit"] = yeni_birim

    def sepet_satir_hesapla(self, idx, miktar_w=None, fiyat_w=None):
        if idx >= len(self.sepet):
            return

        w_miktar_container = self.table_sepet.cellWidget(idx, 1)
        w_qty = w_miktar_container.findChild(SadeceSayiKutusu) if w_miktar_container else None
        w_pr = self.table_sepet.cellWidget(idx, 2)

        qty_val = w_qty.sayi_al() if w_qty else self.sepet[idx]["qty"]
        pr_val = w_pr.sayi_al() if isinstance(w_pr, SadeceSayiKutusu) else self.sepet[idx]["price"]

        self.sepet[idx]["qty"] = qty_val
        self.sepet[idx]["price"] = pr_val
        self.sepet[idx]["total"] = qty_val * pr_val

        it_tot = self.table_sepet.item(idx, 3)
        if it_tot:
            it_tot.setText(f"{self.sepet[idx]['total']:,.2f} ₺")

        toplam = sum(it["total"] for it in self.sepet)
        self.lbl_toplam_tutar.setText(f"{toplam:,.2f}")
        if self.secili_musteri:
            self.btn_tab_m1.setText(f"👤 {self.secili_musteri['name']} ({toplam:,.2f} ₺)")
        else:
            self.btn_tab_m1.setText(f"Müşteri ({toplam:,.2f} ₺)")
        self.para_ustu_hesapla()

    def akilli_urun_onerileri_goster(self, text):
        if len(text.strip()) < 2:
            if hasattr(self, "_urun_model"):
                self._urun_model.setStringList([])
            return

        arama_norm = turkce_toleransli_metin(text.strip())
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT name FROM products")
        all_prods = [r[0] for r in c.fetchall()]
        conn.close()

        kelimeler = arama_norm.split()
        eslesenler = [p_name for p_name in all_prods if all(k in turkce_toleransli_metin(p_name) for k in kelimeler)]

        if not hasattr(self, "_urun_completer"):
            self._urun_model = QStringListModel(self)
            self._urun_completer = QCompleter(self._urun_model, self.txt_barkod_ara)
            self._urun_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            self._urun_completer.setCompletionMode(QCompleter.CompletionMode.UnfilteredPopupCompletion)
            self._urun_completer.setMaxVisibleItems(12)
            self._urun_completer.activated.connect(lambda _t="": self.akilli_arama_sepete_ekle())
            self.txt_barkod_ara.setCompleter(self._urun_completer)
        self._urun_model.setStringList(eslesenler)
        if eslesenler:
            self._urun_completer.complete()

    def akilli_arama_sepete_ekle(self):
        txt = self.txt_barkod_ara.text().strip()
        if not txt:
            return

        txt_norm = turkce_toleransli_metin(txt)
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name, sell_price FROM products")
        rows = c.fetchall()
        conn.close()

        kelimeler = txt_norm.split()
        bulunan = None
        for pid, name, price in rows:
            p_norm = turkce_toleransli_metin(name)
            if all(k in p_norm for k in kelimeler):
                bulunan = (pid, name, price)
                break

        if bulunan:
            self.hizli_urun_sepete_ekle(bulunan[0], bulunan[1], float(bulunan[2] or 0.0))
            self.txt_barkod_ara.clear()
        else:
            QMessageBox.warning(self, "Bulunamadı", f"'{txt}' ürünü bulunamadı!")

    def satisi_yazdir_diyalog(self):
        if not self.sepet:
            QMessageBox.warning(self, "Yazdırılamaz", "Yazdırmak için sepete en az bir ürün ekleyiniz!")
            return

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        print_dialog = QPrintDialog(printer, self)
        print_dialog.setWindowTitle("Yazıcı Seç ve Fiş Yazdır")

        if print_dialog.exec() == QDialog.DialogCode.Accepted:
            doc = QTextDocument()
            toplam = sum(i["total"] for i in self.sepet)
            musteri_ad = self.secili_musteri["name"] if self.secili_musteri else "PERAKENDE SATIŞ"

            html = f"""
            <div style="font-family: Arial, sans-serif; font-size: 11pt;">
                <h3 align="center" style="margin:0;">PLASTİK GERİ DÖNÜŞÜM POS</h3>
                <p align="center" style="margin:2px 0 10px 0; font-size: 9pt;">SATIŞ BİLGİ FİŞİ</p>
                <hr>
                <table width="100%" style="font-size: 10pt;">
                    <tr><td><b>Müşteri:</b> {musteri_ad}</td></tr>
                    <tr><td><b>Ödeme:</b> Peşin / Cari</td></tr>
                </table>
                <hr>
                <table width="100%" border="0" cellspacing="4" style="font-size: 10pt;">
                    <tr style="border-bottom: 1px solid #000;">
                        <th align="left">Ürün</th>
                        <th align="center">Miktar</th>
                        <th align="right">Fiyat</th>
                        <th align="right">Tutar</th>
                    </tr>
            """
            for it in self.sepet:
                html += f"""
                    <tr>
                        <td>{it['name']}</td>
                        <td align="center">{it['qty']:.1f} KG</td>
                        <td align="right">{it['price']:,.2f} ₺</td>
                        <td align="right"><b>{it['total']:,.2f} ₺</b></td>
                    </tr>
                """

            html += f"""
                </table>
                <hr>
                <h2 align="right" style="margin:5px 0;">TOPLAM: {toplam:,.2f} ₺</h2>
                <hr>
                <p align="center" style="font-size: 8pt; margin-top: 10px;">İyi çalışmalar dileriz.</p>
            </div>
            """
            doc.setHtml(html)
            doc.print(printer)
            QMessageBox.information(self, "Başarılı", "Satış fişi seçilen yazıcıya başarıyla gönderildi!")

    def fiyat_gor_popup(self):
        txt = self.txt_barkod_ara.text().strip()
        if not txt:
            QMessageBox.information(self, "Fiyat Gör", "Lütfen bir ürün adı yazın.")
            return
        txt_norm = turkce_toleransli_metin(txt)
        kelimeler = txt_norm.split()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT name, sell_price, stock_kg FROM products")
        rows = c.fetchall()
        conn.close()
        row = None
        for name, price, stock in rows:
            if all(k in turkce_toleransli_metin(name) for k in kelimeler):
                row = (name, price, stock)
                break
        if row:
            QMessageBox.information(self, "Fiyat Bilgisi", f"Ürün: {row[0]}\nSatış Fiyatı: {row[1]:,.2f} ₺\nMevcut Stok: {row[2]:,.2f} KG")
        else:
            QMessageBox.warning(self, "Bulunamadı", "Ürün bulunamadı!")

    def sepetten_cikar(self, idx):
        if idx < len(self.sepet):
            self.sepet.pop(idx)
            self.sepet_guncelle()

    def sepeti_temizle(self):
        self.sepet.clear()
        self.txt_odenen_tutar.setText("")
        self.secili_musteri = None
        self.lbl_secili_musteri.setText("Müşteri Seçilmedi")
        self.btn_tab_m1.setText("Müşteri (0.00 ₺)")
        self.sepet_guncelle()

    def hizli_urun_sepete_ekle(self, p_id, p_name, p_price):
        p_price = float(p_price or 0.0)
        son_fiyat = guncel_fiyati_getir(p_name)
        if son_fiyat is not None:
            p_price = float(son_fiyat)
        for item in self.sepet:
            if item["id"] == p_id:
                item["qty"] += 1.0
                item["total"] = item["qty"] * item["price"]
                self.sepet_guncelle()
                return
        self.sepet.append({
            "id": p_id,
            "name": p_name,
            "qty": 1.0,
            "price": p_price,
            "total": 1.0 * p_price,
            "unit": "KG"
        })
        self.sepet_guncelle()

    def barkod_ile_sepete_ekle(self):
        txt = buyuk_harf(self.txt_barkod_ara.text().strip())
        if not txt:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name, sell_price FROM products WHERE name LIKE ? LIMIT 1", (f"%{txt}%",))
        row = c.fetchone()
        conn.close()
        if row:
            self.hizli_urun_sepete_ekle(row[0], row[1], float(row[2] or 0.0))
            self.txt_barkod_ara.clear()
        else:
            QMessageBox.warning(self, "Bulunamadı", f"'{txt}' ürünü bulunamadı!")

    def load_product_groups(self):
        """Veritabanındaki gerçek ana grupları alır ve satış ekranına sekme olarak dizer"""
        while self.layout_grup_butonlari.count():
            item = self.layout_grup_butonlari.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                nested = item.layout()
                while nested.count():
                    nitem = nested.takeAt(0)
                    if nitem.widget():
                        nitem.widget().deleteLater()

        izgara = QGridLayout()
        izgara.setSpacing(8)

        btn_ana = QPushButton("ANA")
        btn_ana.setFixedHeight(40)
        btn_ana.setMinimumWidth(100)
        btn_ana.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ana.setStyleSheet("background-color: #3b82f6; color: white; font-weight: bold; border-radius: 6px; padding: 0 15px;")
        btn_ana.clicked.connect(lambda: self.urunleri_kategoriye_gore_filtrele("ANA"))
        izgara.addWidget(btn_ana, 0, 0)

        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute(
                "SELECT name FROM product_groups "
                "WHERE name IS NOT NULL AND name != '' "
                "ORDER BY id ASC"
            )
            gruplar = c.fetchall()
            conn.close()

            row = 0
            col = 1
            maksimum_sutun = 6

            for grup in gruplar:
                grup_ismi = str(grup[0]).strip()
                if not grup_ismi or grup_ismi.upper() in ("GRUPSUZLAR", "TUMU", "TÜMÜ", "ANA"):
                    continue

                btn_kategori = QPushButton(grup_ismi.upper())
                btn_kategori.setFixedHeight(40)
                btn_kategori.setMinimumWidth(110)
                btn_kategori.setCursor(Qt.CursorShape.PointingHandCursor)
                btn_kategori.setStyleSheet("""
                    QPushButton {
                        background-color: #f1f5f9; color: #334155;
                        font-weight: bold; border-radius: 6px;
                        padding: 0px 15px; border: 1px solid #cbd5e1;
                    }
                    QPushButton:hover { background-color: #e2e8f0; }
                """)
                btn_kategori.clicked.connect(lambda checked, g=grup_ismi: self.urunleri_kategoriye_gore_filtrele(g))
                izgara.addWidget(btn_kategori, row, col)

                col += 1
                if col >= maksimum_sutun:
                    col = 0
                    row += 1

        except Exception as e:
            print(f"Grup dizilimi hatası: {e}")

        self.layout_grup_butonlari.addLayout(izgara)
        self.layout_grup_butonlari.addStretch()

    def kategori_butonlari_yenile(self):
        self.load_product_groups()

    def kategoriye_gore_urunleri_getir(self, grup_adi):
        self.urunleri_kategoriye_gore_filtrele(grup_adi)

    def urunleri_kategoriye_gore_filtrele(self, kategori_adi):
        """
        Seçilen kategoriye (PC, PMMA, PET, POM vb.) ait kayıtlı ürünleri ve
        satış geçmişinde yer alan çeşitleri ekrana getirir.
        """
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        secilen_urunler = []

        if kategori_adi == "ANA":
            c.execute("""
                SELECT p.id, p.name, p.sell_price, p.stock_kg, COUNT(sh.id) as s_cnt
                FROM products p
                LEFT JOIN sales_history sh ON p.name = sh.product_name
                GROUP BY p.id
                ORDER BY s_cnt DESC, p.id DESC
                LIMIT 24
            """)
            secilen_urunler = [(r[0], r[1], r[2], r[3]) for r in c.fetchall()]
        elif kategori_adi == "GRUPSUZLAR":
            c.execute("""
                SELECT p.id, p.name, p.sell_price, p.stock_kg
                FROM products p
                LEFT JOIN product_groups g ON p.group_id = g.id
                WHERE g.name = 'GRUPSUZLAR' OR g.name IS NULL OR g.name = ''
                ORDER BY p.name ASC
            """)
            secilen_urunler = [(r[0], r[1], r[2], r[3]) for r in c.fetchall()]
        else:
            c.execute("""
                SELECT p.id, p.name, p.sell_price, p.stock_kg
                FROM products p
                JOIN product_groups g ON p.group_id = g.id
                WHERE UPPER(g.name) = UPPER(?)
                   OR UPPER(p.name) LIKE ?
                ORDER BY p.name ASC
            """, (kategori_adi, f"%{kategori_adi.upper()}%"))
            secilen_urunler = [(r[0], r[1], r[2], r[3]) for r in c.fetchall()]

        conn.close()
        self.urun_butonlarini_ciz(secilen_urunler)

    def urun_butonlarini_ciz(self, secilen_urunler):
        while self.grid_urun_layout.count():
            it = self.grid_urun_layout.takeAt(0)
            if it.widget():
                it.widget().deleteLater()

        col = 0
        row = 0
        for pid, name, price, *_rest in secilen_urunler:
            fiyat = float(price or 0.0)
            btn_prod = QPushButton(f"{name}\n\n₺ {fiyat:,.2f}")
            btn_prod.setFixedSize(110, 85)
            btn_prod.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_prod.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    font-weight: bold;
                    font-size: 11.5px;
                    color: #1e293b;
                    padding: 4px;
                }
                QPushButton:hover {
                    border: 2px solid #0284c7;
                    background-color: #f0f9ff;
                }
            """)
            btn_prod.clicked.connect(lambda _, p_id=pid, nm=name, pr=fiyat: self.hizli_urun_sepete_ekle(p_id, nm, pr))
            self.grid_urun_layout.addWidget(btn_prod, row, col)
            col += 1
            if col >= 3:
                col = 0
                row += 1

    def hizli_urunleri_ve_gruplari_yukle(self, group_id=None):
        self.load_product_groups()
        self.urunleri_kategoriye_gore_filtrele("ANA")

    def satisi_tamamla(self, odeme_turu):
        if not self.sepet:
            QMessageBox.warning(self, "Sepet Boş", "Lütfen sepete ürün ekleyiniz!")
            return

        toplam = sum(item["total"] for item in self.sepet)
        m_ad = self.secili_musteri["name"] if self.secili_musteri else "PERAKENDE MÜŞTERİ"
        m_id = self.secili_musteri["id"] if self.secili_musteri else None

        if "AÇIK HESAP" in odeme_turu and not self.secili_musteri:
            QMessageBox.warning(self, "Müşteri Gerekli", "Açık hesap satış için lütfen önce '+ Seç' ile müşteri belirleyiniz!")
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        # Peşin/nakit satışlarda tahsilat = toplam tutar, açık hesapta 0'dır.
        # (WhatsApp aktarımlarındaki toplu_kaydet ile aynı formül; önceden burada
        # sadece AÇIK HESAP durumunda cari güncelleniyordu, NAKİT satışlarda
        # müşteri seçili olsa bile borç/tahsilat/alışveriş sayısı hiç işlenmiyordu.)
        tahsilat_bu_satis = 0.0 if "AÇIK HESAP" in odeme_turu else toplam
        if m_id:
            c.execute("""
                UPDATE customers
                SET debt = debt + ?, payment = payment + ?, remaining_debt = remaining_debt + ?, shopping_count = shopping_count + 1
                WHERE id = ?
            """, (toplam, tahsilat_bu_satis, toplam - tahsilat_bu_satis, m_id))

        for it in self.sepet:
            c.execute("UPDATE products SET stock_kg = MAX(0, stock_kg - ?) WHERE id = ?", (it["qty"], it["id"]))

        sepet_veri = json.dumps(self.sepet)
        urun_adlari = ", ".join(it["name"] for it in self.sepet)
        toplam_kg = sum(it["qty"] for it in self.sepet)
        c.execute("""
            INSERT INTO sales_history (
                personnel_name, customer_id, customer_name, product_name, qty, price,
                total_amount, payment_type, payment_received, items_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, ("Genel", m_id, m_ad, urun_adlari, toplam_kg, (self.sepet[0]["price"] if self.sepet else 0.0), toplam, odeme_turu, tahsilat_bu_satis, sepet_veri))

        conn.commit()
        conn.close()

        self.sepeti_temizle()
        if hasattr(self, "load_products"):
            self.load_products()
        if hasattr(self, "load_customers"):
            self.load_customers()

    def en_son_satis_islem_iptal_et(self):
        """Veritabanındaki en son yapılan satış kaydını bulur, siler ve stokları geri yükler."""
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("""
                SELECT id, product_name, qty, customer_id, customer_name, total_amount, payment_received, items_json
                FROM sales_history ORDER BY id DESC LIMIT 1
            """)
            son_satis = c.fetchone()
            if not son_satis:
                QMessageBox.warning(self, "Uyarı", "İptal edilecek geçmiş satış hareketi bulunamadı!")
                conn.close()
                return

            satis_id, urun_adi, miktar, musteri_id, m_ad, tutar, tahsilat, items_json = son_satis
            cevap = QMessageBox.question(
                self, "İşlem İptal Onayı",
                f"En son yapılan satış iptal edilecek:\nÜrün: {urun_adi} ({miktar})\nMüşteri: {m_ad}\n\nOnaylıyor musunuz?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if cevap != QMessageBox.StandardButton.Yes:
                conn.close()
                return

            items = json.loads(items_json or "[]")
            if items:
                for it in items:
                    adet = abs(float(it.get("qty") or 0))
                    if it.get("id"):
                        c.execute("UPDATE products SET stock_kg = stock_kg + ? WHERE id = ?", (adet, it["id"]))
                    elif it.get("name"):
                        c.execute("UPDATE products SET stock_kg = stock_kg + ? WHERE name = ?", (adet, it["name"]))
            else:
                c.execute(
                    "UPDATE products SET stock_kg = stock_kg + ? WHERE name = ?",
                    (abs(float(miktar or 0)), urun_adi),
                )

            if musteri_id:
                tutar_f = float(tutar or 0)
                tahsilat_f = float(tahsilat or 0)
                c.execute(
                    """UPDATE customers
                       SET debt = MAX(0, debt - ?), payment = MAX(0, payment - ?),
                           remaining_debt = MAX(0, remaining_debt - ?),
                           shopping_count = MAX(0, shopping_count - 1)
                       WHERE id = ?""",
                    (tutar_f, tahsilat_f, tutar_f - tahsilat_f, musteri_id),
                )

            c.execute("DELETE FROM sales_history WHERE id = ?", (satis_id,))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Başarılı", "En son satış başarıyla iptal edildi ve stoklar geri yüklendi.")
            if hasattr(self, "load_products"):
                self.load_products()
            if hasattr(self, "load_customers"):
                self.load_customers()
            if hasattr(self, "urunleri_kategoriye_gore_filtrele"):
                self.urunleri_kategoriye_gore_filtrele("ANA")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"İptal işlemi sırasında hata oluştu:\n{e}")

    def son_islemi_iptal_et_onay(self):
        self.en_son_satis_islem_iptal_et()

    def create_whatsapp_sales_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        top_bar = QHBoxLayout()
        title = QLabel("💬 WHATSAPP SATIŞ AKTARIMI")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1e293b;")
        top_bar.addWidget(title)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        card_giris = QFrame()
        card_giris.setStyleSheet("background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px;")
        l_giris = QVBoxLayout(card_giris)
        l_giris.setContentsMargins(16, 14, 16, 14)
        l_giris.setSpacing(10)

        lbl_bilgi = QLabel("WhatsApp grubunuzdaki mesajları kopyalayıp aşağıdaki alana yapıştırın. Sistem müşteri isimlerindeki yazım hatalarını (örn: alu colak -> ALİ ÇOLAK) otomatik düzeltir, ürün ve standart fiyatları eşler.")
        lbl_bilgi.setStyleSheet("color: #475569; font-size: 13px; font-weight: 500;")
        lbl_bilgi.setWordWrap(True)
        l_giris.addWidget(lbl_bilgi)

        self.txt_wp_mesajlar = QTextEdit()
        self.txt_wp_mesajlar.setPlaceholderText("Örnek Giriş:\nalu colak 300 kg beyaz abs\n500 kg mavi pom 38 tl nakit\nmehmet yavz 2 ton seffaf pp\nbeyaz pom (kilo yazılmazsa eksik uyarısı verir)")
        self.txt_wp_mesajlar.setFixedHeight(120)
        self.txt_wp_mesajlar.setStyleSheet("""
            QTextEdit {
                border: 2px solid #cbd5e1;
                border-radius: 6px;
                padding: 8px;
                font-size: 13.5px;
                background: #f8fafc;
            }
            QTextEdit:focus { border-color: #0284c7; background: #ffffff; }
        """)
        l_giris.addWidget(self.txt_wp_mesajlar)

        h_btn_bar = QHBoxLayout()
        h_btn_bar.addStretch()
        btn_ayristir = QPushButton("  ⚡ MESAJLARI ANALİZ ET VE TABLOYA DÖK  ")
        btn_ayristir.setFixedHeight(38)
        btn_ayristir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ayristir.setStyleSheet("background-color: #0284c7; color: white; font-size: 13px; font-weight: bold; border-radius: 5px; border: none; padding: 0 16px;")
        btn_ayristir.clicked.connect(self.wp_mesajlarini_isle)
        h_btn_bar.addWidget(btn_ayristir)
        l_giris.addLayout(h_btn_bar)
        layout.addWidget(card_giris)

        # ---------------- MODERN PREMIUM AKTARIM PANELİ ----------------
        card_dev_yukleme = QFrame()
        card_dev_yukleme.setObjectName("ModernDevKarti")
        card_dev_yukleme.setStyleSheet("""
            QFrame#ModernDevKarti {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
            }
        """)
        l_dev = QVBoxLayout(card_dev_yukleme)
        l_dev.setContentsMargins(20, 16, 20, 16)
        l_dev.setSpacing(12)

        # Üst Başlık ve Rozet (Badge)
        h_baslik_row = QHBoxLayout()
        h_baslik_row.setSpacing(8)

        lbl_rozet = QLabel("10 GB+ MOTOR")
        lbl_rozet.setFixedHeight(22)
        lbl_rozet.setStyleSheet("""
            background-color: #eff6ff;
            color: #1d4ed8;
            font-size: 11px;
            font-weight: 800;
            padding: 2px 8px;
            border-radius: 4px;
            border: 1px solid #bfdbfe;
        """)
        h_baslik_row.addWidget(lbl_rozet)

        lbl_dev_bilgi = QLabel("Arşiv & Toplu Veri İçe Aktarım Merkezi")
        lbl_dev_bilgi.setStyleSheet("font-size: 13.5px; font-weight: 700; color: #0f172a;")
        h_baslik_row.addWidget(lbl_dev_bilgi)

        lbl_alt_aciklama = QLabel("WhatsApp (.txt), Muhasebe (.csv) veya Excel yedeklerini arka planda sıfır donma ile işler.")
        lbl_alt_aciklama.setStyleSheet("font-size: 12px; color: #64748b; font-weight: 500;")
        h_baslik_row.addWidget(lbl_alt_aciklama)
        h_baslik_row.addStretch()
        l_dev.addLayout(h_baslik_row)

        # Butonlar Alanı
        h_butonlar = QHBoxLayout()
        h_butonlar.setSpacing(10)

        # 1. Buton: WhatsApp TXT
        btn_wp_dev_txt = QPushButton("  WhatsApp Arşivi (.txt) Yükle")
        btn_wp_dev_txt.setFixedHeight(40)
        btn_wp_dev_txt.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_wp_dev_txt.setStyleSheet("""
            QPushButton {
                background-color: #0f172a;
                color: #ffffff;
                font-size: 13px;
                font-weight: 600;
                border-radius: 6px;
                padding: 0 16px;
                border: 1px solid #0f172a;
            }
            QPushButton:hover {
                background-color: #1e293b;
                border-color: #1e293b;
            }
            QPushButton:pressed {
                background-color: #020617;
            }
        """)
        btn_wp_dev_txt.clicked.connect(self.dev_dosya_sec_ve_baslat)
        h_butonlar.addWidget(btn_wp_dev_txt)

        # 2. Buton: CSV / Excel
        btn_dev_csv = QPushButton("  CSV / Excel Satış Yedeği Yükle")
        btn_dev_csv.setFixedHeight(40)
        btn_dev_csv.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_dev_csv.setStyleSheet("""
            QPushButton {
                background-color: #f8fafc;
                color: #1e293b;
                font-size: 13px;
                font-weight: 600;
                border-radius: 6px;
                padding: 0 16px;
                border: 1px solid #cbd5e1;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
                border-color: #94a3b8;
                color: #0f172a;
            }
            QPushButton:pressed {
                background-color: #e2e8f0;
            }
        """)
        btn_dev_csv.clicked.connect(self.dev_dosya_sec_ve_baslat)
        h_butonlar.addWidget(btn_dev_csv)

        # 3. İptal / Durdur Butonu (Yalnızca çalışırken aktifleşir)
        self.btn_aktarim_durdur = QPushButton("İşlemi Durdur")
        self.btn_aktarim_durdur.setFixedHeight(40)
        self.btn_aktarim_durdur.setEnabled(False)
        self.btn_aktarim_durdur.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_aktarim_durdur.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #dc2626;
                font-size: 13px;
                font-weight: 600;
                border-radius: 6px;
                padding: 0 14px;
                border: 1px solid #fecaca;
            }
            QPushButton:hover {
                background-color: #fef2f2;
                border-color: #f87171;
            }
            QPushButton:disabled {
                background-color: #f8fafc;
                color: #cbd5e1;
                border-color: #f1f5f9;
            }
        """)
        self.btn_aktarim_durdur.clicked.connect(self.dev_aktarimi_durdur)
        h_butonlar.addWidget(self.btn_aktarim_durdur)

        h_butonlar.addStretch()
        l_dev.addLayout(h_butonlar)

        # Durum ve İlerleme Çubuğu Bölümü
        h_status_bar = QHBoxLayout()
        h_status_bar.setSpacing(12)

        self.lbl_dev_canli_durum = QLabel("Aktarım için dosya bekleniyor")
        self.lbl_dev_canli_durum.setStyleSheet("font-size: 12.5px; font-weight: 600; color: #64748b;")
        h_status_bar.addWidget(self.lbl_dev_canli_durum)

        self.progress_dev = QProgressBar()
        self.progress_dev.setFixedHeight(6)
        self.progress_dev.setTextVisible(False)
        self.progress_dev.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 3px;
                background-color: #e2e8f0;
            }
            QProgressBar::chunk {
                background-color: #0284c7;
                border-radius: 3px;
            }
        """)
        self.progress_dev.setVisible(False)
        h_status_bar.addWidget(self.progress_dev, stretch=1)

        l_dev.addLayout(h_status_bar)
        layout.addWidget(card_dev_yukleme)

        card_tablo = QFrame()
        card_tablo.setStyleSheet("background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px;")
        l_tablo = QVBoxLayout(card_tablo)
        l_tablo.setContentsMargins(16, 14, 16, 14)
        l_tablo.setSpacing(10)

        lbl_onizleme = QLabel("📋 SATIŞ ÖNİZLEME VE DÜZELTME TABLOSU (Eksik kiloları veya hatalı alanları doğrudan hücreye tıklayarak düzeltebilirsiniz)")
        lbl_onizleme.setStyleSheet("font-size: 13px; font-weight: bold; color: #0f172a;")
        l_tablo.addWidget(lbl_onizleme)

        self.table_wp_onizleme = QTableWidget()
        self.table_wp_onizleme.setColumnCount(9)
        self.table_wp_onizleme.setHorizontalHeaderLabels([
            "DURUM", "TARİH", "PERSONEL", "MÜŞTERİ / FİRMA", "ÜRÜN", "MİKTAR (KG)", "BİRİM FİYAT", "TUTAR", "TAHSİLAT"
        ])
        self.table_wp_onizleme.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_wp_onizleme.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table_wp_onizleme.setColumnWidth(0, 110)
        self.table_wp_onizleme.verticalHeader().setVisible(False)
        self.table_wp_onizleme.setStyleSheet("""
            QTableWidget { border: 1px solid #e2e8f0; font-size: 13px; background: #ffffff; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 38px; border-bottom: 2px solid #cbd5e1; }
            QTableWidget::item { padding: 4px; }
        """)
        self.table_wp_onizleme.cellChanged.connect(self.wp_tablo_hucre_degisti)
        self.table_wp_onizleme.setEditTriggers(
            QAbstractItemView.EditTrigger.CurrentChanged |
            QAbstractItemView.EditTrigger.SelectedClicked
        )
        l_tablo.addWidget(self.table_wp_onizleme)

        h_alt_bar = QHBoxLayout()
        self.lbl_wp_durum_ozet = QLabel("Henüz mesaj işlenmedi.")
        self.lbl_wp_durum_ozet.setStyleSheet("font-weight: bold; color: #475569; font-size: 13px;")
        h_alt_bar.addWidget(self.lbl_wp_durum_ozet)
        h_alt_bar.addStretch()
        self.btn_wp_onayla = QPushButton("  ✓ TÜM SATIŞLARI SİSTEME KAYDET (STOK & BORÇLARA İŞLE)  ")
        self.btn_wp_onayla.setFixedHeight(42)
        self.btn_wp_onayla.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_wp_onayla.setStyleSheet("background-color: #10b981; color: white; font-size: 14px; font-weight: bold; border-radius: 6px; border: none; padding: 0 20px;")
        self.btn_wp_onayla.clicked.connect(self.wp_satislarini_veritabanina_kaydet)
        h_alt_bar.addWidget(self.btn_wp_onayla)
        l_tablo.addLayout(h_alt_bar)
        layout.addWidget(card_tablo)

        self.wp_cozumlenen_veriler = []
        return page

    def musteri_veya_musterisiz_al(self, m_id, m_ad, cursor=None):
        kendi_baglantisi = False
        if cursor is None:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            kendi_baglantisi = True
        else:
            c = cursor

        if m_id:
            if kendi_baglantisi:
                conn.close()
            return m_id, m_ad

        m_ad_temiz = buyuk_harf(str(m_ad or "").replace("🏢", "").replace("👤", "").strip())
        musteri_siz = buyuk_harf("MÜŞTERİSİZ SATIŞ")
        genel_polimer = buyuk_harf("GENEL POLİMER")
        if not m_ad_temiz or m_ad_temiz in [musteri_siz, genel_polimer]:
            c.execute("SELECT id, name FROM customers WHERE name = ? LIMIT 1", (musteri_siz,))
            row = c.fetchone()
            if row:
                cid = row[0]
            else:
                c.execute("INSERT INTO customers (name, debt, remaining_debt) VALUES (?, 0.0, 0.0)", (musteri_siz,))
                cid = c.lastrowid
            if kendi_baglantisi:
                conn.commit()
                conn.close()
            return cid, musteri_siz

        c.execute("SELECT id, name FROM customers WHERE upper(name) = ? LIMIT 1", (m_ad_temiz,))
        row = c.fetchone()
        if row:
            cid, cname = row[0], row[1]
        else:
            c.execute("""
                INSERT INTO customers (name, debt, payment, remaining_debt, shopping_count)
                VALUES (?, 0, 0, 0, 0)
            """, (m_ad_temiz,))
            cid = c.lastrowid
            cname = m_ad_temiz

        if kendi_baglantisi:
            conn.commit()
            conn.close()
        return cid, cname

    def wp_mesajlarini_isle(self):
        try:
            ham_metin = self.txt_wp_mesajlar.toPlainText().strip()
            if not ham_metin:
                QMessageBox.warning(self, "Uyarı", "Lütfen önce WhatsApp mesajlarını kutuya yapıştırınız!")
                return
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("SELECT id, name FROM customers")
            tum_musteriler = c.fetchall()
            c.execute("SELECT id, name, sell_price FROM products")
            tum_urunler = c.fetchall()
            conn.close()
            sonuclar = WhatsAppSatisAyristirici.bloklari_ayristir(ham_metin, tum_musteriler, tum_urunler)
            self.wp_cozumlenen_veriler = sonuclar if sonuclar else []
            if not self.wp_cozumlenen_veriler:
                QMessageBox.information(self, "Bilgi", "Mesajlarda geçerli bir satış satırı tespit edilemedi.")
                return
            self.wp_tabloyu_doldur()
        except Exception as err:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Ayrıştırma Hatası", f"Mesajlar işlenirken bir sorun oluştu:\n{str(err)}")

    def wp_tabloyu_doldur(self):
        self.table_wp_onizleme.blockSignals(True)
        try:
            self.table_wp_onizleme.cellClicked.disconnect()
        except Exception:
            pass
        try:
            self.table_wp_onizleme.cellChanged.disconnect()
        except Exception:
            pass
        try:
            self.table_wp_onizleme.currentCellChanged.disconnect()
        except Exception:
            pass
        self.table_wp_onizleme.setColumnCount(9)
        self.table_wp_onizleme.setHorizontalHeaderLabels([
            "DURUM", "TARİH", "PERSONEL", "MÜŞTERİ ADI", "ÜRÜN TANIMI", "MİKTAR (KG)", "BİRİM FİYAT", "TUTAR", "TAHSİLAT"
        ])
        toplam_kayit = len(self.wp_cozumlenen_veriler)
        self.table_wp_onizleme.setRowCount(toplam_kayit)

        toplam_kg = 0.0
        toplam_tl = 0.0
        for r_idx in range(toplam_kayit):
            d = self.wp_cozumlenen_veriler[r_idx]
            self.table_wp_onizleme.setRowHeight(r_idx, 38)

            miktar = float(d.get("miktar", 0.0) or 0.0)
            fiyat = float(d.get("fiyat", 0.0) or 0.0)
            tutar = round(miktar * fiyat, 2)
            d["tutar"] = tutar
            tahsilat = float(d.get("odeme_alindi", 0.0) or 0.0)

            toplam_kg += miktar
            toplam_tl += tutar

            durum_metin = "✓ HAZIR" if miktar > 0 else "⚠ KİLO YOK"
            it_st = QTableWidgetItem(durum_metin)
            it_st.setForeground(QColor("#15803d" if miktar > 0 else "#b45309"))
            it_st.setBackground(QColor("#dcfce7" if miktar > 0 else "#fef3c7"))
            it_st.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_st.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            it_st.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_wp_onizleme.setItem(r_idx, 0, it_st)

            w_dt = QWidget()
            l_dt = QHBoxLayout(w_dt)
            l_dt.setContentsMargins(2, 2, 2, 2)
            l_dt.setAlignment(Qt.AlignmentFlag.AlignCenter)

            dt_cell = QDateEdit()
            dt_cell.setCalendarPopup(True)
            dt_cell.setDisplayFormat("dd.MM.yyyy")
            dt_cell.setFixedHeight(30)

            parcalar = str(d.get("tarih", "")).split()
            ham_tarih = parcalar[0] if parcalar else ""
            q_date = QDate.fromString(ham_tarih, "dd.MM.yyyy")
            if not q_date.isValid():
                q_date = QDate.fromString(ham_tarih, "dd.MM.yy")
            if q_date.isValid():
                dt_cell.setDate(q_date)
            elif hasattr(self, "dt_wp_tarih"):
                dt_cell.setDate(self.dt_wp_tarih.date())
            else:
                dt_cell.setDate(QDate.currentDate())

            dt_cell.dateChanged.connect(lambda qd, row_no=r_idx: self.wp_satir_tarih_degisti(row_no, qd))
            l_dt.addWidget(dt_cell)
            self.table_wp_onizleme.setCellWidget(r_idx, 1, w_dt)

            it_per = QTableWidgetItem(f"👤 {d.get('personel', 'Genel')}")
            it_per.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            it_per.setForeground(QColor("#0369a1"))
            it_per.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_wp_onizleme.setItem(r_idx, 2, it_per)

            it_mus = QTableWidgetItem(str(d.get("musteri_ad", "MÜŞTERİSİZ SATIŞ")))
            it_mus.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_mus.setForeground(QColor("#0f172a"))
            it_mus.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable)
            self.table_wp_onizleme.setItem(r_idx, 3, it_mus)

            it_ur = QTableWidgetItem(str(d.get("urun_ad", "-")))
            it_ur.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_ur.setForeground(QColor("#047857"))
            it_ur.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable)
            self.table_wp_onizleme.setItem(r_idx, 4, it_ur)

            miktar_str = f"{int(miktar)}" if miktar.is_integer() and miktar > 0 else (f"{miktar:.1f}" if miktar > 0 else "KİLO GİRİN")
            it_q = QTableWidgetItem(miktar_str)
            it_q.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_q.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_q.setForeground(QColor("#dc2626" if miktar == 0 else "#0f172a"))
            it_q.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable)
            self.table_wp_onizleme.setItem(r_idx, 5, it_q)

            it_p = QTableWidgetItem(f"{fiyat:,.2f} ₺")
            it_p.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it_p.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable)
            self.table_wp_onizleme.setItem(r_idx, 6, it_p)

            it_t = QTableWidgetItem(f"{tutar:,.2f} ₺")
            it_t.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_t.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it_t.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_wp_onizleme.setItem(r_idx, 7, it_t)

            it_pay = QTableWidgetItem(f"✓ {tahsilat:,.2f} ₺" if tahsilat > 0 else "-")
            it_pay.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_pay.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table_wp_onizleme.setItem(r_idx, 8, it_pay)

        self.table_wp_onizleme.blockSignals(False)
        self.table_wp_onizleme.cellClicked.connect(self.wp_metinde_satiri_vurgula)
        self.table_wp_onizleme.currentCellChanged.connect(lambda r, c, pr, pc: self.wp_metinde_satiri_vurgula(r, c))
        self.table_wp_onizleme.cellChanged.connect(self.wp_tablo_hucre_degisti)
        if hasattr(self, "lbl_wp_durum_ozet"):
            self.lbl_wp_durum_ozet.setText(f"Toplam Tonaj: {toplam_kg:,.0f} KG | Ciro: {toplam_tl:,.2f} ₺")

    def wp_satir_tarih_degisti(self, idx, qdate):
        if 0 <= idx < len(self.wp_cozumlenen_veriler):
            self.wp_cozumlenen_veriler[idx]["tarih"] = qdate.toString("dd.MM.yyyy")

    def wp_metinde_satiri_vurgula(self, row, col):
        """Tabloda bir satıra/hücreye tıklandığında üst kutuda ilgili WhatsApp mesajını mavi seçer"""
        if row < 0 or row >= len(self.wp_cozumlenen_veriler):
            return
        d = self.wp_cozumlenen_veriler[row]
        c_start = d.get("char_start", 0)
        c_end = d.get("char_end", 0)
        # 1. ÜST KUTUDAKİ METNİ MAVİ VURGULA (SEÇİM ALANI)
        try:
            doc_len = len(self.txt_wp_mesajlar.toPlainText())
            if c_end > c_start and c_start < doc_len:
                cursor = self.txt_wp_mesajlar.textCursor()
                cursor.setPosition(min(c_start, doc_len))
                cursor.setPosition(min(c_end, doc_len), QTextCursor.MoveMode.KeepAnchor)
                self.txt_wp_mesajlar.setTextCursor(cursor)
                self.txt_wp_mesajlar.ensureCursorVisible()
        except Exception:
            pass
        # 2. EĞER MİKTAR (KİLO) VEYA FİYAT HÜCRESİYSE DÜZENLEMEYE HAZIR HALE GETİR
        if col in [5, 6]:  # 5: Miktar (KG), 6: Fiyat
            item = self.table_wp_onizleme.item(row, col)
            if item:
                if "KİLO" in item.text():
                    # 'KİLO GİRİN' yazısını temizle ki klavyeden basılan rakam doğrudan yazılsın
                    item.setText("")
                self.table_wp_onizleme.editItem(item)

    def wp_tablo_hucre_degisti(self, row, col):
        if row < 0 or row >= len(self.wp_cozumlenen_veriler):
            return
        try:
            it_per = self.table_wp_onizleme.item(row, 2)
            it_m = self.table_wp_onizleme.item(row, 3)
            it_u = self.table_wp_onizleme.item(row, 4)
            it_q = self.table_wp_onizleme.item(row, 5)
            it_p = self.table_wp_onizleme.item(row, 6)
            it_pay = self.table_wp_onizleme.item(row, 8)
            if it_per:
                self.wp_cozumlenen_veriler[row]["personel"] = it_per.text().replace("👤", "").strip()
            if it_m:
                self.wp_cozumlenen_veriler[row]["musteri_ad"] = buyuk_harf(it_m.text().replace("🏢", "").replace("👤", "").strip())
                self.wp_cozumlenen_veriler[row]["musteri_tur"] = WhatsAppSatisAyristirici.tur_tespit_et(
                    self.wp_cozumlenen_veriler[row]["musteri_ad"]
                )
            if it_u:
                self.wp_cozumlenen_veriler[row]["urun_ad"] = buyuk_harf(it_u.text().strip())

            if it_q:
                q_text = it_q.text().replace("KİLO GİRİN", "").replace("KİLOGİRİN", "").replace("KG", "").replace(" ", "").strip()
                if "." in q_text and "," not in q_text:
                    parca = q_text.split(".")
                    if len(parca) == 2 and len(parca[1]) == 3 and parca[0].lstrip("-").isdigit():
                        q_text = q_text.replace(".", "")
                else:
                    q_text = q_text.replace(",", ".")
                q_val = float(q_text) if q_text else 0.0
                self.wp_cozumlenen_veriler[row]["miktar"] = q_val
            else:
                q_val = float(self.wp_cozumlenen_veriler[row].get("miktar", 0.0) or 0.0)

            if it_p:
                p_text = it_p.text().replace("₺", "").replace(" ", "").strip()
                if "," in p_text and "." in p_text:
                    p_text = p_text.replace(".", "").replace(",", ".")
                else:
                    p_text = p_text.replace(",", ".")
                p_val = float(p_text) if p_text else 0.0
                self.wp_cozumlenen_veriler[row]["fiyat"] = p_val
            else:
                p_val = float(self.wp_cozumlenen_veriler[row].get("fiyat", 0.0) or 0.0)

            if q_val > 0:
                self.wp_cozumlenen_veriler[row]["durum"] = "TAMAM"

            pay_txt = (it_pay.text() if it_pay else "").replace("-", "0").replace("₺", "").replace("✓", "").replace(" ", "").strip()
            if "," in pay_txt and "." in pay_txt:
                pay_txt = pay_txt.replace(".", "").replace(",", ".")
            else:
                pay_txt = pay_txt.replace(",", ".")
            self.wp_cozumlenen_veriler[row]["odeme_alindi"] = float(pay_txt) if pay_txt else 0.0

            yeni_tutar = round(q_val * p_val, 2)
            self.wp_cozumlenen_veriler[row]["tutar"] = yeni_tutar

            self.table_wp_onizleme.blockSignals(True)
            it_tot = self.table_wp_onizleme.item(row, 7)
            if it_tot:
                it_tot.setText(f"{yeni_tutar:,.2f} ₺")
            it_st = self.table_wp_onizleme.item(row, 0)
            if q_val > 0 and it_st:
                it_st.setText("✓ HAZIR")
                it_st.setForeground(QColor("#15803d"))
                it_st.setBackground(QColor("#dcfce7"))
                if it_q:
                    it_q.setForeground(QColor("#0f172a"))
            self.table_wp_onizleme.blockSignals(False)

            toplam_kg = sum(float(x.get("miktar", 0.0) or 0.0) for x in self.wp_cozumlenen_veriler)
            toplam_tl = sum(float(x.get("tutar", 0.0) or 0.0) for x in self.wp_cozumlenen_veriler)
            self.lbl_wp_durum_ozet.setText(f"Toplam Tonaj: {toplam_kg:,.0f} KG | Ciro: {toplam_tl:,.2f} ₺")
        except (ValueError, AttributeError):
            pass

    def urunu_tekil_al_veya_olustur(self, urun_adi, standart_fiyat, cursor=None):
        temiz_grup, urun_adi = standart_urun_ve_grup_belirle(str(urun_adi or ""), "")
        if temiz_grup == "PP":
            hedef_grup = "PP MOBLEN"
        elif temiz_grup in ["DİĞER", ""]:
            hedef_grup = "GRUPSUZLAR"
        else:
            hedef_grup = temiz_grup
        kendi_baglantisi = False
        if cursor is None:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            kendi_baglantisi = True
        else:
            c = cursor
        c.execute("SELECT id, group_id, sell_price FROM products WHERE name = ? COLLATE NOCASE LIMIT 1", (urun_adi,))
        row = c.fetchone()

        if hedef_grup.upper() in ["WHATSAPP", "GENEL", "GENEL POLİMER", "TÜMÜ"]:
            hedef_grup = "GRUPSUZLAR"
        c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (hedef_grup,))
        g_row = c.fetchone()
        if g_row:
            gid = g_row[0]
        else:
            c.execute("INSERT INTO product_groups (name) VALUES (?)", (hedef_grup,))
            gid = c.lastrowid
        if row:
            p_id = row[0]
            if row[1] != gid:
                c.execute("UPDATE products SET group_id = ? WHERE id = ?", (gid, p_id))
        else:
            c.execute("""
                INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price)
                VALUES (?, ?, 0.0, 0.0, ?)
            """, (gid, urun_adi, float(standart_fiyat or 40.0)))
            p_id = c.lastrowid
        if kendi_baglantisi:
            conn.commit()
            conn.close()
        return p_id

    def wp_satislarini_veritabanina_kaydet(self):
        if not self.wp_cozumlenen_veriler:
            QMessageBox.warning(self, "Kayıt Yok", "İşlenecek satış bulunamadı!")
            return

        eksik_kilo = [d for d in self.wp_cozumlenen_veriler if float(d.get("miktar", 0.0) or 0.0) <= 0]
        if eksik_kilo:
            QMessageBox.warning(self, "Eksik Kilo!", f"{len(eksik_kilo)} satırda kilo eksik. Lütfen tablodan kiloları giriniz.")
            return

        reply = QMessageBox.question(
            self,
            "TOPLU SATIŞ ONAYI",
            f"{len(self.wp_cozumlenen_veriler)} adet WhatsApp satışı kaydedilecek.\nOnaylıyor musunuz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        conn = None
        try:
            conn = sqlite3.connect(DB_NAME, timeout=30)
            c = conn.cursor()
            c.execute("SELECT id, name, sell_price FROM products")
            urun_map = {p[1]: (p[0], p[2]) for p in c.fetchall()}
            c.execute("SELECT id, name FROM customers")
            musteri_map = {m[1]: m[0] for m in c.fetchall()}
            grup_id_map = {}

            def grup_id_bulucu(cur, u_ad):
                hedef = "GRUPSUZLAR"
                u_up = (u_ad or "").upper()
                if "POM" in u_up:
                    hedef = "POM"
                elif "ABS" in u_up:
                    hedef = "ABS"
                if hedef not in grup_id_map:
                    cur.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (hedef,))
                    row = cur.fetchone()
                    if row:
                        grup_id_map[hedef] = row[0]
                    else:
                        cur.execute("INSERT INTO product_groups (name) VALUES (?)", (hedef,))
                        grup_id_map[hedef] = cur.lastrowid
                return grup_id_map[hedef]

            ozet = wp_satis_kayit_servisi(c, conn, self.wp_cozumlenen_veriler, musteri_map, urun_map, grup_id_bulucu)
            QMessageBox.information(
                self, "Başarılı",
                f"Kaydedilen: {ozet['kaydedilen']}\nDaha önce mevcut: {ozet['mevcut']}\nİnceleme: {ozet['inceleme']}"
            )
            self.txt_wp_mesajlar.clear()
            self.wp_cozumlenen_veriler.clear()
            self.table_wp_onizleme.setRowCount(0)
            if hasattr(self, "lbl_wp_durum_ozet"):
                self.lbl_wp_durum_ozet.setText("Tüm işlemler başarıyla kaydedildi.")
            if hasattr(self, "load_products"):
                self.load_products()
            if hasattr(self, "load_customers"):
                self.load_customers()
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Veritabanı Hatası", f"Kayıt sırasında hata oluştu:\n{str(e)}")
        finally:
            if conn is not None:
                conn.close()

    def dev_dosya_sec_ve_baslat(self):
        dosya, _ = QFileDialog.getOpenFileName(
            self,
            "Yüklenecek Dev Arşiv Dosyasını Seç",
            "",
            "Desteklenen Arşivler (*.txt *.csv *.xlsx);;Metin Dosyaları (*.txt);;CSV Dosyaları (*.csv);;Excel (*.xlsx)"
        )
        if not dosya:
            return

        reply = QMessageBox.question(
            self,
            "DEV AKTARIM BAŞLATILIYOR",
            f"Seçilen dosya doğrudan arka plan akış motoruyla okunacak.\n"
            f"İşlem sırasında dükkan satışlarınızı yapmaya devam edebilirsiniz.\n\n"
            f"Dosya: {os.path.basename(dosya)}\nBoyut: {os.path.getsize(dosya) / (1024*1024):,.2f} MB\n\nBaşlatılsın mı?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.progress_dev.setVisible(True)
        self.progress_dev.setValue(0)
        self.btn_aktarim_durdur.setEnabled(True)
        self.lbl_dev_canli_durum.setText("Aktarım başlatıldı...")

        self.dev_thread = WhatsAppDevAktarimThread(dosya, parent=self)
        self.dev_thread.ilerleme_sinyali.connect(self.progress_dev.setValue)
        self.dev_thread.durum_sinyali.connect(self.lbl_dev_canli_durum.setText)
        self.dev_thread.tamamlandi_sinyali.connect(self.dev_aktarim_bitti)
        self.dev_thread.start()

    def dev_aktarimi_durdur(self):
        if hasattr(self, "dev_thread") and self.dev_thread.isRunning():
            self.dev_thread.iptal_edildi = True
            self.lbl_dev_canli_durum.setText("İşlem kullanıcı tarafından durduruldu! O ana kadar olan veriler kaydedildi.")
            self.btn_aktarim_durdur.setEnabled(False)

    def dev_aktarim_bitti(self, sonuc):
        self.btn_aktarim_durdur.setEnabled(False)
        self.progress_dev.setValue(100)
        if sonuc.get("hata"):
            QMessageBox.critical(self, "Aktarım Hatası", str(sonuc.get("hata")))
            return
        toplam_satir = sonuc.get("toplam_satir", 0)
        toplam_satis = sonuc.get("toplam_satis", 0)
        QMessageBox.information(
            self,
            "BÜYÜK AKTARIM TAMAMLANDI!",
            f"✅ {toplam_satir:,} satırlık dev veri akışı başarıyla tarandı.\n\n"
            f"📦 Toplam {toplam_satis:,} adet satış başarıyla veritabanına, cari borçlara ve stoklara işlendi!"
        )
        if hasattr(self, "load_products"):
            self.load_products()
        if hasattr(self, "load_customers"):
            self.load_customers()
        if hasattr(self, "load_personnel_data"):
            self.load_personnel_data()

    def create_reports_hub_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        ust = QHBoxLayout()
        ust.addStretch()
        btn_excel = QPushButton("📊 Excel'e Aktar")
        btn_excel.setFixedHeight(36)
        btn_excel.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_excel.setStyleSheet("background: #16a34a; color: white; font-weight: bold; padding: 0 14px; border-radius: 6px;")
        btn_excel.clicked.connect(self.grafikleri_excelle_aktar)
        ust.addWidget(btn_excel)
        layout.addLayout(ust)
        self.report_stack = QStackedWidget()
        for builder in (
            self._build_gunluk_rapor_sayfasi, self._build_tarihsel_rapor_sayfasi,
            self._build_urunsel_rapor_sayfasi, self._build_grupsal_rapor_sayfasi,
            self._build_stok_hareket_sayfasi, self._build_personel_hareket_sayfasi,
        ):
            self.report_stack.addWidget(builder())
        layout.addWidget(self.report_stack)
        return page

    def grafikleri_excelle_aktar(self):
        """Tüm rapor sayfalarındaki verileri tek Excel dosyasında farklı sekmeler halinde dışarı aktarır."""
        if not OPENPYXL_VAR:
            QMessageBox.warning(self, "Eksik Kütüphane", "Excel aktarımı için openpyxl yüklü değil.")
            return
        try:
            dosya_yolu, _ = QFileDialog.getSaveFileName(
                self, "Raporları Excel'e Aktar", "BenimPOS_Detayli_Raporlar.xlsx", "Excel Dosyası (*.xlsx)"
            )
            if not dosya_yolu:
                return
            wb = openpyxl.Workbook()
            ws_satis = wb.active
            ws_satis.title = "Satış Raporu"
            ws_satis.append(["Tarih", "Personel", "Müşteri", "Ürün", "Miktar (KG)", "Birim Fiyat", "Toplam Tutar (TL)"])

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT created_at, personnel_name, customer_name, product_name, qty, price, total_amount FROM sales_history"
            )
            for satir in cursor.fetchall():
                ws_satis.append(list(satir))

            try:
                ws_finans = wb.create_sheet(title="Çek ve Senetler")
                ws_finans.append(["Evrak Tipi", "Vade Tarihi", "Tutar", "Banka", "Keşideci/Sahip", "Durum"])
                cursor.execute(
                    "SELECT evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi_kesideci, durum FROM finans_evraklar"
                )
                for satir in cursor.fetchall():
                    ws_finans.append(list(satir))
            except Exception:
                pass

            try:
                ws_stok = wb.create_sheet(title="Ürün Bazlı Stok Özeti")
                ws_stok.append(["Ürün Adı", "Grup Adı", "Toplam Satış Miktarı (KG)"])
                cursor.execute(
                    """
                    SELECT sh.product_name, COALESCE(g.name, ''), SUM(sh.qty)
                    FROM sales_history sh
                    LEFT JOIN products p ON p.name = sh.product_name
                    LEFT JOIN product_groups g ON g.id = p.group_id
                    GROUP BY sh.product_name
                    """
                )
                for satir in cursor.fetchall():
                    ws_stok.append(list(satir))
            except Exception:
                pass

            conn.close()
            wb.save(dosya_yolu)
            QMessageBox.information(self, "Başarılı", f"Tüm raporlar başarıyla Excel'e aktarıldı:\n{dosya_yolu}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel'e aktarım sırasında bir hata oluştu:\n{e}")

    def create_finance_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        baslik_layout = QHBoxLayout()
        baslik = QLabel("💼 ÇEK & SENET YÖNETİMİ")
        baslik.setStyleSheet("font-size: 20px; font-weight: bold; color: #1e293b;")
        baslik_layout.addWidget(baslik)
        baslik_layout.addStretch()
        self.cmb_siralama = QComboBox()
        self.cmb_siralama.addItems([
            "📅 Vade Tarihi (Yakından Uzağa)",
            "📅 Vade Tarihi (Uzaktan Yakına)",
            "💰 Tutar (Yüksekten Düşüğe)",
            "💰 Tutar (Düşükten Yükseğe)",
        ])
        self.cmb_siralama.setFixedHeight(36)
        self.cmb_siralama.setStyleSheet("background: white; border: 1px solid #cbd5e1; border-radius: 6px; padding: 0 8px; font-weight: bold;")
        self.cmb_siralama.currentIndexChanged.connect(self.finans_tablosunu_guncelle)
        baslik_layout.addWidget(self.cmb_siralama)
        self.btn_finans_secilen_sil = QPushButton("🗑️ Seçilenleri Sil")
        self.btn_finans_secilen_sil.setFixedHeight(36)
        self.btn_finans_secilen_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_finans_secilen_sil.setStyleSheet(
            "QPushButton { background-color: #fff7ed; color: #c2410c; font-weight: bold; "
            "border: 1px solid #fdba74; border-radius: 6px; padding: 0 10px; }"
        )
        self.btn_finans_secilen_sil.clicked.connect(self.finans_secilenleri_sil)
        baslik_layout.addWidget(self.btn_finans_secilen_sil)
        self.btn_finans_tum_sil = QPushButton("⚠️ Tümünü Sil")
        self.btn_finans_tum_sil.setFixedHeight(36)
        self.btn_finans_tum_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_finans_tum_sil.setStyleSheet(
            "QPushButton { background-color: #fef2f2; color: #b91c1c; font-weight: bold; "
            "border: 1px solid #fecaca; border-radius: 6px; padding: 0 10px; }"
        )
        self.btn_finans_tum_sil.clicked.connect(self.finans_tumunu_sil)
        baslik_layout.addWidget(self.btn_finans_tum_sil)
        btn_aktar = QPushButton("📊 Excel'e Aktar")
        btn_aktar.setFixedHeight(36)
        btn_aktar.setStyleSheet("background: #16a34a; color: white; font-weight: bold; padding: 0 12px; border-radius: 6px;")
        btn_aktar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_aktar.clicked.connect(self.finans_evraklarini_disari_aktar)
        baslik_layout.addWidget(btn_aktar)
        btn_mail_ayar = QPushButton("📧 E-Posta Ayarları")
        btn_mail_ayar.setFixedHeight(36)
        btn_mail_ayar.setStyleSheet("background: #0284c7; color: white; font-weight: bold; padding: 0 12px; border-radius: 6px;")
        btn_mail_ayar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_mail_ayar.clicked.connect(self.email_ayarlari_ac)
        baslik_layout.addWidget(btn_mail_ayar)
        btn_ice_aktar = QPushButton("📥 İçe Aktar")
        btn_ice_aktar.setFixedHeight(36)
        btn_ice_aktar.setStyleSheet("background: #0284c7; color: white; font-weight: bold; padding: 0 12px; border-radius: 6px;")
        btn_ice_aktar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ice_aktar.clicked.connect(self.finans_evraklarini_ice_aktar)
        baslik_layout.addWidget(btn_ice_aktar)
        btn_toplu = QPushButton("📁 Toplu Çek Yükle")
        btn_toplu.setFixedHeight(36)
        btn_toplu.setStyleSheet("background: #8b5cf6; color: white; font-weight: bold; padding: 0 12px; border-radius: 6px;")
        btn_toplu.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_toplu.clicked.connect(self.toplu_evrak_yukle)
        baslik_layout.addWidget(btn_toplu)
        btn_yeni = QPushButton("+ Tek Evrak Gir")
        btn_yeni.setFixedHeight(36)
        btn_yeni.setStyleSheet("background: #0284c7; color: white; font-weight: bold; padding: 0 12px; border-radius: 6px;")
        btn_yeni.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yeni.clicked.connect(self.yeni_evrak_ekle)
        baslik_layout.addWidget(btn_yeni)
        layout.addLayout(baslik_layout)
        kpi_layout = QHBoxLayout()
        self.lbl_bekleyen_cek = QLabel("BEKLEYEN ÇEK:<br>0,00 ₺<br><span style='font-size:10px; color:#64748b;'>(SIFIR TÜRK LİRASI)</span>")
        self.lbl_bekleyen_senet = QLabel("BEKLEYEN SENET:<br>0,00 ₺<br><span style='font-size:10px; color:#64748b;'>(SIFIR TÜRK LİRASI)</span>")
        self.lbl_yaklasan = QLabel("İLK 30 GÜN TAHSİLAT:<br>0,00 ₺<br><span style='font-size:10px; color:#64748b;'>(SIFIR TÜRK LİRASI)</span>")
        for lbl in (self.lbl_bekleyen_cek, self.lbl_bekleyen_senet, self.lbl_yaklasan):
            lbl.setTextFormat(Qt.TextFormat.RichText)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 12px; font-weight: bold; font-size: 13px; color: #0f172a;")
            kpi_layout.addWidget(lbl)
        layout.addLayout(kpi_layout)

        lbl_projeksiyon_baslik = QLabel("📅 Gelecek 12 Aylık Tahsilat Projeksiyonu (Önümüzdeki 1 Yıl)")
        lbl_projeksiyon_baslik.setStyleSheet("font-size: 13px; font-weight: bold; color: #334155; margin-top: 6px;")
        layout.addWidget(lbl_projeksiyon_baslik)

        self.tbl_aylik_projeksiyon = QTableWidget()
        self.tbl_aylik_projeksiyon.setColumnCount(12)
        self.tbl_aylik_projeksiyon.setRowCount(1)
        self.tbl_aylik_projeksiyon.setFixedHeight(65)
        self.tbl_aylik_projeksiyon.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_aylik_projeksiyon.verticalHeader().setVisible(False)
        self.tbl_aylik_projeksiyon.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl_aylik_projeksiyon.setStyleSheet(
            "QTableWidget { background: #f8fafc; border: 1px solid #cbd5e1; font-size: 11px; } "
            "QHeaderView::section { background: #e2e8f0; font-weight: bold; color: #1e293b; border: none; padding: 4px; }"
        )
        layout.addWidget(self.tbl_aylik_projeksiyon)

        self.tbl_finans = QTableWidget()
        self.tbl_finans.setColumnCount(10)
        self.tbl_finans.setHorizontalHeaderLabels(["SEÇ", "ID", "Tip", "Vade Tarihi", "Kalan Gün", "Tutar", "Banka", "Sahibi", "Durum", "İŞLEM"])
        self.tbl_finans.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_finans.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.tbl_finans.setColumnWidth(0, 50)
        self.tbl_finans.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeMode.Fixed)
        self.tbl_finans.setColumnWidth(9, 140)
        self.tbl_finans.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl_finans.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tbl_finans.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl_finans.cellDoubleClicked.connect(self.finans_satir_cift_tikla)
        layout.addWidget(self.tbl_finans)
        self.finans_tablosunu_guncelle()
        return page

    def email_ayarlari_ac(self):
        dialog = EmailAyarlariDialog(self)
        dialog.exec()

    def finans_evraklarini_disari_aktar(self):
        dosya_yolu, _ = QFileDialog.getSaveFileName(
            self,
            "Çek ve Senet Portföyünü Kaydet",
            "cek_senet_portfoy.csv",
            "CSV Dosyaları (*.csv);;Tüm Dosyalar (*.*)",
        )
        if not dosya_yolu:
            return
        try:
            siralama_secimi = self.cmb_siralama.currentText() if hasattr(self, "cmb_siralama") else "📅 Vade Tarihi (Yakından Uzağa)"
            sql_order = "ORDER BY vade_tarihi ASC"
            if "Uzaktan Yakına" in siralama_secimi:
                sql_order = "ORDER BY vade_tarihi DESC"
            elif "Yüksekten Düşüğe" in siralama_secimi:
                sql_order = "ORDER BY tutar DESC"
            elif "Düşükten Yükseğe" in siralama_secimi:
                sql_order = "ORDER BY tutar ASC"
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            c.execute(
                f"""SELECT evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi_kesideci, sehir, durum, eklenme_tarihi
                    FROM finans_evraklar {sql_order}"""
            )
            kayitlar = c.fetchall()
            conn.close()
            genel_toplam = 0.0
            aylik_dagilim = defaultdict(float)
            ay_isimleri = {
                "01": "Ocak", "02": "Şubat", "03": "Mart", "04": "Nisan",
                "05": "Mayıs", "06": "Haziran", "07": "Temmuz", "08": "Ağustos",
                "09": "Eylül", "10": "Ekim", "11": "Kasım", "12": "Aralık",
            }

            def csv_tl(deger):
                try:
                    return f"{float(deger or 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except (TypeError, ValueError):
                    return "0,00"

            with open(dosya_yolu, mode="w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Sıra", "Evrak Tipi", "Vade Tarihi", "Tutar (TL)", "Banka Adı",
                    "Keşideci / Sahibi", "Şehir", "Durum", "Sisteme Eklenme Tarihi",
                ])
                for idx, satir in enumerate(kayitlar, 1):
                    tip, vade, tutar, banka, sahip, sehir, durum, eklenme = satir
                    tutar_val = float(tutar or 0.0)
                    genel_toplam += tutar_val
                    ham_vade = str(vade or "")[:10]
                    if ham_vade and "-" in ham_vade:
                        parcalar = ham_vade.split("-")
                        if len(parcalar) == 3:
                            yil, ay = parcalar[0], parcalar[1]
                            aylik_dagilim[f"{ay_isimleri.get(ay, ay)} {yil}"] += tutar_val
                    formatli_vade = vade
                    if ham_vade and "-" in ham_vade:
                        p = ham_vade.split("-")
                        if len(p) == 3:
                            formatli_vade = f"{p[2]}.{p[1]}.{p[0]}"
                    writer.writerow([
                        idx, tip, formatli_vade, csv_tl(tutar_val),
                        banka or "", sahip or "", sehir or "", durum or "", eklenme or "",
                    ])
                writer.writerow([])
                writer.writerow(["================ FİNANSAL ÖZET ================"])
                writer.writerow(["GENEL TOPLAM ALACAK:", csv_tl(genel_toplam) + " TL"])
                writer.writerow([])
                writer.writerow(["--- AYLARA GÖRE TAHSİLAT PLANI ---"])
                writer.writerow(["Ay / Yıl", "Toplam Tutar (TL)"])
                for ay_yil, tutar in aylik_dagilim.items():
                    writer.writerow([ay_yil, csv_tl(tutar) + " TL"])
            QMessageBox.information(
                self, "Başarılı",
                f"Çek ve senet portföyü aylık döküm özetleriyle birlikte dışarı aktarıldı:\n{dosya_yolu}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarma sırasında hata oluştu:\n{str(e)}")

    def finans_evraklarini_ice_aktar(self):
        dosya_yolu, _ = QFileDialog.getOpenFileName(
            self,
            "Çek ve Senet Portföyünü İçe Aktar",
            "",
            "CSV Dosyaları (*.csv);;Tüm Dosyalar (*.*)",
        )
        if not dosya_yolu:
            return

        def tutar_oku(ham):
            s = str(ham or "").replace("₺", "").strip()
            if not s:
                return 0.0
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            return float(s)

        try:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            eklenen_sayisi = 0
            with open(dosya_yolu, mode="r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    conn.close()
                    QMessageBox.warning(self, "Hata", "Seçilen dosya boş veya geçersiz formatta.")
                    return
                for row in reader:
                    if len(row) < 4:
                        continue
                    if len(row) >= 8:
                        evrak_tipi = row[1].strip() or "ÇEK"
                        vade_tarihi = row[2].strip()
                        if "." in vade_tarihi:
                            p = vade_tarihi.split(".")
                            if len(p) == 3:
                                vade_tarihi = f"{p[2]}-{p[1]}-{p[0]}"
                        try:
                            tutar = tutar_oku(row[3])
                        except ValueError:
                            tutar = 0.0
                        banka_adi = row[4].strip()
                        sahibi = row[5].strip()
                        sehir = row[6].strip()
                        durum = row[7].strip() or "BEKLİYOR"
                    else:
                        evrak_tipi = row[0].strip() or "ÇEK"
                        vade_tarihi = row[1].strip()
                        if "." in vade_tarihi:
                            p = vade_tarihi.split(".")
                            if len(p) == 3:
                                vade_tarihi = f"{p[2]}-{p[1]}-{p[0]}"
                        try:
                            tutar = tutar_oku(row[2])
                        except ValueError:
                            tutar = 0.0
                        banka_adi = row[3].strip() if len(row) > 3 else ""
                        sahibi = row[4].strip() if len(row) > 4 else ""
                        sehir = ""
                        durum = "BEKLİYOR"
                    customer_id = None
                    if sahibi:
                        c.execute("SELECT id, debt, payment FROM customers WHERE UPPER(name) = ?", (sahip.upper(),))
                        m_row = c.fetchone()
                        if m_row:
                            customer_id, m_debt, m_pay = m_row
                            m_pay = (m_pay or 0.0) + tutar
                            m_rem = max(0.0, (m_debt or 0.0) - m_pay)
                            c.execute(
                                "UPDATE customers SET payment = ?, remaining_debt = ? WHERE id = ?",
                                (m_pay, m_rem, customer_id),
                            )
                    c.execute(
                        """INSERT INTO finans_evraklar
                           (evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi_kesideci, sehir, durum, customer_id)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi, sehir, durum, customer_id),
                    )
                    eklenen_sayisi += 1
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Başarılı", f"Dosyadan {eklenen_sayisi} adet evrak başarıyla içe aktarıldı.")
            self.finans_tablosunu_guncelle()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"İçe aktarma sırasında bir hata oluştu:\n{str(e)}")

    def yeni_evrak_ekle(self):
        dialog = EvrakEkleDialog(self)
        if dialog.exec():
            tip = dialog.cmb_tip.currentText()
            vade = dialog.dt_vade.date().toString("yyyy-MM-dd")
            tutar = turk_para_coz(dialog.inp_tutar.text())
            sahip = buyuk_harf(dialog.inp_sahip.text().strip())
            banka = dialog.cmb_banka.currentText()
            sehir = buyuk_harf(dialog.inp_sehir.text().strip())
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            customer_id = None
            if sahip:
                c.execute("SELECT id, debt, payment FROM customers WHERE UPPER(name) = ?", (sahip,))
                m_row = c.fetchone()
                if m_row:
                    customer_id = m_row[0]
                    m_debt = m_row[1] or 0.0
                    m_pay = (m_row[2] or 0.0) + tutar
                    m_rem = max(0.0, m_debt - m_pay)
                    c.execute(
                        "UPDATE customers SET payment = ?, remaining_debt = ? WHERE id = ?",
                        (m_pay, m_rem, customer_id),
                    )
            c.execute(
                """INSERT INTO finans_evraklar
                   (evrak_tipi, vade_tarihi, tutar, sahibi_kesideci, banka_adi, sehir, customer_id, foto_yolu)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (tip, vade, tutar, sahip, banka, sehir, customer_id, getattr(dialog, "foto_yolu", None)),
            )
            conn.commit()
            conn.close()
            self.finans_tablosunu_guncelle()
            vade_hatirlatici_mailleri_gonder()
            if customer_id:
                QMessageBox.information(
                    self, "Cari Hesap Güncellendi",
                    f"Evrak kaydedildi ve '{sahip}' adlı müşterinin cari hesabından {tutar:,.2f} ₺ düşüldü.",
                )
            else:
                QMessageBox.information(
                    self, "Başarılı",
                    "Evrak sisteme başarıyla kaydedildi (Eşleşen müşteri bulunamadığı için cari hesaba yansıtılmadı).",
                )

    def kucuk_resim_widgeti(self, resim_yolu, boyut=60):
        lbl = QLabel()
        lbl.setFixedSize(boyut, boyut)
        lbl.setCursor(Qt.CursorShape.PointingHandCursor)
        if resim_yolu and os.path.exists(resim_yolu):
            pix = QPixmap(resim_yolu).scaled(boyut, boyut, Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation)
            lbl.setPixmap(pix)
            lbl.mousePressEvent = lambda e, yol=resim_yolu: self._resmi_buyuk_goster(yol)
        else:
            lbl.setText("Yok")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("color:#94a3b8; border:1px dashed #cbd5e1;")
        return lbl

    def _resmi_buyuk_goster(self, resim_yolu):
        dlg = QDialog(self)
        dlg.setWindowTitle("Fotoğraf")
        v = QVBoxLayout(dlg)
        lbl = QLabel()
        pix = QPixmap(resim_yolu).scaled(700, 700, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        lbl.setPixmap(pix)
        v.addWidget(lbl)
        dlg.exec()

    def finans_tablosunu_guncelle(self):
        siralama_secimi = self.cmb_siralama.currentText() if hasattr(self, "cmb_siralama") else "📅 Vade Tarihi (Yakından Uzağa)"
        sql_order = "ORDER BY vade_tarihi ASC"
        if "Uzaktan Yakına" in siralama_secimi:
            sql_order = "ORDER BY vade_tarihi DESC"
        elif "Yüksekten Düşüğe" in siralama_secimi:
            sql_order = "ORDER BY tutar DESC"
        elif "Düşükten Yükseğe" in siralama_secimi:
            sql_order = "ORDER BY tutar ASC"
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(
            f"SELECT id, evrak_tipi, vade_tarihi, tutar, banka_adi, sahibi_kesideci, durum, foto_yolu "
            f"FROM finans_evraklar WHERE durum = 'BEKLİYOR' {sql_order}"
        )
        kayitlar = c.fetchall()
        conn.close()

        bugun_dt = datetime.now()
        aylar_keys = []
        aylik_toplamlar = {}
        ay_isimleri_tr = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
        for i in range(12):
            m = bugun_dt.month + i
            y = bugun_dt.year + (m - 1) // 12
            m = (m - 1) % 12 + 1
            key = f"{y}-{m:02d}"
            aylar_keys.append(key)
            aylik_toplamlar[key] = 0.0
        if hasattr(self, "tbl_aylik_projeksiyon"):
            self.tbl_aylik_projeksiyon.setHorizontalHeaderLabels(
                [f"{ay_isimleri_tr[int(k.split('-')[1])]} {k.split('-')[0]}" for k in aylar_keys]
            )
            for _, _, vade, tutar, _, _, _, _ in kayitlar:
                if vade and "-" in str(vade):
                    parcalar = str(vade)[:10].split("-")
                    if len(parcalar) == 3:
                        yil_ay = f"{parcalar[0]}-{parcalar[1]}"
                        if yil_ay in aylik_toplamlar:
                            aylik_toplamlar[yil_ay] += float(tutar or 0.0)
            for col_idx, key in enumerate(aylar_keys):
                tutar_deger = aylik_toplamlar[key]
                item = QTableWidgetItem(format_tl(tutar_deger) if tutar_deger > 0 else "-")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if tutar_deger > 0:
                    item.setForeground(QColor("#0369a1"))
                    item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                else:
                    item.setForeground(QColor("#94a3b8"))
                self.tbl_aylik_projeksiyon.setItem(0, col_idx, item)

        self.tbl_finans.setRowCount(len(kayitlar))
        t_cek = t_senet = bu_ay = 0.0
        bugun = bugun_dt.date()
        uyarilar = []
        for r, (eid, tip, vade, tutar, banka, sahip, durum, foto_yolu) in enumerate(kayitlar):
            try:
                vade_dt = datetime.strptime(vade, "%Y-%m-%d").date() if vade else bugun
            except ValueError:
                vade_dt = bugun
            kalan_gun = (vade_dt - bugun).days
            if tip == "ÇEK":
                t_cek += float(tutar or 0)
            elif tip == "SENET":
                t_senet += float(tutar or 0)
            if 0 <= kalan_gun <= 30:
                bu_ay += float(tutar or 0)
            if 0 <= kalan_gun <= 7:
                uyarilar.append(f"• {sahip} - {float(tutar or 0):,.2f} ₺ ({kalan_gun} gün kaldı)")
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            chk_item.setCheckState(Qt.CheckState.Unchecked)
            chk_item.setData(Qt.ItemDataRole.UserRole, eid)
            self.tbl_finans.setItem(r, 0, chk_item)
            self.tbl_finans.setItem(r, 1, QTableWidgetItem(str(eid)))
            self.tbl_finans.setItem(r, 2, QTableWidgetItem(tip))
            self.tbl_finans.setItem(r, 3, QTableWidgetItem(vade_dt.strftime("%d.%m.%Y")))
            item_kalan = QTableWidgetItem(f"{kalan_gun} Gün")
            if kalan_gun < 0:
                item_kalan.setForeground(QColor("red"))
                item_kalan.setText("VADESİ GEÇTİ")
            elif kalan_gun <= 7:
                item_kalan.setForeground(QColor("#f59e0b"))
            self.tbl_finans.setItem(r, 4, item_kalan)
            self.tbl_finans.setItem(r, 5, QTableWidgetItem(format_tl(tutar)))
            self.tbl_finans.setItem(r, 6, QTableWidgetItem(str(banka or "-")))
            self.tbl_finans.setItem(r, 7, QTableWidgetItem(str(sahip or "-")))
            self.tbl_finans.setItem(r, 8, QTableWidgetItem(durum))
            btn_duzenle = QPushButton("✏️ Düzenle")
            btn_duzenle.setFixedHeight(28)
            btn_duzenle.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_duzenle.setStyleSheet("background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; border-radius: 4px; font-weight: bold;")
            btn_duzenle.clicked.connect(lambda _, _id=eid: self.evrak_duzenle(_id))
            w_widget = QWidget()
            w_lay = QHBoxLayout(w_widget)
            w_lay.setContentsMargins(4, 2, 4, 2)
            w_lay.addWidget(self.kucuk_resim_widgeti(foto_yolu, 36))
            w_lay.addWidget(btn_duzenle)
            self.tbl_finans.setCellWidget(r, 9, w_widget)
        self.lbl_bekleyen_cek.setText(
            f"BEKLEYEN ÇEK:<br>{format_tl(t_cek)}<br>"
            f"<span style='font-size:9.5px; color:#64748b; font-weight:normal;'>({para_yazisi_olustur(t_cek)})</span>"
        )
        self.lbl_bekleyen_senet.setText(
            f"BEKLEYEN SENET:<br>{format_tl(t_senet)}<br>"
            f"<span style='font-size:9.5px; color:#64748b; font-weight:normal;'>({para_yazisi_olustur(t_senet)})</span>"
        )
        self.lbl_yaklasan.setText(
            f"İLK 30 GÜN TAHSİLAT:<br>{format_tl(bu_ay)}<br>"
            f"<span style='font-size:9.5px; color:#64748b; font-weight:normal;'>({para_yazisi_olustur(bu_ay)})</span>"
        )
        if uyarilar and not getattr(self, "ilk_uyari_yapildi", False) and self.isVisible():
            self.ilk_uyari_yapildi = True
            QMessageBox.warning(self, "🚨 VADE HATIRLATICISI", "Vadesine 7 günden az kalmış evraklarınız var:\n\n" + "\n".join(uyarilar))

    def finans_satir_cift_tikla(self, r, _c):
        item = self.tbl_finans.item(r, 0)
        if item:
            eid = item.data(Qt.ItemDataRole.UserRole)
            if eid:
                self.evrak_duzenle(eid)

    def evrak_duzenle(self, evrak_id):
        dialog = EvrakEkleDialog(parent=self, evrak_id=evrak_id)
        if dialog.exec():
            tip = dialog.cmb_tip.currentText()
            vade = dialog.dt_vade.date().toString("yyyy-MM-dd")
            tutar = turk_para_coz(dialog.inp_tutar.text())
            sahip = buyuk_harf(dialog.inp_sahip.text().strip())
            banka = dialog.cmb_banka.currentText()
            sehir = buyuk_harf(dialog.inp_sehir.text().strip())
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            c.execute(
                """UPDATE finans_evraklar
                   SET evrak_tipi = ?, vade_tarihi = ?, tutar = ?, sahibi_kesideci = ?, banka_adi = ?, sehir = ?
                   WHERE id = ?""",
                (tip, vade, tutar, sahip, banka, sehir, evrak_id),
            )
            conn.commit()
            conn.close()
            self.finans_tablosunu_guncelle()
            QMessageBox.information(self, "Başarılı", "Evrak bilgileri güncellendi.")

    def finans_secilenleri_sil(self):
        ids = []
        for r in range(self.tbl_finans.rowCount()):
            item = self.tbl_finans.item(r, 0)
            if item and item.checkState() == Qt.CheckState.Checked:
                ids.append(item.data(Qt.ItemDataRole.UserRole))
        if not ids:
            QMessageBox.information(self, "Seçim Yok", "Silmek için satırdaki kutucuğu işaretleyin.")
            return
        cevap = QMessageBox.question(
            self, "Seçilenleri Sil",
            f"{len(ids)} evrak silinsin mi?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if cevap != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.executemany("DELETE FROM finans_evraklar WHERE id = ?", [(i,) for i in ids])
        conn.commit()
        conn.close()
        self.finans_tablosunu_guncelle()

    def finans_tumunu_sil(self):
        cevap = QMessageBox.question(
            self, "Tümünü Sil",
            "Bekleyen tüm çek ve senetler silinsin mi? Bu işlem geri alınamaz.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if cevap != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM finans_evraklar WHERE durum = 'BEKLİYOR'")
        conn.commit()
        conn.close()
        self.finans_tablosunu_guncelle()

    def toplu_evrak_yukle(self):
        dosyalar, _ = QFileDialog.getOpenFileNames(
            self,
            "Toplu Çek Fotoğrafları Seç (Birden fazla seçebilirsiniz)",
            "",
            "Resim Dosyaları (*.png *.jpg *.jpeg)",
        )
        if not dosyalar:
            return
        toplam = len(dosyalar)
        self.progress_dialog = QProgressDialog(
            "Sıfır hata kontrolü yapılıyor ve çekler okunuyor...", "İptal", 0, toplam, self
        )
        self.progress_dialog.setWindowTitle("Toplu Evrak Okuma")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        self.toplu_thread = TopluEvrakOkuyucuThread(dosyalar)
        self.progress_dialog.canceled.connect(self.toplu_thread.requestInterruption)
        self.toplu_thread.progress_guncelle.connect(self.toplu_islem_guncelle)
        self.toplu_thread.islem_bitti.connect(self.toplu_islem_tamamlandi)
        self.toplu_thread.start()

    def toplu_islem_guncelle(self, mevcut, toplam, dosya_adi):
        self.progress_dialog.setValue(mevcut)
        self.progress_dialog.setLabelText(f"İşlenen Çek: {mevcut} / {toplam}\nDosya: {dosya_adi}")

    def toplu_islem_tamamlandi(self, basarili_sayisi, hatalar):
        self.progress_dialog.close()
        self.finans_tablosunu_guncelle()
        if hatalar:
            dialog = TopluSonucDialog(basarili_sayisi, hatalar, self)
            dialog.exec()
        else:
            QMessageBox.information(
                self, "Kusursuz İşlem",
                f"Seçilen {basarili_sayisi} evrakın tamamı sıfır hata ile doğrulandı ve portföye kaydedildi.",
            )

    def _golge_ekle(self, widget, bulaniklik=20, y_offset=3, opaklik=40):
        golge = QGraphicsDropShadowEffect(widget)
        golge.setBlurRadius(bulaniklik)
        golge.setOffset(0, y_offset)
        golge.setColor(QColor(15, 23, 42, opaklik))
        widget.setGraphicsEffect(golge)
        return widget

    def _sayfa_basligi_widget(self, ikon, metin, renk="#0284c7"):
        kutu = QWidget()
        v = QVBoxLayout(kutu)
        v.setContentsMargins(0, 0, 0, 4)
        v.setSpacing(8)
        ust = QHBoxLayout()
        ust.setSpacing(12)
        rozet = QLabel(ikon)
        rozet.setFixedSize(38, 38)
        rozet.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rozet.setStyleSheet(f"background-color: {renk}1a; border-radius: 10px; font-size: 18px;")
        ust.addWidget(rozet)
        baslik = QLabel(metin)
        baslik.setStyleSheet("font-size: 21px; font-weight: bold; color: #0f172a;")
        ust.addWidget(baslik)
        ust.addStretch()
        v.addLayout(ust)
        cizgi = QFrame()
        cizgi.setFixedHeight(3)
        cizgi.setStyleSheet(f"""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 {renk}, stop:1 transparent);
            border-radius: 1px;
        """)
        v.addWidget(cizgi)
        return kutu

    def _rapor_ozet_karti(self, baslik, renk="#0284c7"):
        kutu = QFrame()
        kutu.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 8px;")
        kutu_layout = QVBoxLayout(kutu)
        kutu_layout.setContentsMargins(14, 10, 14, 10)
        lbl_baslik = QLabel(baslik)
        lbl_baslik.setStyleSheet("color: #64748b; font-size: 12px; font-weight: bold;")
        lbl_deger = QLabel("₺ 0.00")
        lbl_deger.setStyleSheet(f"color: {renk}; font-size: 20px; font-weight: bold;")
        kutu_layout.addWidget(lbl_baslik)
        kutu_layout.addWidget(lbl_deger)
        kutu.lbl_deger = lbl_deger
        return self._golge_ekle(kutu, bulaniklik=16, y_offset=2, opaklik=25)

    def _tablo_grafik_govdesi_kur(self, layout, tablo, grafik, tur="bar"):
        govde = QHBoxLayout()
        govde.setSpacing(10)
        tablo.setMinimumWidth(260)
        grafik.setMinimumHeight(420)
        grafik2 = GoogleStilGrafik(tur=tur)
        grafik2.setMinimumHeight(420)
        govde.addWidget(tablo, 6)
        sag = QVBoxLayout()
        sag.setSpacing(8)
        sag.addWidget(grafik)
        sag.addWidget(grafik2)
        govde.addLayout(sag, 5)
        layout.addLayout(govde)
        return grafik2

    @staticmethod
    def _tarihi_qdate_yap(created_at):
        if not created_at:
            return None
        ham = str(created_at).strip().split(" ")[0]
        for fmt in ("yyyy-MM-dd", "dd.MM.yyyy", "dd/MM/yyyy", "d.M.yyyy"):
            qd = QDate.fromString(ham, fmt)
            if qd.isValid():
                return qd
        return None

    def _build_gunluk_rapor_sayfasi(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("📅", "GÜNLÜK RAPOR", "#0284c7"))
        satir = QHBoxLayout()
        self.dt_gunluk_tarih = QDateEdit(QDate.currentDate())
        self.dt_gunluk_tarih.setCalendarPopup(True)
        self.dt_gunluk_tarih.setFixedHeight(36)
        satir.addWidget(QLabel("Tarih:"))
        satir.addWidget(self.dt_gunluk_tarih)
        self.cmb_gunluk_odeme = QComboBox()
        self.cmb_gunluk_odeme.addItems(["TÜMÜ", "NAKİT", "POS", "AÇIK HESAP", "PARÇALI", "DİĞER"])
        satir.addWidget(QLabel("Ödeme Tipi:"))
        satir.addWidget(self.cmb_gunluk_odeme)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.gunluk_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        orta_bolum = QHBoxLayout()
        self.tbl_gunluk_rapor = QTableWidget()
        self.tbl_gunluk_rapor.setColumnCount(7)
        self.tbl_gunluk_rapor.setHorizontalHeaderLabels(["Tarih", "Personel", "Müşteri", "Ürün", "Miktar", "Tutar", "Ödeme Tipi"])
        self.tbl_gunluk_rapor.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_gunluk_rapor.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        orta_bolum.addWidget(self.tbl_gunluk_rapor, 3)
        self.chart_gunluk = DinamikRaporGrafigi()
        orta_bolum.addWidget(self.chart_gunluk, 2)
        layout.addLayout(orta_bolum)
        ozet = QHBoxLayout()
        self.kart_gunluk_nakit = self._rapor_ozet_karti("NAKİT", "#16a34a")
        self.kart_gunluk_pos = self._rapor_ozet_karti("POS", "#0284c7")
        self.kart_gunluk_acik = self._rapor_ozet_karti("AÇIK HESAP", "#dc2626")
        self.kart_gunluk_toplam = self._rapor_ozet_karti("TOPLAM CİRO", "#7c3aed")
        for k in (self.kart_gunluk_nakit, self.kart_gunluk_pos, self.kart_gunluk_acik, self.kart_gunluk_toplam):
            ozet.addWidget(k)
        layout.addLayout(ozet)
        self.grafik_gunluk = GoogleStilGrafik(tur="bar")
        layout.addWidget(self.grafik_gunluk)
        return page

    def gunluk_rapor_yenile(self):
        secilen = self.dt_gunluk_tarih.date()
        odeme_filtre = self.cmb_gunluk_odeme.currentText()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT created_at, personnel_name, customer_name, product_name, qty, total_amount, payment_type
                     FROM sales_history ORDER BY id DESC""")
        kayitlar = c.fetchall()
        conn.close()
        satirlar = []
        for tarih, per, m_ad, urun, qty, tutar, pay_type in kayitlar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd != secilen:
                continue
            if odeme_filtre != "TÜMÜ" and (pay_type or "") != odeme_filtre:
                continue
            satirlar.append((tarih, per, m_ad, urun, qty, tutar, pay_type))
        self.tbl_gunluk_rapor.setRowCount(len(satirlar))
        nakit = pos = acik = 0.0
        for r, (tarih, per, m_ad, urun, qty, tutar, pay_type) in enumerate(satirlar):
            self.tbl_gunluk_rapor.setItem(r, 0, QTableWidgetItem(str(tarih or "-")))
            self.tbl_gunluk_rapor.setItem(r, 1, QTableWidgetItem(str(per or "-")))
            self.tbl_gunluk_rapor.setItem(r, 2, QTableWidgetItem(str(m_ad or "-")))
            self.tbl_gunluk_rapor.setItem(r, 3, QTableWidgetItem(str(urun or "-")))
            self.tbl_gunluk_rapor.setItem(r, 4, QTableWidgetItem(f"{float(qty or 0):,.1f} KG"))
            self.tbl_gunluk_rapor.setItem(r, 5, QTableWidgetItem(f"{float(tutar or 0):,.2f} ₺"))
            self.tbl_gunluk_rapor.setItem(r, 6, QTableWidgetItem(str(pay_type or "-")))
            t = float(tutar or 0)
            if pay_type == "NAKİT": nakit += t
            elif pay_type == "POS": pos += t
            elif pay_type == "AÇIK HESAP": acik += t
        self.kart_gunluk_nakit.lbl_deger.setText(f"₺ {nakit:,.2f}")
        self.kart_gunluk_pos.lbl_deger.setText(f"₺ {pos:,.2f}")
        self.kart_gunluk_acik.lbl_deger.setText(f"₺ {acik:,.2f}")
        self.kart_gunluk_toplam.lbl_deger.setText(f"₺ {(nakit + pos + acik):,.2f}")
        self.chart_gunluk.veri_yukle(
            satirlar,
            col_map={"tarih": 0, "per": 1, "musteri": 2, "urun": 3, "kg": 4, "tutar": 5, "odeme": 6},
        )
        self.grafik_gunluk.veri_ayarla(["NAKİT", "POS", "AÇIK HESAP"], [nakit, pos, acik], "Ödeme Tipine Göre Dağılım")

    def _build_tarihsel_rapor_sayfasi(self):
        page = QWidget()
        dis_layout = QVBoxLayout(page)
        dis_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        icerik_widget = QWidget()
        layout = QVBoxLayout(icerik_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)
        layout.addWidget(self._sayfa_basligi_widget("🗓", "TARİHSEL RAPOR", "#7c3aed"))
        satir = QHBoxLayout()
        self.dt_tarihsel_bas = QDateEdit(QDate.currentDate().addDays(-7))
        self.dt_tarihsel_bas.setCalendarPopup(True)
        self.dt_tarihsel_bas.setFixedHeight(36)
        self.dt_tarihsel_bit = QDateEdit(QDate.currentDate())
        self.dt_tarihsel_bit.setCalendarPopup(True)
        self.dt_tarihsel_bit.setFixedHeight(36)
        satir.addWidget(QLabel("Başlangıç:"))
        satir.addWidget(self.dt_tarihsel_bas)
        satir.addWidget(QLabel("Bitiş:"))
        satir.addWidget(self.dt_tarihsel_bit)
        self.cmb_tarihsel_personel = QComboBox()
        self.cmb_tarihsel_personel.addItem("Tüm Personeller")
        satir.addWidget(QLabel("Personel:"))
        satir.addWidget(self.cmb_tarihsel_personel)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.tarihsel_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        self.tbl_tarihsel_rapor = QTableWidget()
        self.tbl_tarihsel_rapor.setColumnCount(7)
        self.tbl_tarihsel_rapor.setHorizontalHeaderLabels(["Tarih", "Personel", "Müşteri", "Ürün", "Miktar", "Tutar", "Ödeme Tipi"])
        header = self.tbl_tarihsel_rapor.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_tarihsel_rapor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.tbl_tarihsel_rapor.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        orta_bolum = QHBoxLayout()
        self.chart_tarihsel = DinamikRaporGrafigi()
        orta_bolum.addWidget(self.chart_tarihsel)
        layout.addLayout(orta_bolum)
        ozet = QHBoxLayout()
        self.kart_tarihsel_toplam = self._rapor_ozet_karti("TOPLAM CİRO", "#7c3aed")
        self.kart_tarihsel_tahsilat = self._rapor_ozet_karti("TOPLAM TAHSİLAT", "#16a34a")
        self.kart_tarihsel_kalan = self._rapor_ozet_karti("KALAN (AÇIK HESAP)", "#dc2626")
        for k in (self.kart_tarihsel_toplam, self.kart_tarihsel_tahsilat, self.kart_tarihsel_kalan):
            ozet.addWidget(k)
        layout.addLayout(ozet)
        self.grafik_tarihsel = GoogleStilGrafik(tur="cizgi")
        self.grafik_tarihsel.setMinimumHeight(420)
        self.grafik_tarihsel2 = self._tablo_grafik_govdesi_kur(layout, self.tbl_tarihsel_rapor, self.grafik_tarihsel, "cizgi")
        scroll_area.setWidget(icerik_widget)
        dis_layout.addWidget(scroll_area)
        return page

    def tarihsel_rapor_yenile(self):
        bas, bit = self.dt_tarihsel_bas.date(), self.dt_tarihsel_bit.date()
        personel_filtre = self.cmb_tarihsel_personel.currentText()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT created_at, personnel_name, customer_name, product_name, qty,
                            total_amount, payment_type, payment_received
                     FROM sales_history ORDER BY id DESC""")
        kayitlar = c.fetchall()
        c.execute("SELECT DISTINCT name FROM personnel ORDER BY name")
        personel_isimleri = [r[0] for r in c.fetchall()]
        conn.close()
        mevcut = {self.cmb_tarihsel_personel.itemText(i) for i in range(self.cmb_tarihsel_personel.count())}
        for isim in personel_isimleri:
            if isim not in mevcut:
                self.cmb_tarihsel_personel.addItem(isim)
        satirlar = []
        for tarih, per, m_ad, urun, qty, tutar, pay_type, tahsilat in kayitlar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd is None or qd < bas or qd > bit:
                continue
            if personel_filtre != "Tüm Personeller" and (per or "") != personel_filtre:
                continue
            satirlar.append((tarih, per, m_ad, urun, qty, tutar, pay_type, tahsilat))
        self.tbl_tarihsel_rapor.setRowCount(len(satirlar))
        toplam = tahsilat_toplam = 0.0
        for r, (tarih, per, m_ad, urun, qty, tutar, pay_type, tahsilat) in enumerate(satirlar):
            self.tbl_tarihsel_rapor.setItem(r, 0, QTableWidgetItem(str(tarih or "-")))
            self.tbl_tarihsel_rapor.setItem(r, 1, QTableWidgetItem(str(per or "-")))
            self.tbl_tarihsel_rapor.setItem(r, 2, QTableWidgetItem(str(m_ad or "-")))
            self.tbl_tarihsel_rapor.setItem(r, 3, QTableWidgetItem(str(urun or "-")))
            self.tbl_tarihsel_rapor.setItem(r, 4, QTableWidgetItem(f"{float(qty or 0):,.1f} KG"))
            self.tbl_tarihsel_rapor.setItem(r, 5, QTableWidgetItem(f"{float(tutar or 0):,.2f} ₺"))
            self.tbl_tarihsel_rapor.setItem(r, 6, QTableWidgetItem(str(pay_type or "-")))
            toplam += float(tutar or 0)
            tahsilat_toplam += float(tahsilat or 0)
        self.kart_tarihsel_toplam.lbl_deger.setText(f"₺ {toplam:,.2f}")
        self.kart_tarihsel_tahsilat.lbl_deger.setText(f"₺ {tahsilat_toplam:,.2f}")
        self.kart_tarihsel_kalan.lbl_deger.setText(f"₺ {max(0.0, toplam - tahsilat_toplam):,.2f}")
        self.chart_tarihsel.veri_yukle(
            satirlar,
            col_map={"tarih": 0, "per": 1, "musteri": 2, "urun": 3, "kg": 4, "tutar": 5, "odeme": 6},
        )
        gunluk_toplam = defaultdict(float)
        gun_sayisi = bas.daysTo(bit) + 1
        for i in range(gun_sayisi):
            gunluk_toplam[bas.addDays(i).toString("dd.MM")] = 0.0
        for tarih, per, m_ad, urun, qty, tutar, pay_type, tahsilat in satirlar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd:
                gunluk_toplam[qd.toString("dd.MM")] += float(tutar or 0)
        etiketler = list(gunluk_toplam.keys())
        degerler = list(gunluk_toplam.values())
        self.grafik_tarihsel.veri_ayarla(etiketler, degerler, "Günlük Ciro Trendi")
        self.grafik_tarihsel2.veri_ayarla(etiketler, degerler, "Günlük Ciro Trendi")

    def _build_urunsel_rapor_sayfasi(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("📦", "ÜRÜNSEL RAPOR", "#16a34a"))
        satir = QHBoxLayout()
        self.dt_urunsel_bas = QDateEdit(QDate.currentDate().addMonths(-1))
        self.dt_urunsel_bas.setCalendarPopup(True)
        self.dt_urunsel_bas.setFixedHeight(36)
        self.dt_urunsel_bit = QDateEdit(QDate.currentDate())
        self.dt_urunsel_bit.setCalendarPopup(True)
        self.dt_urunsel_bit.setFixedHeight(36)
        satir.addWidget(QLabel("Başlangıç:"))
        satir.addWidget(self.dt_urunsel_bas)
        satir.addWidget(QLabel("Bitiş:"))
        satir.addWidget(self.dt_urunsel_bit)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.urunsel_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        orta = QHBoxLayout()
        self.tbl_urunsel_rapor = QTableWidget()
        self.tbl_urunsel_rapor.setColumnCount(3)
        self.tbl_urunsel_rapor.setHorizontalHeaderLabels([
            "Ürün Adı",
            "Satılan Miktar (KG)",
            "Toplam Satış Tutarı (TL)",
        ])
        header = self.tbl_urunsel_rapor.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_urunsel_rapor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.tbl_urunsel_rapor.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        orta.addWidget(self.tbl_urunsel_rapor, 5)
        self.grafik_urunsel = GoogleStilGrafik(tur="bar")
        orta.addWidget(self.grafik_urunsel, 5)
        layout.addLayout(orta)
        return page

    def urunsel_rapor_yenile(self):
        bas, bit = self.dt_urunsel_bas.date(), self.dt_urunsel_bit.date()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT created_at, personnel_name, customer_name, product_name, qty, total_amount, payment_type
                     FROM sales_history""")
        satislar = c.fetchall()
        conn.close()
        ozet = defaultdict(lambda: {"miktar": 0.0, "tutar": 0.0})
        for tarih, per, m_ad, urun, qty, tutar, pay_type in satislar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd is None or qd < bas or qd > bit:
                continue
            ozet[urun]["miktar"] += float(qty or 0)
            ozet[urun]["tutar"] += float(tutar or 0)
        satirlar = sorted(ozet.items(), key=lambda kv: kv[1]["tutar"], reverse=True)
        self.tbl_urunsel_rapor.setRowCount(len(satirlar))
        for i, (urun, veri) in enumerate(satirlar):
            item_ad = QTableWidgetItem(str(urun or "-"))
            self.tbl_urunsel_rapor.setItem(i, 0, item_ad)
            kg_deger = f"{float(veri['miktar']):,.1f} KG".replace(",", "X").replace(".", ",").replace("X", ".")
            item_kg = QTableWidgetItem(kg_deger)
            item_kg.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.tbl_urunsel_rapor.setItem(i, 1, item_kg)
            tl_deger = f"{float(veri['tutar']):,.2f} ₺".replace(",", "X").replace(".", ",").replace("X", ".")
            item_tl = QTableWidgetItem(tl_deger)
            item_tl.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.tbl_urunsel_rapor.setItem(i, 2, item_tl)
        ilk_10 = satirlar[:10]
        self.grafik_urunsel.veri_ayarla([str(u or "-") for u, _ in ilk_10], [v["tutar"] for _, v in ilk_10], "En Çok Ciro Yapan 10 Ürün")

    def _build_grupsal_rapor_sayfasi(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("🗂", "GRUPSAL RAPOR", "#ea580c"))
        satir = QHBoxLayout()
        self.dt_grupsal_bas = QDateEdit(QDate.currentDate().addMonths(-1))
        self.dt_grupsal_bas.setCalendarPopup(True)
        self.dt_grupsal_bas.setFixedHeight(36)
        self.dt_grupsal_bit = QDateEdit(QDate.currentDate())
        self.dt_grupsal_bit.setCalendarPopup(True)
        self.dt_grupsal_bit.setFixedHeight(36)
        satir.addWidget(QLabel("Başlangıç:"))
        satir.addWidget(self.dt_grupsal_bas)
        satir.addWidget(QLabel("Bitiş:"))
        satir.addWidget(self.dt_grupsal_bit)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.grupsal_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        orta_govde = QHBoxLayout()
        self.tbl_grupsal_rapor = QTableWidget()
        self.tbl_grupsal_rapor.setColumnCount(4)
        self.tbl_grupsal_rapor.setHorizontalHeaderLabels(["Grup Adı", "Toplam Miktar", "Toplam Satış", "Satılan Çeşit"])
        self.tbl_grupsal_rapor.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_grupsal_rapor.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        orta_govde.addWidget(self.tbl_grupsal_rapor, 1)
        self.grafik_grupsal = GoogleStilGrafik(tur="pasta")
        orta_govde.addWidget(self.grafik_grupsal, 3)
        layout.addLayout(orta_govde)
        return page

    def grupsal_rapor_yenile(self):
        bas, bit = self.dt_grupsal_bas.date(), self.dt_grupsal_bit.date()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT g.name, p.name FROM products p
                     JOIN product_groups g ON p.group_id = g.id""")
        urun_grup_map = {}
        for gname, pname in c.fetchall():
            urun_grup_map.setdefault(pname, gname)
        c.execute("SELECT created_at, product_name, qty, total_amount FROM sales_history")
        satislar = c.fetchall()
        conn.close()
        ozet = defaultdict(lambda: {"miktar": 0.0, "tutar": 0.0, "urunler": set()})
        for tarih, urun, qty, tutar in satislar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd is None or qd < bas or qd > bit:
                continue
            grup = urun_grup_map.get(urun, "GRUPSUZLAR")
            ozet[grup]["miktar"] += float(qty or 0)
            ozet[grup]["tutar"] += float(tutar or 0)
            ozet[grup]["urunler"].add(urun)
        satirlar = sorted(ozet.items(), key=lambda kv: kv[1]["tutar"], reverse=True)
        self.tbl_grupsal_rapor.setRowCount(len(satirlar))
        for r, (grup, veri) in enumerate(satirlar):
            self.tbl_grupsal_rapor.setItem(r, 0, QTableWidgetItem(grup))
            self.tbl_grupsal_rapor.setItem(r, 1, QTableWidgetItem(f"{veri['miktar']:,.1f} KG"))
            self.tbl_grupsal_rapor.setItem(r, 2, QTableWidgetItem(f"{veri['tutar']:,.2f} ₺"))
            self.tbl_grupsal_rapor.setItem(r, 3, QTableWidgetItem(str(len(veri["urunler"]))))
        self.grafik_grupsal.veri_ayarla([g for g, _ in satirlar], [v["tutar"] for _, v in satirlar], "Gruplara Göre Satış Dağılımı")

    def _build_stok_hareket_sayfasi(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("📉", "STOK HAREKET RAPORU", "#dc2626"))
        bilgi = QLabel("Bu rapor satış geçmişinden türetilir (her satış = stoktan düşüş).")
        bilgi.setStyleSheet("color: #64748b; font-size: 12px;")
        layout.addWidget(bilgi)
        satir = QHBoxLayout()
        self.dt_stok_bas = QDateEdit(QDate.currentDate().addMonths(-1))
        self.dt_stok_bas.setCalendarPopup(True)
        self.dt_stok_bas.setFixedHeight(36)
        self.dt_stok_bit = QDateEdit(QDate.currentDate())
        self.dt_stok_bit.setCalendarPopup(True)
        self.dt_stok_bit.setFixedHeight(36)
        satir.addWidget(QLabel("Başlangıç:"))
        satir.addWidget(self.dt_stok_bas)
        satir.addWidget(QLabel("Bitiş:"))
        satir.addWidget(self.dt_stok_bit)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.stok_hareket_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        self.tbl_stok_hareket = QTableWidget()
        self.tbl_stok_hareket.setColumnCount(3)
        self.tbl_stok_hareket.setHorizontalHeaderLabels(["Ürün Adı", "Toplam Miktar", "Satış Sayısı"])
        self.tbl_stok_hareket.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_stok_hareket.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tbl_stok_hareket)
        return page

    def stok_hareket_rapor_yenile(self):
        bas, bit = self.dt_stok_bas.date(), self.dt_stok_bit.date()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT product_name, qty, created_at FROM sales_history")
        kayitlar = c.fetchall()
        conn.close()
        ozet = defaultdict(lambda: {"miktar": 0.0, "sayi": 0})
        for urun, qty, tarih in kayitlar:
            qd = self._tarihi_qdate_yap(tarih)
            if qd is None or qd < bas or qd > bit:
                continue
            ozet[urun or "-"]["miktar"] += float(qty or 0)
            ozet[urun or "-"]["sayi"] += 1
        satirlar = sorted(ozet.items(), key=lambda kv: kv[1]["miktar"])
        self.tbl_stok_hareket.setRowCount(len(satirlar))
        for r, (urun, veri) in enumerate(satirlar):
            self.tbl_stok_hareket.setItem(r, 0, QTableWidgetItem(str(urun)))
            miktar_item = QTableWidgetItem(f"{veri['miktar']:,.1f} KG")
            miktar_item.setForeground(QColor("#dc2626"))
            self.tbl_stok_hareket.setItem(r, 1, miktar_item)
            self.tbl_stok_hareket.setItem(r, 2, QTableWidgetItem(str(veri["sayi"])))

    def _build_personel_hareket_sayfasi(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("👤", "PERSONEL HAREKET RAPORU", "#0369a1"))
        satir = QHBoxLayout()
        self.cmb_ph_personel = QComboBox()
        self.cmb_ph_personel.addItem("Tüm Personeller")
        satir.addWidget(QLabel("Personel:"))
        satir.addWidget(self.cmb_ph_personel)
        btn = QPushButton("🔍 Listele")
        btn.setFixedHeight(36)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 5px; padding: 0 16px;")
        btn.clicked.connect(self.personel_hareket_rapor_yenile)
        satir.addWidget(btn)
        satir.addStretch()
        layout.addLayout(satir)
        self.tbl_personel_hareket = QTableWidget()
        self.tbl_personel_hareket.setColumnCount(6)
        self.tbl_personel_hareket.setHorizontalHeaderLabels(["Tarih", "Personel", "İşlem", "Müşteri", "Ürün", "Tutar"])
        self.tbl_personel_hareket.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_personel_hareket.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tbl_personel_hareket)
        return page

    def personel_hareket_rapor_yenile(self):
        personel_filtre = self.cmb_ph_personel.currentText()
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT created_at, personnel_name, customer_name, product_name, total_amount
                     FROM sales_history ORDER BY id DESC""")
        kayitlar = c.fetchall()
        c.execute("SELECT DISTINCT name FROM personnel ORDER BY name")
        personel_isimleri = [r[0] for r in c.fetchall()]
        conn.close()
        mevcut = {self.cmb_ph_personel.itemText(i) for i in range(self.cmb_ph_personel.count())}
        for isim in personel_isimleri:
            if isim not in mevcut:
                self.cmb_ph_personel.addItem(isim)
        satirlar = [k for k in kayitlar if personel_filtre == "Tüm Personeller" or (k[1] or "") == personel_filtre]
        self.tbl_personel_hareket.setRowCount(len(satirlar))
        for r, (tarih, per, m_ad, urun, tutar) in enumerate(satirlar):
            self.tbl_personel_hareket.setItem(r, 0, QTableWidgetItem(str(tarih or "-")))
            self.tbl_personel_hareket.setItem(r, 1, QTableWidgetItem(str(per or "-")))
            self.tbl_personel_hareket.setItem(r, 2, QTableWidgetItem("SATIŞ"))
            self.tbl_personel_hareket.setItem(r, 3, QTableWidgetItem(str(m_ad or "-")))
            self.tbl_personel_hareket.setItem(r, 4, QTableWidgetItem(str(urun or "-")))
            self.tbl_personel_hareket.setItem(r, 5, QTableWidgetItem(f"{float(tutar or 0):,.2f} ₺"))

    def supheli_islemler_sayfasini_ac(self):
        self.stack.setCurrentWidget(self.sayfa_supheli_islemler)
        self.supheli_islemleri_listele()

    def create_supheli_islemler_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)
        layout.addWidget(self._sayfa_basligi_widget("⚠️", "ŞÜPHELİ İŞLEMLER VE ANOMALİ ANALİZİ", "#ef4444"))

        ust_bar = QHBoxLayout()
        aciklama = QLabel("Hammadde birim fiyatı 250 ₺'den yüksek, tek satırda 35 tonu aşan veya 500.000 ₺ üzeri fahiş kayıtlar:")
        aciklama.setStyleSheet("color: #64748b; font-size: 13px;")
        ust_bar.addWidget(aciklama)
        ust_bar.addStretch()

        btn_yenile = QPushButton("🔄 Listeyi Yenile")
        btn_yenile.setFixedHeight(34)
        btn_yenile.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yenile.setStyleSheet("background: #0284c7; color: white; font-weight: bold; border-radius: 6px; padding: 0 14px;")
        btn_yenile.clicked.connect(self.supheli_islemleri_listele)
        ust_bar.addWidget(btn_yenile)

        btn_oto_onar = QPushButton("⚡ Hepsini Otomatik Onar (40 ₺ Sabitle)")
        btn_oto_onar.setFixedHeight(34)
        btn_oto_onar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_oto_onar.setStyleSheet("background: #16a34a; color: white; font-weight: bold; border-radius: 6px; padding: 0 14px;")
        btn_oto_onar.clicked.connect(self.supheli_islemleri_toplu_onar)
        ust_bar.addWidget(btn_oto_onar)
        layout.addLayout(ust_bar)

        self.tablo_supheli = QTableWidget()
        self.tablo_supheli.setColumnCount(8)
        self.tablo_supheli.setHorizontalHeaderLabels([
            "ID", "Tarih", "Müşteri", "Ürün", "Miktar (KG)", "Birim Fiyat", "Toplam Tutar", "İşlem"
        ])
        self.tablo_supheli.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tablo_supheli.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.tablo_supheli.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        self.tablo_supheli.setStyleSheet("""
            QTableWidget {
                background-color: white; border: 1px solid #e2e8f0; border-radius: 8px; gridline-color: #f1f5f9;
            }
            QHeaderView::section {
                background-color: #f8fafc; font-weight: bold; color: #334155; height: 38px; border: none;
            }
        """)
        layout.addWidget(self.tablo_supheli)
        return page

    def supheli_islemleri_listele(self):
        if not hasattr(self, "tablo_supheli"):
            return
        self.tablo_supheli.setRowCount(0)
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("""
                SELECT id, created_at, customer_name, product_name, qty, price, total_amount
                FROM sales_history
                WHERE price > 250.0 OR total_amount > 500000.0 OR qty > 5000
                ORDER BY total_amount DESC
            """)
            satirlar = c.fetchall()
            conn.close()
            self.tablo_supheli.setRowCount(len(satirlar))
            for row_idx, row in enumerate(satirlar):
                s_id, tar, cust, urun, miktar, fiyat, tutar = row
                self.tablo_supheli.setItem(row_idx, 0, QTableWidgetItem(str(s_id)))
                self.tablo_supheli.setItem(row_idx, 1, QTableWidgetItem(str(tar or "-")))
                self.tablo_supheli.setItem(row_idx, 2, QTableWidgetItem(str(cust or "MÜŞTERİSİZ")))
                self.tablo_supheli.setItem(row_idx, 3, QTableWidgetItem(str(urun or "-")))
                self.tablo_supheli.setItem(row_idx, 4, QTableWidgetItem(f"{float(miktar or 0):,.1f} KG"))
                item_fiyat = QTableWidgetItem(f"{float(fiyat or 0):,.2f} ₺")
                if float(fiyat or 0) > 250:
                    item_fiyat.setForeground(QColor("#dc2626"))
                self.tablo_supheli.setItem(row_idx, 5, item_fiyat)
                item_tutar = QTableWidgetItem(f"{float(tutar or 0):,.2f} ₺")
                item_tutar.setForeground(QColor("#b91c1c"))
                self.tablo_supheli.setItem(row_idx, 6, item_tutar)
                btn_sil = QPushButton("🗑 Sil")
                btn_sil.setFixedHeight(28)
                btn_sil.setCursor(Qt.CursorShape.PointingHandCursor)
                btn_sil.setStyleSheet("background: #fee2e2; color: #dc2626; font-weight: bold; border-radius: 4px; padding: 0 8px;")
                btn_sil.clicked.connect(lambda _, id_val=s_id: self.supheli_islemi_sil(id_val))
                self.tablo_supheli.setCellWidget(row_idx, 7, btn_sil)
        except Exception as e:
            print(f"Şüpheli işlemler sorgu hatası: {e}")

    def supheli_islemi_sil(self, satir_id):
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM sales_history WHERE id = ?", (satir_id,))
            conn.commit()
            conn.close()
            self.supheli_islemleri_listele()
        except Exception as e:
            print(f"İşlem silinemedi: {e}")

    def supheli_islemleri_toplu_onar(self):
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("""
                UPDATE sales_history
                SET price = 40.0, total_amount = qty * 40.0
                WHERE price > 250
            """)
            c.execute("DELETE FROM sales_history WHERE qty > 50000")
            conn.commit()
            conn.close()
            self.supheli_islemleri_listele()
            QMessageBox.information(self, "Başarılı", "Tüm şüpheli uçuk fiyatlar 40 ₺ hammadde rayicine çekildi ve cirolar düzeltildi!")
        except Exception as e:
            print(f"Toplu onarım hatası: {e}")

    def create_personnel_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        top_bar = QHBoxLayout()
        title = QLabel("👥 PERSONEL YÖNETİMİ & SATIŞ RAPORLARI")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1e293b;")
        top_bar.addWidget(title)
        top_bar.addStretch()

        btn_yeni_personel = QPushButton("  + YENİ PERSONEL EKLE  ")
        btn_yeni_personel.setFixedHeight(36)
        btn_yeni_personel.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yeni_personel.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: #ffffff;
                font-size: 13px;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 0 16px;
            }
            QPushButton:hover { background-color: #059669; }
        """)
        btn_yeni_personel.clicked.connect(self.popup_yeni_personel_ekle)
        top_bar.addWidget(btn_yeni_personel)
        layout.addLayout(top_bar)

        lbl_ozet = QLabel("📊 PERSONEL LİSTESİ VE PERFORMANS ÖZETİ")
        lbl_ozet.setStyleSheet("font-size: 13px; font-weight: bold; color: #475569;")
        layout.addWidget(lbl_ozet)

        self.table_personnel_summary = QTableWidget()
        self.table_personnel_summary.setColumnCount(6)
        self.table_personnel_summary.setHorizontalHeaderLabels([
            "PERSONEL ADI", "GÖREVİ", "TELEFON", "TOPLAM TONAJ (KG)", "TOPLAM CİRO (TL)", "İŞLEMLER"
        ])
        self.table_personnel_summary.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_personnel_summary.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table_personnel_summary.setColumnWidth(5, 150)
        self.table_personnel_summary.verticalHeader().setVisible(False)
        self.table_personnel_summary.setFixedHeight(220)
        self.table_personnel_summary.setStyleSheet("""
            QTableWidget { border: 1px solid #cbd5e1; background: white; font-size: 13px; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 36px; border-bottom: 2px solid #cbd5e1; }
        """)
        layout.addWidget(self.table_personnel_summary)

        lbl_detay = QLabel("📋 PERSONEL DETAYLI SATIŞ VE TAHSİLAT HAREKETLERİ")
        lbl_detay.setStyleSheet("font-size: 13px; font-weight: bold; color: #475569; margin-top: 10px;")
        layout.addWidget(lbl_detay)

        self.table_personnel_details = QTableWidget()
        self.table_personnel_details.setColumnCount(7)
        self.table_personnel_details.setHorizontalHeaderLabels([
            "TARİH", "PERSONEL", "MÜŞTERİ", "ÜRÜN", "MİKTAR (KG)", "FİYAT", "TUTAR"
        ])
        self.table_personnel_details.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_personnel_details.verticalHeader().setVisible(False)
        self.table_personnel_details.setStyleSheet("""
            QTableWidget { border: 1px solid #cbd5e1; background: white; font-size: 13px; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 36px; border-bottom: 2px solid #cbd5e1; }
        """)
        layout.addWidget(self.table_personnel_details)
        return page

    def popup_yeni_personel_ekle(self):
        dlg = PersonelEkleDialog(self)
        if dlg.exec():
            self.load_personnel_data()

    def popup_personel_duzenle(self, p_id):
        dlg = PersonelEkleDialog(self, p_id=p_id)
        if dlg.exec():
            self.load_personnel_data()

    def personel_sil_onay(self, p_id, p_name):
        reply = QMessageBox.question(
            self,
            "Personel Sil",
            f"'{p_name}' isimli personeli silmek istediğinize emin misiniz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM personnel WHERE id = ?", (p_id,))
            conn.commit()
            conn.close()
            self.load_personnel_data()
            QMessageBox.information(self, "Silindi", "Personel başarıyla silindi.")

    def load_personnel_data(self):
        if not hasattr(self, "table_personnel_summary"):
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT DISTINCT personnel_name FROM sales_history WHERE personnel_name IS NOT NULL AND personnel_name != ''")
        wp_personeller = [r[0].strip() for r in c.fetchall() if r[0] and r[0].strip()]
        for wp_p in wp_personeller:
            c.execute("INSERT OR IGNORE INTO personnel (name, full_name, role) VALUES (?, ?, 'WhatsApp Satış')", (wp_p, wp_p))
        conn.commit()

        c.execute("""
            SELECT p.id, p.name, p.role, p.phone,
                   COALESCE(SUM(s.qty), 0.0) as total_kg,
                   COALESCE(SUM(s.total_amount), 0.0) as total_tl
            FROM personnel p
            LEFT JOIN sales_history s ON s.personnel_name = p.name
            GROUP BY p.id
            ORDER BY total_tl DESC
        """)
        rows = c.fetchall()
        self.table_personnel_summary.setRowCount(len(rows))
        for r_idx, (p_id, p_name, p_role, p_phone, total_kg, total_tl) in enumerate(rows):
            self.table_personnel_summary.setRowHeight(r_idx, 38)
            it_name = QTableWidgetItem(f"👤 {p_name}")
            it_name.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.table_personnel_summary.setItem(r_idx, 0, it_name)
            self.table_personnel_summary.setItem(r_idx, 1, QTableWidgetItem(p_role or "-"))
            it_tel = QTableWidgetItem(p_phone or "-")
            it_tel.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_personnel_summary.setItem(r_idx, 2, it_tel)
            it_kg = QTableWidgetItem(f"{total_kg:,.0f} KG")
            it_kg.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_kg.setForeground(QColor("#0284c7"))
            it_kg.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_personnel_summary.setItem(r_idx, 3, it_kg)
            it_tl = QTableWidgetItem(f"{total_tl:,.2f} ₺")
            it_tl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_tl.setForeground(QColor("#16a34a"))
            it_tl.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_personnel_summary.setItem(r_idx, 4, it_tl)

            action_w = QWidget()
            action_l = QHBoxLayout(action_w)
            action_l.setContentsMargins(4, 2, 4, 2)
            action_l.setSpacing(6)
            btn_ed = QPushButton("Düzenle")
            btn_ed.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_ed.setStyleSheet("background: #0ea5e9; color: white; border-radius: 4px; font-size: 11px; font-weight: bold; padding: 3px 8px;")
            btn_ed.clicked.connect(lambda _, pid=p_id: self.popup_personel_duzenle(pid))
            action_l.addWidget(btn_ed)
            btn_del = QPushButton("Sil")
            btn_del.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_del.setStyleSheet("background: #ef4444; color: white; border-radius: 4px; font-size: 11px; font-weight: bold; padding: 3px 8px;")
            btn_del.clicked.connect(lambda _, pid=p_id, pnm=p_name: self.personel_sil_onay(pid, pnm))
            action_l.addWidget(btn_del)
            self.table_personnel_summary.setCellWidget(r_idx, 5, action_w)

        c.execute("""
            SELECT created_at, personnel_name, customer_name, product_name, qty, price, total_amount
            FROM sales_history
            ORDER BY id DESC LIMIT 100
        """)
        d_rows = c.fetchall()
        conn.close()
        self.table_personnel_details.setRowCount(len(d_rows))
        for r_idx, (dt, p_name, c_name, prod_name, q, pr, tot) in enumerate(d_rows):
            self.table_personnel_details.setRowHeight(r_idx, 36)
            self.table_personnel_details.setItem(r_idx, 0, QTableWidgetItem(str(dt)[:16] if dt else "-"))
            self.table_personnel_details.setItem(r_idx, 1, QTableWidgetItem(p_name or "Genel"))
            self.table_personnel_details.setItem(r_idx, 2, QTableWidgetItem(c_name or "-"))
            self.table_personnel_details.setItem(r_idx, 3, QTableWidgetItem(prod_name or "-"))
            it_q = QTableWidgetItem(f"{(q or 0.0):,.0f} KG")
            it_q.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_personnel_details.setItem(r_idx, 4, it_q)
            it_pr = QTableWidgetItem(f"{(pr or 0.0):,.2f} ₺")
            it_pr.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_personnel_details.setItem(r_idx, 5, it_pr)
            it_tot = QTableWidgetItem(f"{(tot or 0.0):,.2f} ₺")
            it_tot.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_tot.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_personnel_details.setItem(r_idx, 6, it_tot)

    # ================= SAYFA 3: MÜŞTERİLER (BENİMPOS ŞABLONU) =================
    def create_customers_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        top_bar = QHBoxLayout()
        title = QLabel("MÜŞTERİ LİSTESİ")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        top_bar.addWidget(title)
        top_bar.addStretch()

        self.btn_delete_page = QPushButton("  🗑 BU SAYFAYI SİL  ")
        self.btn_delete_page.setFixedHeight(36)
        self.btn_delete_page.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_delete_page.setStyleSheet("""
            QPushButton {
                background-color: #fff1f2; color: #e11d48; font-size: 12px; font-weight: bold;
                border: 1px solid #fecdd3; border-radius: 5px; padding: 0 10px;
            }
            QPushButton:hover { background-color: #ffe4e6; }
        """)
        self.btn_delete_page.clicked.connect(self.sayfadaki_musterileri_sil_onay)
        top_bar.addWidget(self.btn_delete_page)

        btn_delete_all = QPushButton("  🗑 TÜMÜNÜ SİL  ")
        btn_delete_all.setFixedHeight(36)
        btn_delete_all.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_delete_all.setStyleSheet("""
            QPushButton {
                background-color: #fee2e2; color: #dc2626; font-size: 12px; font-weight: bold;
                border: 1px solid #fca5a5; border-radius: 5px; padding: 0 10px;
            }
            QPushButton:hover { background-color: #fecaca; }
        """)
        btn_delete_all.clicked.connect(self.tum_musterileri_sil_onay)
        top_bar.addWidget(btn_delete_all)

        btn_export = QPushButton("  📥 EXCEL DIŞARI AKTAR  ")
        btn_export.setFixedHeight(36)
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.setStyleSheet("""
            QPushButton {
                background-color: #f8fafc; color: #0284c7; font-size: 12px; font-weight: bold;
                border: 1px solid #bae6fd; border-radius: 5px; padding: 0 12px;
            }
            QPushButton:hover { background-color: #e0f2fe; }
        """)
        btn_export.clicked.connect(self.excel_musteri_disari_aktar)
        top_bar.addWidget(btn_export)

        btn_import = QPushButton("  📤 EXCEL İÇERİ AKTAR  ")
        btn_import.setFixedHeight(36)
        btn_import.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_import.setStyleSheet("""
            QPushButton {
                background-color: #f8fafc; color: #059669; font-size: 12px; font-weight: bold;
                border: 1px solid #a7f3d0; border-radius: 5px; padding: 0 12px;
            }
            QPushButton:hover { background-color: #d1fae5; }
        """)
        btn_import.clicked.connect(self.excel_musteri_iceri_aktar)
        top_bar.addWidget(btn_import)

        btn_new_cust = QPushButton("  + YENİ MÜŞTERİ EKLE  ")
        btn_new_cust.setFixedHeight(40)
        btn_new_cust.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_new_cust.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: #ffffff;
                font-size: 14px;
                font-weight: bold;
                border: none;
                border-radius: 6px;
                padding: 0 18px;
            }
            QPushButton:hover {
                background-color: #059669;
            }
        """)
        btn_new_cust.clicked.connect(self.popup_yeni_musteri_ekle)
        top_bar.addWidget(btn_new_cust)

        layout.addLayout(top_bar)

        main_card = QFrame()
        main_card.setStyleSheet("background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px;")
        card_layout = QVBoxLayout(main_card)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(12)

        search_bar = QHBoxLayout()
        lbl_ara = QLabel("ARA:")
        lbl_ara.setStyleSheet("font-size: 12px; font-weight: bold; color: #475569;")
        search_bar.addWidget(lbl_ara)

        self.txt_cust_search = BuyukHarfKutusu("🔍 MÜŞTERİ ADI VEYA TELEFON...")
        self.txt_cust_search.setFixedWidth(280)
        self.txt_cust_search.setFixedHeight(34)
        self.txt_cust_search.setStyleSheet("background: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 10px; font-size: 12.5px; font-weight: bold;")
        self.txt_cust_search.textChanged.connect(lambda: self.musteri_sayfa_degistir(1))
        search_bar.addWidget(self.txt_cust_search)

        search_bar.addStretch()

        self.lbl_toplam_musteri_bilgi = QLabel("0 MÜŞTERİ")
        self.lbl_toplam_musteri_bilgi.setStyleSheet("color: #64748b; font-size: 12px; font-weight: bold;")
        search_bar.addWidget(self.lbl_toplam_musteri_bilgi)
        card_layout.addLayout(search_bar)

        self.table_customers = QTableWidget()
        self.table_customers.setColumnCount(9)
        self.table_customers.setHorizontalHeaderLabels([
            "Sıra", "Müşteri", "Alışveriş Sayısı", "Açık Hesap", "Ödeme", "Kalan Borcu", "Son Ödeme Tarihi", "Detay", "İşlem"
        ])
        self.table_customers.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_customers.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table_customers.setColumnWidth(0, 60)
        self.table_customers.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeMode.Fixed)
        self.table_customers.setColumnWidth(8, 140)
        self.table_customers.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_customers.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_customers.verticalHeader().setVisible(False)
        self.table_customers.verticalHeader().setDefaultSectionSize(44)
        self.table_customers.setStyleSheet("""
            QTableWidget { border: 1px solid #e2e8f0; font-size: 13px; background-color: #ffffff; }
            QHeaderView::section { background-color: #f8fafc; color: #334155; font-weight: bold; border: none; border-bottom: 2px solid #e2e8f0; height: 38px; padding-left: 6px; font-size: 12.5px; }
            QTableWidget::item { border-bottom: 1px solid #f1f5f9; padding: 6px; }
            QTableWidget::item:selected { background-color: #f0f9ff; color: #0369a1; }
        """)
        card_layout.addWidget(self.table_customers)

        footer_row = QHBoxLayout()
        self.lbl_musteri_alt_bilgi = QLabel("0 kayıttan 0 ile 0 arasındakiler")
        self.lbl_musteri_alt_bilgi.setStyleSheet("color: #64748b; font-size: 12px;")
        footer_row.addWidget(self.lbl_musteri_alt_bilgi)
        footer_row.addStretch()

        self.cust_pagination_box = QHBoxLayout()
        self.cust_pagination_box.setSpacing(4)
        footer_row.addLayout(self.cust_pagination_box)

        card_layout.addLayout(footer_row)
        layout.addWidget(main_card)
        return page

    def popup_yeni_musteri_ekle(self):
        dlg = MusteriEkleDialog(self)
        if dlg.exec():
            self.load_customers()
            if hasattr(self, "load_customer_details"):
                self.load_customer_details()

    def popup_musteri_duzenle(self, cust_id):
        dlg = MusteriEkleDialog(self, customer_id=cust_id)
        if dlg.exec():
            self.load_customers()

    def load_customers(self):
        """Müşterileri görseldeki 9 sütunlu BenimPOS düzeninde listeler"""
        if not hasattr(self, "table_customers"):
            return

        search = buyuk_harf(self.txt_cust_search.text().strip())
        self.table_customers.setRowCount(0)

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        count_query = "SELECT COUNT(*) FROM customers WHERE 1=1"
        params = []
        if search:
            count_query += " AND (name LIKE ? OR phone LIKE ?)"
            params.extend([f"%{search}%", f"%{search}%"])
        c.execute(count_query, params)
        toplam_kayit = c.fetchone()[0]

        toplam_sayfa = max(1, (toplam_kayit + self.sayfa_boyutu - 1) // self.sayfa_boyutu)
        if self.musteri_aktif_sayfa > toplam_sayfa:
            self.musteri_aktif_sayfa = toplam_sayfa

        offset = (self.musteri_aktif_sayfa - 1) * self.sayfa_boyutu

        query = """
            SELECT id, name, shopping_count, debt, payment, remaining_debt, last_payment_date, detail
            FROM customers WHERE 1=1
        """
        q_params = []
        if search:
            query += " AND (name LIKE ? OR phone LIKE ?)"
            q_params.extend([f"%{search}%", f"%{search}%"])
        query += " ORDER BY id ASC LIMIT ? OFFSET ?"
        q_params.extend([self.sayfa_boyutu, offset])

        c.execute(query, q_params)
        rows = c.fetchall()
        conn.close()

        self.lbl_toplam_musteri_bilgi.setText(f"{toplam_kayit} MÜŞTERİ")
        baslangic = offset + 1 if toplam_kayit else 0
        bitis = min(offset + self.sayfa_boyutu, toplam_kayit)
        if toplam_kayit > 0:
            self.lbl_musteri_alt_bilgi.setText(f"{toplam_kayit} kayıttan {baslangic} ile {bitis} arasındakiler")
        else:
            self.lbl_musteri_alt_bilgi.setText("Kayıt bulunamadı")

        self.table_customers.setRowCount(len(rows))
        for r_idx, (cid, name, shop_cnt, debt, pay, rem_debt, pay_date, detail) in enumerate(rows):
            it_seq = QTableWidgetItem(str(baslangic + r_idx))
            it_seq.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_customers.setItem(r_idx, 0, it_seq)

            ad_btn = QPushButton(name or "")
            ad_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            ad_btn.setFlat(True)
            ad_font = QFont("Segoe UI", 10, QFont.Weight.Bold)
            ad_font.setUnderline(True)
            ad_btn.setFont(ad_font)
            ad_btn.setStyleSheet("""
                QPushButton {
                    background: transparent; border: none; color: #0ea5e9;
                    text-align: left; padding: 0 4px;
                }
                QPushButton:hover { color: #0284c7; }
            """)
            ad_btn.clicked.connect(lambda _, c_id=cid: self.popup_musteri_duzenle(c_id))
            self.table_customers.setCellWidget(r_idx, 1, ad_btn)

            it_sc = QTableWidgetItem(str(shop_cnt or 0))
            it_sc.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_customers.setItem(r_idx, 2, it_sc)

            it_debt = QTableWidgetItem(f"{(debt or 0.0):,.2f}")
            it_debt.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_customers.setItem(r_idx, 3, it_debt)

            it_pay = QTableWidgetItem(f"{(pay or 0.0):,.2f}")
            it_pay.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_customers.setItem(r_idx, 4, it_pay)

            it_rem = QTableWidgetItem(f"{(rem_debt or 0.0):,.2f}")
            it_rem.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_rem.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_customers.setItem(r_idx, 5, it_rem)

            it_dt = QTableWidgetItem(pay_date or "")
            it_dt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_customers.setItem(r_idx, 6, it_dt)

            self.table_customers.setItem(r_idx, 7, QTableWidgetItem(detail or ""))

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(0, 0, 0, 0)
            action_layout.setSpacing(6)
            action_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            action_layout.addWidget(self._islem_butonu("✏", "Düzenle", "mavi", lambda _, c_id=cid: self.popup_musteri_duzenle(c_id)))
            action_layout.addWidget(self._islem_butonu("🗑", "Sil", "kirmizi", lambda _, c_id=cid, c_nm=name: self.delete_customer_confirm(c_id, c_nm)))
            self.table_customers.setCellWidget(r_idx, 8, action_widget)

        self._sayfalama_ciz(self.cust_pagination_box, self.musteri_aktif_sayfa, toplam_sayfa, self.musteri_sayfa_degistir)

    def musteri_sayfa_degistir(self, yeni_sayfa):
        self.musteri_aktif_sayfa = max(1, yeni_sayfa)
        self.load_customers()

    def delete_customer_confirm(self, cid, cname):
        reply = QMessageBox.question(
            self, "ONAY", f"'{cname}' müşterisini silmek istediğinize emin misiniz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM customers WHERE id = ?", (cid,))
            conn.commit()
            conn.close()
            self.load_customers()
            if hasattr(self, "table_customer_details"):
                self.load_customer_details()

    def tum_musterileri_sil_onay(self):
        """Tüm müşterileri veritabanından kalıcı olarak siler"""
        reply = QMessageBox.question(
            self,
            "DİKKAT - TÜM MÜŞTERİLER SİLİNECEK",
            "Sistemdeki TÜM MÜŞTERİLER kalıcı olarak silinecek!\n\nSilmek istediğinize emin misiniz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("DELETE FROM customers")
            conn.commit()
            conn.close()
            self.musteri_aktif_sayfa = 1
            self.load_customers()
            if hasattr(self, "load_customer_details"):
                self.load_customer_details()
            QMessageBox.information(self, "Bilgi", "Tüm müşteri kayıtları başarıyla silindi.")

    def sayfadaki_musterileri_sil_onay(self):
        """O an görüntülenen 10'luk aralıktaki müşterileri siler (örn: 1-10 veya 11-20)"""
        offset = (self.musteri_aktif_sayfa - 1) * self.sayfa_boyutu
        baslangic = offset + 1
        bitis = offset + self.sayfa_boyutu

        reply = QMessageBox.question(
            self,
            "SAYFAYI SİLME ONAYI",
            f"Şu anki sayfada yer alan ({baslangic} ile {bitis} arasındaki) müşteriler silinecek.\n\nEmin misiniz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        search = buyuk_harf(self.txt_cust_search.text().strip())
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        query = "SELECT id FROM customers WHERE 1=1"
        q_params = []
        if search:
            query += " AND (name LIKE ? OR phone LIKE ?)"
            q_params.extend([f"%{search}%", f"%{search}%"])
        query += " ORDER BY id ASC LIMIT ? OFFSET ?"
        q_params.extend([self.sayfa_boyutu, offset])

        c.execute(query, q_params)
        silinecek_idler = [row[0] for row in c.fetchall()]
        silinen = len(silinecek_idler)

        if silinecek_idler:
            placeholders = ",".join("?" for _ in silinecek_idler)
            c.execute(f"DELETE FROM customers WHERE id IN ({placeholders})", silinecek_idler)
            conn.commit()
        conn.close()

        self.load_customers()
        if hasattr(self, "load_customer_details"):
            self.load_customer_details()
        QMessageBox.information(self, "Bilgi", f"{silinen} müşteri bu sayfadan silindi.")


    def _musteri_sayi_oku(self, val):
        if val is None:
            return 0.0
        metin = str(val).strip().replace("₺", "").replace(" ", "")
        if not metin or metin in ("-", "None"):
            return 0.0
        if "," in metin:
            metin = metin.replace(".", "").replace(",", ".")
        try:
            return float(metin)
        except ValueError:
            return 0.0

    def _musteri_hucre(self, row, idx):
        if len(row) <= idx or row[idx] is None:
            return ""
        return str(row[idx]).strip()

    def _musteri_excel_baslik_mi(self, row):
        if not row or len(row) < 2:
            return True
        c0 = str(row[0] or "").strip()
        c1 = str(row[1] or "").strip()
        if not c1:
            return True
        if c0 == "0" and c1 == "1":
            return True
        c0u, c1u = buyuk_harf(c0), buyuk_harf(c1)
        if c0u == "SIRA" and c1u in ("MÜŞTERİ", "MUSTERI"):
            return True
        if len(row) > 2 and "ALIŞVERİŞ" in buyuk_harf(str(row[2] or "")) and c1u in ("MÜŞTERİ", "MUSTERI"):
            return True
        return False

    def _musteri_excel_satirini_ekle(self, cursor, row):
        """BenimPOS 17 sütun: 0 Sıra, 1 Müşteri, 2 Alışveriş, 3 Borç ... 16 Boylam."""
        if self._musteri_excel_baslik_mi(row):
            return False
        r = [str(hucre).strip() if hucre is not None else "" for hucre in row]
        musteri_adi = buyuk_harf(r[1] if len(r) > 1 else "")
        if not musteri_adi or musteri_adi.isdigit():
            return False

        def hucre(idx):
            return r[idx] if len(r) > idx else ""

        cursor.execute("""
            INSERT INTO customers (
                name, shopping_count, debt, payment, remaining_debt,
                last_payment_date, detail, address, phone, tax_office,
                tax_number, term_days, credit_limit, note, latitude, longitude
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            musteri_adi,
            int(self._musteri_sayi_oku(hucre(2))),
            self._musteri_sayi_oku(hucre(3)),
            self._musteri_sayi_oku(hucre(4)),
            self._musteri_sayi_oku(hucre(5)),
            hucre(6),
            hucre(7),
            hucre(8),
            hucre(9),
            buyuk_harf(hucre(10)),
            hucre(11),
            int(self._musteri_sayi_oku(hucre(12))),
            self._musteri_sayi_oku(hucre(13)),
            hucre(14),
            hucre(15),
            hucre(16),
        ))
        return True

    def _musteri_xlsx_satirlari(self, path):
        try:
            import pandas as pd
            df = pd.read_excel(path, header=None, dtype=str)
            return [list(r) for r in df.fillna("").itertuples(index=False, name=None)]
        except ImportError:
            pass
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            rows = [list(row) for row in ws.iter_rows(values_only=True)]
            wb.close()
            return rows
        except ImportError:
            raise RuntimeError(
                "xlsx dosyası için pandas veya openpyxl yüklü değil. "
                "Lütfen UTF-8 BOM'lu CSV dosyası seçin."
            )

    def excel_musteri_iceri_aktar(self):
        """BenimPOS dosyasını akıllı tarayarak müşteri adını ve tutarları doğru sütundan çeker."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Müşteri Excel/CSV Dosyasını Seç", "",
            "CSV / Excel (*.csv *.txt *.xlsx *.xls)"
        )
        if not path:
            return

        cevap = QMessageBox.question(
            self, "Temizleme Onayı",
            "Yeni aktarım yapmadan önce veritabanındaki mevcut müşteriler temizlensin mi?\n"
            "(Evet derseniz az önceki hatalı kayıtlar silinip tertemiz yüklenir)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        conn = None
        try:
            uzanti = path.lower()
            if uzanti.endswith(".xlsx") or uzanti.endswith(".xls"):
                satirlar = self._musteri_xlsx_satirlari(path)
            else:
                with open(path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
                    content = f.read(4096)
                    f.seek(0)
                    if "\t" in content:
                        delim = "\t"
                    elif ";" in content:
                        delim = ";"
                    else:
                        delim = ","
                    satirlar = list(csv.reader(f, delimiter=delim))

            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            if cevap == QMessageBox.StandardButton.Yes:
                c.execute("DELETE FROM customers")

            eklenen = 0
            for row in satirlar:
                if self._musteri_excel_satirini_ekle(c, row):
                    eklenen += 1
            conn.commit()
            conn.close()
            conn = None

            self.load_customers()
            if hasattr(self, "load_customer_details"):
                self.load_customer_details()
            QMessageBox.information(self, "Başarılı", f"{eklenen} adet müşteri adı ve bilgileriyle eksiksiz aktarıldı!")
        except Exception as e:
            if conn is not None:
                conn.rollback()
                conn.close()
            QMessageBox.critical(self, "Hata", f"Aktarım sırasında hata oluştu:\n{e}")

    def excel_musteri_disari_aktar(self):
        """Tüm müşterileri BenimPOS Excel formatında (17 sütun) dışarı aktarır."""
        path, _ = QFileDialog.getSaveFileName(
            self, "Müşteri Listesini Kaydet", "Musteriler-BenimPOS.csv",
            "Excel Uyumlu CSV (*.csv)"
        )
        if not path:
            return
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("""
                SELECT id, name, shopping_count, debt, payment, remaining_debt,
                       last_payment_date, detail, address, phone, tax_office,
                       tax_number, term_days, credit_limit, note, latitude, longitude
                FROM customers ORDER BY id ASC
            """)
            rows = c.fetchall()
            conn.close()

            with open(path, mode="w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"])
                writer.writerow([
                    "Sıra", "Müşteri", "Alışveriş Sayısı", "Borç", "Ödeme", "Kalan Borç",
                    "Son Ödeme Tarihi", "Detay", "Adres", "Telefon", "Vergi Dairesi",
                    "Vergi No", "Vade", "Limit", "Not", "Enlem", "Boylam"
                ])
                for idx, r in enumerate(rows):
                    writer.writerow([
                        idx + 1, r[1], r[2], f"{float(r[3] or 0):.2f}", f"{float(r[4] or 0):.2f}",
                        f"{float(r[5] or 0):.2f}", r[6] or "", r[7] or "", r[8] or "", r[9] or "",
                        r[10] or "", r[11] or "", r[12] or 0, f"{float(r[13] or 0):.2f}",
                        r[14] or "", r[15] or "", r[16] or ""
                    ])

            QMessageBox.information(self, "Başarılı", f"{len(rows)} müşteri Excel şablonunda dışarı aktarıldı!")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışarı aktarılırken hata oluştu:\n{e}")

    # ================= SAYFA 4: MÜŞTERİ DETAY SAYFASI =================
    def create_customer_details_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        self.detay_musteri_id = None
        self.detay_tum_satislar = []
        self.sayfa_basi_kayit = 10
        self.aktif_sayfa = 1

        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        lbl_sec = QLabel("Müşteri Seç:")
        lbl_sec.setStyleSheet("font-size: 13px; font-weight: bold; color: #1e293b;")
        top_bar.addWidget(lbl_sec)

        # Eski QComboBox yerine yazılabilir ve tahmin motorlu kutuyu koyuyoruz
        self.cmb_detay_musteriler = TahminliMusteriKutusu()
        self.cmb_detay_musteriler.setFixedHeight(36)
        self.cmb_detay_musteriler.setMinimumWidth(320)
        self.cmb_detay_musteriler.setStyleSheet("""
            QComboBox {
                border: 1.5px solid #cbd5e1;
                border-radius: 6px;
                padding: 0 10px;
                font-size: 13px;
                font-weight: 600;
                background-color: #ffffff;
                color: #0f172a;
            }
            QComboBox:focus {
                border: 2px solid #0284c7;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #cbd5e1;
                background: #ffffff;
                selection-background-color: #e0f2fe;
                selection-color: #0369a1;
                padding: 4px;
                max-height: 220px; /* Ekranı boydan boya kaplamasını engeller */
            }
        """)
        self.cmb_detay_musteriler.currentIndexChanged.connect(self.detay_musteri_degisti)
        top_bar.addWidget(self.cmb_detay_musteriler)

        btn_kart = QPushButton("📇 İletişim & Cari Bilgileri")
        btn_kart.setFixedHeight(36)
        btn_kart.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_kart.setStyleSheet("background: #f0f9ff; color: #0284c7; border: 1px solid #bae6fd; border-radius: 5px; font-weight: bold; padding: 0 12px;")
        btn_kart.clicked.connect(self.detay_musteri_karti_ac)
        top_bar.addWidget(btn_kart)

        top_bar.addStretch()

        btn_export = QPushButton("📤 Excel'e Aktar")
        btn_export.setFixedHeight(36)
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.setStyleSheet("background: #ecfdf5; color: #059669; border: 1px solid #a7f3d0; border-radius: 5px; font-weight: bold; padding: 0 14px;")
        btn_export.clicked.connect(self.detay_excel_disa_aktar)
        top_bar.addWidget(btn_export)

        btn_import = QPushButton("📥 Excel'den Yükle")
        btn_import.setFixedHeight(36)
        btn_import.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_import.setStyleSheet("background: #f8fafc; color: #475569; border: 1px solid #cbd5e1; border-radius: 5px; font-weight: bold; padding: 0 14px;")
        btn_import.clicked.connect(self.detay_excel_ice_aktar)
        top_bar.addWidget(btn_import)

        layout.addLayout(top_bar)

        h_cards = QHBoxLayout()
        h_cards.setSpacing(10)

        self.lbl_kart_tonaj = self._bilgi_karti_olustur("TOPLAM ALINAN TONAJ", "0 KG", "#0284c7")
        self.lbl_kart_ciro = self._bilgi_karti_olustur("TOPLAM SATIŞ TUTARI", "0.00 ₺", "#0f172a")
        self.lbl_kart_tahsilat = self._bilgi_karti_olustur("YAPILAN TAHSİLAT", "0.00 ₺", "#16a34a")
        self.lbl_kart_bakiye = self._bilgi_karti_olustur("KALAN NET BORÇ", "0.00 ₺", "#dc2626")

        h_cards.addWidget(self.lbl_kart_tonaj)
        h_cards.addWidget(self.lbl_kart_ciro)
        h_cards.addWidget(self.lbl_kart_tahsilat)
        h_cards.addWidget(self.lbl_kart_bakiye)
        layout.addLayout(h_cards)

        self.table_detay_satislar = QTableWidget()
        self.table_customer_details = self.table_detay_satislar
        self.table_detay_satislar.setColumnCount(8)
        self.table_detay_satislar.setHorizontalHeaderLabels([
            "TARİH", "PERSONEL", "ÜRÜN / HAMMADDE", "MİKTAR (KG)", "BİRİM FİYAT", "TOPLAM TUTAR", "TAHSİLAT", "ÖDEME ŞEKLİ"
        ])
        self.table_detay_satislar.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_detay_satislar.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_detay_satislar.verticalHeader().setVisible(False)
        self.table_detay_satislar.setStyleSheet("""
            QTableWidget { border: 1px solid #cbd5e1; background: white; font-size: 13px; }
            QHeaderView::section { background: #f8fafc; font-weight: bold; height: 38px; border-bottom: 2px solid #cbd5e1; color: #1e293b; }
        """)
        layout.addWidget(self.table_detay_satislar)

        h_page_bar = QHBoxLayout()
        self.lbl_sayfa_bilgisi = QLabel("Kayıt bulunamadı.")
        self.lbl_sayfa_bilgisi.setStyleSheet("font-size: 12px; font-weight: bold; color: #64748b;")
        h_page_bar.addWidget(self.lbl_sayfa_bilgisi)
        h_page_bar.addStretch()
        self.layout_sayfa_butonlari = QHBoxLayout()
        self.layout_sayfa_butonlari.setSpacing(4)
        h_page_bar.addLayout(self.layout_sayfa_butonlari)
        layout.addLayout(h_page_bar)

        return page

    def _bilgi_karti_olustur(self, baslik, ilk_deger, renk):
        frame = QFrame()
        frame.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 6px;")
        l = QVBoxLayout(frame)
        l.setContentsMargins(12, 10, 12, 10)
        l.setSpacing(2)
        lbl_t = QLabel(baslik)
        lbl_t.setStyleSheet("font-size: 11px; font-weight: bold; color: #64748b;")
        lbl_v = QLabel(ilk_deger)
        lbl_v.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {renk};")
        l.addWidget(lbl_t)
        l.addWidget(lbl_v)
        frame.lbl_val = lbl_v
        return frame

    def load_customer_details(self):
        """Müşteriler açılır kutusunu veritabanından doldurur"""
        if not hasattr(self, 'cmb_detay_musteriler'):
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name FROM customers ORDER BY name ASC")
        rows = c.fetchall()
        conn.close()
        self.cmb_detay_musteriler.veri_yukle(rows)
        self.detay_musteri_degisti()

    def detay_musteri_degisti(self):
        m_id = self.cmb_detay_musteriler.currentData()
        self.detay_musteri_id = m_id
        self.aktif_sayfa = 1

        if not m_id:
            self.detay_tum_satislar = []
            self.lbl_kart_tonaj.lbl_val.setText("0 KG")
            self.lbl_kart_ciro.lbl_val.setText("0.00 ₺")
            self.lbl_kart_tahsilat.lbl_val.setText("0.00 ₺")
            self.lbl_kart_bakiye.lbl_val.setText("0.00 ₺")
            self.detay_sayfayi_ciz()
            return

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""
            SELECT created_at, personnel_name, product_name, qty, price, total_amount, payment_received, payment_type
            FROM sales_history
            WHERE customer_id = ?
            ORDER BY id DESC
        """, (m_id,))
        self.detay_tum_satislar = c.fetchall()
        c.execute("SELECT remaining_debt, debt, payment FROM customers WHERE id = ?", (m_id,))
        c_row = c.fetchone()
        conn.close()

        toplam_kg = sum(float(r[3] or 0.0) for r in self.detay_tum_satislar)
        toplam_ciro = sum(float(r[5] or 0.0) for r in self.detay_tum_satislar)
        toplam_tahsilat = sum(float(r[6] or 0.0) for r in self.detay_tum_satislar)
        kalan_borc = float(c_row[0] or 0.0) if c_row else (toplam_ciro - toplam_tahsilat)

        self.lbl_kart_tonaj.lbl_val.setText(f"{toplam_kg:,.0f} KG")
        self.lbl_kart_ciro.lbl_val.setText(f"{toplam_ciro:,.2f} ₺")
        self.lbl_kart_tahsilat.lbl_val.setText(f"{toplam_tahsilat:,.2f} ₺")
        self.lbl_kart_bakiye.lbl_val.setText(f"{kalan_borc:,.2f} ₺")
        self.detay_sayfayi_ciz()

    def detay_sayfayi_ciz(self):
        toplam_kayit = len(self.detay_tum_satislar)
        while self.layout_sayfa_butonlari.count():
            item = self.layout_sayfa_butonlari.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if toplam_kayit == 0:
            self.table_detay_satislar.setRowCount(0)
            self.lbl_sayfa_bilgisi.setText("Kayıtlı satış hareketi bulunamadı.")
            return

        toplam_sayfa = (toplam_kayit + self.sayfa_basi_kayit - 1) // self.sayfa_basi_kayit
        for s in range(1, toplam_sayfa + 1):
            bas = (s - 1) * self.sayfa_basi_kayit + 1
            bit = min(s * self.sayfa_basi_kayit, toplam_kayit)
            btn = QPushButton(f"{bas}-{bit}")
            btn.setFixedHeight(28)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            if s == self.aktif_sayfa:
                btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; border-radius: 4px; padding: 0 8px;")
            else:
                btn.setStyleSheet("background-color: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px;")
            btn.clicked.connect(lambda _, sayfa_no=s: self.detay_sayfa_degistir(sayfa_no))
            self.layout_sayfa_butonlari.addWidget(btn)

        start_idx = (self.aktif_sayfa - 1) * self.sayfa_basi_kayit
        end_idx = min(start_idx + self.sayfa_basi_kayit, toplam_kayit)
        sayfa_verisi = self.detay_tum_satislar[start_idx:end_idx]
        self.lbl_sayfa_bilgisi.setText(f"Toplam {toplam_kayit} satış arasından {start_idx+1}-{end_idx} arası gösteriliyor")

        self.table_detay_satislar.setRowCount(len(sayfa_verisi))
        for r_idx, row in enumerate(sayfa_verisi):
            self.table_detay_satislar.setRowHeight(r_idx, 38)
            dt, per, urun, q, pr, tot, tah, pay_type = row

            it_dt = QTableWidgetItem(str(dt)[:16] if dt else "-")
            it_dt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_detay_satislar.setItem(r_idx, 0, it_dt)

            it_per = QTableWidgetItem(f"👤 {per or 'Genel'}")
            it_per.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            it_per.setForeground(QColor("#0369a1"))
            self.table_detay_satislar.setItem(r_idx, 1, it_per)

            it_ur = QTableWidgetItem(urun or "-")
            it_ur.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_ur.setForeground(QColor("#047857"))
            self.table_detay_satislar.setItem(r_idx, 2, it_ur)

            qf = float(q or 0)
            it_q = QTableWidgetItem(f"{int(qf)}" if qf.is_integer() else f"{qf:.1f}")
            it_q.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_q.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.table_detay_satislar.setItem(r_idx, 3, it_q)

            it_pr = QTableWidgetItem(f"{float(pr or 0):,.2f} ₺")
            it_pr.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_detay_satislar.setItem(r_idx, 4, it_pr)

            it_tot = QTableWidgetItem(f"{float(tot or 0):,.2f} ₺")
            it_tot.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            it_tot.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_detay_satislar.setItem(r_idx, 5, it_tot)

            tahf = float(tah or 0)
            it_tah = QTableWidgetItem(f"{tahf:,.2f} ₺" if tahf > 0 else "-")
            it_tah.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if tahf > 0:
                it_tah.setForeground(QColor("#16a34a"))
                it_tah.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            self.table_detay_satislar.setItem(r_idx, 6, it_tah)

            it_pay = QTableWidgetItem(str(pay_type or "AÇIK HESAP"))
            it_pay.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_detay_satislar.setItem(r_idx, 7, it_pay)

    def detay_sayfa_degistir(self, sayfa_no):
        self.aktif_sayfa = sayfa_no
        self.detay_sayfayi_ciz()

    def detay_musteri_karti_ac(self):
        if not self.detay_musteri_id:
            QMessageBox.information(self, "Bilgi", "Lütfen önce bir müşteri seçiniz.")
            return
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM customers WHERE id = ?", (self.detay_musteri_id,))
        row = c.fetchone()
        conn.close()
        if row:
            dlg = MusteriIletisimDialog(dict(row), self)
            dlg.exec()

    def detay_excel_disa_aktar(self):
        if not self.detay_musteri_id or not self.detay_tum_satislar:
            QMessageBox.warning(self, "Uyarı", "Dışa aktarılacak müşteri veya satış kaydı bulunamadı!")
            return
        m_ad = self.cmb_detay_musteriler.currentText().replace("👤 ", "").strip()
        dosya_yolu, _ = QFileDialog.getSaveFileName(
            self, "Excel Satış Listesini Kaydet", f"{m_ad}_Satis_Gecmisi.xlsx",
            "Excel Dosyası (*.xlsx);;CSV Dosyası (*.csv)"
        )
        if not dosya_yolu:
            return
        try:
            if dosya_yolu.endswith(".xlsx") and OPENPYXL_VAR:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Satış Geçmişi"
                ws.append(["Tarih", "Personel", "Ürün / Hammadde", "Miktar (KG)", "Birim Fiyat (TL)", "Toplam Tutar (TL)", "Tahsilat (TL)", "Ödeme Türü"])
                for r in self.detay_tum_satislar:
                    ws.append([
                        str(r[0]), str(r[1] or ""), str(r[2] or ""),
                        float(r[3] or 0), float(r[4] or 0), float(r[5] or 0),
                        float(r[6] or 0), str(r[7] or "")
                    ])
                wb.save(dosya_yolu)
            else:
                with open(dosya_yolu, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow(["Tarih", "Personel", "Urun", "Miktar KG", "Fiyat TL", "Tutar TL", "Tahsilat TL", "Odeme"])
                    for r in self.detay_tum_satislar:
                        writer.writerow([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]])
            QMessageBox.information(self, "Başarılı", f"'{m_ad}' satış dökümü başarıyla kaydedildi:\n{dosya_yolu}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel kaydedilirken hata oluştu:\n{str(e)}")

    def detay_excel_ice_aktar(self):
        if not self.detay_musteri_id:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce satışların ekleneceği müşteriyi seçiniz!")
            return
        dosya_yolu, _ = QFileDialog.getOpenFileName(self, "İçe Aktarılacak Excel / CSV Dosyası", "", "Veri Dosyaları (*.xlsx *.csv)")
        if not dosya_yolu:
            return
        try:
            satirlar_eklenecek = []
            m_ad = buyuk_harf(self.cmb_detay_musteriler.currentText().replace("👤", "").strip())

            def _satir_ekle(row):
                if not row or len(row) < 4 or not row[2]:
                    return
                dt = str(row[0]) if row[0] else ""
                per = str(row[1]) if row[1] else "Excel İçe Aktarım"
                urun = str(row[2])
                qty = float(row[3] or 0.0)
                pr = float(row[4] or 0.0) if len(row) > 4 and row[4] else 40.0
                tot = qty * pr
                satirlar_eklenecek.append((per, self.detay_musteri_id, m_ad, urun, qty, pr, tot, dt))

            if dosya_yolu.endswith(".xlsx") and OPENPYXL_VAR:
                wb = openpyxl.load_workbook(dosya_yolu)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    _satir_ekle(row)
            elif dosya_yolu.endswith(".csv"):
                with open(dosya_yolu, newline="", encoding="utf-8-sig") as f:
                    reader = csv.reader(f, delimiter=";")
                    next(reader, None)
                    for row in reader:
                        _satir_ekle(row)

            if satirlar_eklenecek:
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                toplam_borc = 0.0
                for s in satirlar_eklenecek:
                    c.execute("""
                        INSERT INTO sales_history (personnel_name, customer_id, customer_name, product_name, qty, price, total_amount, payment_type, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 'EXCEL AKTARIM', ?)
                    """, (s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7]))
                    toplam_borc += s[6]
                c.execute(
                    "UPDATE customers SET debt = debt + ?, remaining_debt = remaining_debt + ? WHERE id = ?",
                    (toplam_borc, toplam_borc, self.detay_musteri_id)
                )
                conn.commit()
                conn.close()
                QMessageBox.information(self, "Başarılı", f"{len(satirlar_eklenecek)} adet satış hareketi müşteriye başarıyla işlendi.")
                self.detay_musteri_degisti()
            else:
                QMessageBox.warning(self, "Uyarı", "Dosyada aktarılacak satış satırı bulunamadı.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya okunurken hata oluştu:\n{str(e)}")

    def switch_page(self, index):
        self.stack.setCurrentIndex(index)
        if index == 0:
            self.btn_sub_add.setStyleSheet("background-color: #eef6ff; color: #0088cc; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
            self.btn_sub_groups.setStyleSheet("background-color: transparent; color: #555555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
            self.load_products()
        else:
            self.btn_sub_groups.setStyleSheet("background-color: #eef6ff; color: #0088cc; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
            self.btn_sub_add.setStyleSheet("background-color: transparent; color: #555555; font-weight: bold; text-align: left; padding-left: 20px; border: none; font-size: 11.5px;")
            self.load_groups_page_list()

    # ================= SAYFA 1: ÜRÜN LİSTESİ SAYFASI =================
    def create_products_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        # Üst Başlık ve Aksiyon Butonları Alanı
        header_layout = QHBoxLayout()
        lbl_title = QLabel("ÜRÜN LİSTESİ")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1e293b;")
        header_layout.addWidget(lbl_title)
        header_layout.addStretch()
        # 1. BUTON: TÜM SATIŞLARI SIFIRLAMA BUTONU
        btn_tum_satis_sil = QPushButton("🗑 Tüm Satışları Sıfırla")
        btn_tum_satis_sil.setFixedHeight(36)
        btn_tum_satis_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_tum_satis_sil.setStyleSheet("""
            QPushButton {
                background-color: #fef2f2; color: #b91c1c; font-size: 12.5px; font-weight: bold;
                border: 1px solid #fecaca; border-radius: 6px; padding: 0 12px;
            }
            QPushButton:hover { background-color: #fee2e2; border-color: #ef4444; }
        """)
        btn_tum_satis_sil.clicked.connect(self.tum_satis_kayitlarini_temizle)
        header_layout.addWidget(btn_tum_satis_sil)
        # 2. BUTON: TÜM ÜRÜNLERİ SİLME BUTONU
        btn_tum_urun_sil = QPushButton("⚠️ Tüm Ürünleri Sil")
        btn_tum_urun_sil.setFixedHeight(36)
        btn_tum_urun_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_tum_urun_sil.setStyleSheet("""
            QPushButton {
                background-color: #fff1f2; color: #be123c; font-size: 12.5px; font-weight: bold;
                border: 1px solid #fecdd3; border-radius: 6px; padding: 0 12px;
            }
            QPushButton:hover { background-color: #ffe4e6; border-color: #f43f5e; }
        """)
        btn_tum_urun_sil.clicked.connect(self.tum_urunleri_tamamen_sil)
        header_layout.addWidget(btn_tum_urun_sil)
        # 3. BUTON: + YENİ ÜRÜN EKLE
        btn_yeni_urun = QPushButton("+ YENİ ÜRÜN EKLE")
        btn_yeni_urun.setFixedHeight(36)
        btn_yeni_urun.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_yeni_urun.setStyleSheet("""
            QPushButton {
                background-color: #10b981; color: white; font-size: 12.5px; font-weight: bold;
                border-radius: 6px; padding: 0 16px; border: none;
            }
            QPushButton:hover { background-color: #059669; }
        """)
        btn_yeni_urun.clicked.connect(self.popup_yeni_urun_ekle)
        header_layout.addWidget(btn_yeni_urun)
        layout.addLayout(header_layout)

        # 2. ANA BEYAZ KART (Arama + Tablo)
        main_card = QFrame()
        main_card.setStyleSheet("background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px;")
        card_layout = QVBoxLayout(main_card)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(12)

        # Arama ve Filtre Çubuğu
        search_bar = QHBoxLayout()
        lbl_ara = QLabel("ARA:")
        lbl_ara.setStyleSheet("font-size: 12px; font-weight: bold; color: #475569;")
        search_bar.addWidget(lbl_ara)

        self.search_input = BuyukHarfKutusu("🔍 HAMMADDE ADI İLE ARA...")
        self.search_input.setFixedWidth(280)
        self.search_input.setFixedHeight(34)
        self.search_input.setStyleSheet("background: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 10px; font-size: 12.5px; font-weight: bold;")
        self.search_input.textChanged.connect(lambda: self.urun_sayfa_degistir(1))
        search_bar.addWidget(self.search_input)

        search_bar.addStretch()

        self.lbl_toplam_urun_bilgi = QLabel("0 ÜRÜN")
        self.lbl_toplam_urun_bilgi.setStyleSheet("color: #64748b; font-size: 12px; font-weight: bold;")
        search_bar.addWidget(self.lbl_toplam_urun_bilgi)
        card_layout.addLayout(search_bar)

        # TABLO (3 İŞLEM BUTONLU)
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "GRUP", "HAMMADDE / ÇEŞİT ADI", "BİRİM", "STOK", "SATIŞ FİYATI", "ALIŞ FİYATI", "KÂR MARJI", "İŞLEMLER"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)  # ad kolonu boşluğu doldurur
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(7, 145)  # 3 ikonun rahatça sığması için 145px
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(44)  # 34x32 ikonlar sıkışmasın
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e2e8f0;
                font-size: 13px;
                background-color: #ffffff;
            }
            QHeaderView::section {
                background-color: #f8fafc;
                color: #334155;
                font-weight: bold;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                height: 38px;
                padding-left: 6px;
                font-size: 12.5px;
            }
            QTableWidget::item {
                border-bottom: 1px solid #f1f5f9;
                padding: 6px;
            }
            QTableWidget::item:selected {
                background-color: #f0f9ff;
                color: #0369a1;
            }
        """)
        card_layout.addWidget(self.table)

        # Alt Sayfalama (Pagination)
        footer_row = QHBoxLayout()
        self.lbl_urun_alt_bilgi = QLabel("0 kayıttan 0 ile 0 arasındakiler")
        self.lbl_urun_alt_bilgi.setStyleSheet("color: #64748b; font-size: 12px;")
        footer_row.addWidget(self.lbl_urun_alt_bilgi)
        footer_row.addStretch()

        self.urun_pagination_box = QHBoxLayout()
        self.urun_pagination_box.setSpacing(4)
        footer_row.addLayout(self.urun_pagination_box)

        card_layout.addLayout(footer_row)
        layout.addWidget(main_card)
        return page

    # ================= SAYFA 2: ÜRÜN GRUPLARI (BENİMPOS ŞABLONU) =================
    def create_groups_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(14)

        # 1. Üst Bar: Başlık ve "+ YENİ GRUP EKLE"
        top_bar = QHBoxLayout()
        title = QLabel("ÜRÜN GRUPLARI")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        top_bar.addWidget(title)
        top_bar.addStretch()

        self.grup_secili_idler = set()
        btn_secilenleri_sil = QPushButton("  🗑 SEÇİLENLERİ SİL  ")
        btn_secilenleri_sil.setFixedHeight(40)
        btn_secilenleri_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_secilenleri_sil.setStyleSheet("background-color:#f97316; color:white; font-weight:bold; border-radius:5px;")
        btn_secilenleri_sil.clicked.connect(self.grup_secilenleri_sil)
        top_bar.addWidget(btn_secilenleri_sil)

        btn_hepsini_sil = QPushButton("  🗑 TÜMÜNÜ SİL  ")
        btn_hepsini_sil.setFixedHeight(40)
        btn_hepsini_sil.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_hepsini_sil.setStyleSheet("background-color:#dc2626; color:white; font-weight:bold; border-radius:5px;")
        btn_hepsini_sil.clicked.connect(self.grup_hepsini_sil)
        top_bar.addWidget(btn_hepsini_sil)

        btn_new_group = QPushButton("  + YENİ GRUP EKLE  ")
        btn_new_group.setFixedHeight(40)
        btn_new_group.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_new_group.setStyleSheet("""
            QPushButton {
                background-color: #0ea5e9;
                color: #ffffff;
                font-size: 14px;
                font-weight: bold;
                border: none;
                border-radius: 6px;
                padding: 0 18px;
            }
            QPushButton:hover {
                background-color: #0284c7;
            }
        """)
        btn_new_group.clicked.connect(self.popup_yeni_grup_ekle)
        top_bar.addWidget(btn_new_group)
        layout.addLayout(top_bar)

        # 2. Beyaz Ana Kart
        main_card = QFrame()
        main_card.setStyleSheet("background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px;")
        card_layout = QVBoxLayout(main_card)
        card_layout.setContentsMargins(20, 18, 20, 18)
        card_layout.setSpacing(14)

        # Filtre Hapları
        pill_row = QHBoxLayout()
        self.btn_pill_all = QPushButton("  TÜMÜ  ")
        self.btn_pill_all.setFixedHeight(32)
        self.btn_pill_all.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold; font-size: 11.5px; border-radius: 16px; padding: 0 14px; border: none;")
        self.btn_pill_all.clicked.connect(lambda: self.filtrele_gruplar("TUMU"))
        pill_row.addWidget(self.btn_pill_all)

        self.btn_pill_empty = QPushButton("  ÜRÜNÜ OLMAYANLAR  ")
        self.btn_pill_empty.setFixedHeight(32)
        self.btn_pill_empty.setStyleSheet("background-color: #f1f5f9; color: #475569; font-weight: bold; font-size: 11.5px; border-radius: 16px; padding: 0 14px; border: 1px solid #cbd5e1;")
        self.btn_pill_empty.clicked.connect(lambda: self.filtrele_gruplar("BOS"))
        pill_row.addWidget(self.btn_pill_empty)

        pill_row.addStretch()
        self.lbl_toplam_grup_sayisi = QLabel("0 GRUP")
        self.lbl_toplam_grup_sayisi.setStyleSheet("color: #64748b; font-size: 12px; font-weight: bold;")
        pill_row.addWidget(self.lbl_toplam_grup_sayisi)
        card_layout.addLayout(pill_row)

        # Arama Satırı
        search_row = QHBoxLayout()
        lbl_show = QLabel("10 KAYIT GÖSTER")
        lbl_show.setStyleSheet("color: #64748b; font-size: 12px; font-weight: bold;")
        search_row.addWidget(lbl_show)
        search_row.addStretch()

        lbl_ara = QLabel("ARA:")
        lbl_ara.setStyleSheet("color: #475569; font-size: 12px; font-weight: bold;")
        search_row.addWidget(lbl_ara)

        self.txt_group_search = BuyukHarfKutusu("GRUP ADI ARA...")
        self.txt_group_search.setFixedWidth(200)
        self.txt_group_search.setFixedHeight(32)
        self.txt_group_search.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 4px; padding: 0 8px; font-size: 12px;")
        self.txt_group_search.textChanged.connect(lambda: self.grup_sayfa_degistir(1))
        search_row.addWidget(self.txt_group_search)
        card_layout.addLayout(search_row)

        # Tablo
        self.table_groups = QTableWidget()
        self.table_groups.setColumnCount(6)
        self.table_groups.setHorizontalHeaderLabels(["", "SIRA", "GRUP ADI", "SATIŞ SAYFASI", "GRUPTAKİ TOPLAM ÜRÜN", "İŞLEM"])
        self.table_groups.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_groups.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table_groups.setColumnWidth(0, 40)
        self.table_groups.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.table_groups.setColumnWidth(1, 70)
        self.table_groups.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table_groups.setColumnWidth(5, 220)
        self.table_groups.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_groups.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_groups.verticalHeader().setVisible(False)
        self.table_groups.verticalHeader().setDefaultSectionSize(44)
        self.table_groups.setStyleSheet("""
            QTableWidget { border: 1px solid #e2e8f0; font-size: 13px; background-color: #ffffff; }
            QHeaderView::section { background-color: #f8fafc; color: #334155; font-weight: bold; border: none; border-bottom: 2px solid #e2e8f0; height: 38px; padding-left: 8px; }
            QTableWidget::item { border-bottom: 1px solid #f1f5f9; padding: 6px; }
            QTableWidget::item:selected { background-color: #f0f9ff; color: #0369a1; }
        """)
        self.table_groups.itemChanged.connect(self._grup_checkbox_degisti)
        card_layout.addWidget(self.table_groups)

        # Alt Bilgi ve 10'ar Sayfalama
        footer_row = QHBoxLayout()
        self.lbl_grup_alt_bilgi = QLabel("0 kayıttan 0 ile 0 arasındakiler")
        self.lbl_grup_alt_bilgi.setStyleSheet("color: #64748b; font-size: 12px;")
        footer_row.addWidget(self.lbl_grup_alt_bilgi)
        footer_row.addStretch()

        self.grp_pagination_box = QHBoxLayout()
        self.grp_pagination_box.setSpacing(4)
        footer_row.addLayout(self.grp_pagination_box)

        card_layout.addLayout(footer_row)
        layout.addWidget(main_card)
        return page

    def filtrele_gruplar(self, tur):
        self.grup_aktif_sayfa = 1
        aktif = "background-color: #0284c7; color: white; font-weight: bold; font-size: 11.5px; border-radius: 16px; padding: 0 14px; border: none;"
        pasif = "background-color: #f1f5f9; color: #475569; font-weight: bold; font-size: 11.5px; border-radius: 16px; padding: 0 14px; border: 1px solid #cbd5e1;"
        if tur == "TUMU":
            self.grup_filtre_modu = "TUMU"
            self.btn_pill_all.setStyleSheet(aktif)
            self.btn_pill_empty.setStyleSheet(pasif)
        else:
            self.grup_filtre_modu = "BOS"
            self.btn_pill_empty.setStyleSheet(aktif)
            self.btn_pill_all.setStyleSheet(pasif)
        self.load_groups_page_list()

    def load_groups_page_list(self):
        """Ürün Gruplarını 10'ar 10'ar listeler ve Düzenle/Sil butonlarını tam sığdırır"""
        if not hasattr(self, "table_groups"):
            return

        search_txt = buyuk_harf(self.txt_group_search.text().strip())
        if not hasattr(self, "grup_secili_idler"):
            self.grup_secili_idler = set()
        self.table_groups.setRowCount(0)

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""
            SELECT g.id, g.name, COUNT(p.id) as urun_sayisi
            FROM product_groups g
            LEFT JOIN products p ON g.id = p.group_id
            GROUP BY g.id
            ORDER BY g.id ASC
        """)
        rows = c.fetchall()
        conn.close()

        filtrelenmis = []
        for g_id, g_name, u_count in rows:
            if search_txt and search_txt not in g_name:
                continue
            if self.grup_filtre_modu == "BOS" and u_count > 0:
                continue
            filtrelenmis.append((g_id, g_name, u_count))

        toplam_kayit = len(filtrelenmis)
        sayfa_verisi, toplam_sayfa, baslangic, bitis = self._sayfa_dilimi(
            filtrelenmis, "grup_aktif_sayfa", toplam_kayit
        )

        self.lbl_toplam_grup_sayisi.setText(f"{toplam_kayit} GRUP")
        if toplam_kayit > 0:
            self.lbl_grup_alt_bilgi.setText(f"{toplam_kayit} kayıttan {baslangic + 1} ile {bitis} arasındakiler")
        else:
            self.lbl_grup_alt_bilgi.setText("Kayıt bulunamadı")

        self.table_groups.setRowCount(len(sayfa_verisi))
        self.table_groups.blockSignals(True)
        for r_idx, (g_id, g_name, u_count) in enumerate(sayfa_verisi):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk.setCheckState(Qt.CheckState.Checked if g_id in self.grup_secili_idler else Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, g_id)
            self.table_groups.setItem(r_idx, 0, chk)

            gercek_sira = baslangic + r_idx + 1
            item_sira = QTableWidgetItem(str(gercek_sira))
            item_sira.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_groups.setItem(r_idx, 1, item_sira)

            item_name = QTableWidgetItem(g_name)
            item_name.setFont(QFont("Arial", 11, QFont.Weight.Bold))
            item_name.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            self.table_groups.setItem(r_idx, 2, item_name)

            item_sat = QTableWidgetItem("✔")
            item_sat.setForeground(QColor("#16a34a"))
            item_sat.setFont(QFont("Arial", 11, QFont.Weight.Bold))
            item_sat.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_groups.setItem(r_idx, 3, item_sat)

            item_cnt = QTableWidgetItem(str(u_count))
            item_cnt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_groups.setItem(r_idx, 4, item_cnt)

            action_widget = QWidget()
            a_layout = QHBoxLayout(action_widget)
            a_layout.setContentsMargins(0, 0, 0, 0)
            a_layout.setSpacing(6)
            a_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            btn_incele = QPushButton("🔍 İncele")
            btn_incele.setFixedHeight(28)
            btn_incele.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_incele.setStyleSheet("""
                QPushButton { background-color: #0ea5e9; color: white; border-radius: 4px; padding: 0 10px; font-weight: bold;}
                QPushButton:hover { background-color: #0284c7; }
            """)
            btn_incele.clicked.connect(lambda _, g=g_name: self.grup_icerigini_goster(g))
            a_layout.addWidget(btn_incele)
            a_layout.addWidget(self._islem_butonu(
                "✏", "Düzenle", "mavi",
                lambda _, gid=g_id, gnm=g_name: self.popup_grup_duzenle(gid, gnm)
            ))
            a_layout.addWidget(self._islem_butonu(
                "🗑", "Sil", "kirmizi",
                lambda _, gid=g_id, gnm=g_name: self.grup_sil(gid, gnm)
            ))
            self.table_groups.setCellWidget(r_idx, 5, action_widget)

        self.table_groups.blockSignals(False)
        self._sayfalama_ciz(self.grp_pagination_box, self.grup_aktif_sayfa, toplam_sayfa, self.grup_sayfa_degistir)

    def _grup_checkbox_degisti(self, item):
        if item.column() != 0:
            return
        g_id = item.data(Qt.ItemDataRole.UserRole)
        if item.checkState() == Qt.CheckState.Checked:
            self.grup_secili_idler.add(g_id)
        else:
            self.grup_secili_idler.discard(g_id)

    def grup_secilenleri_sil(self):
        if not self.grup_secili_idler:
            QMessageBox.information(self, "Bilgi", "Önce silinecek grupları işaretleyin.")
            return
        onay = QMessageBox.question(self, "Onay", f"{len(self.grup_secili_idler)} grup ve içindeki ürünler silinsin mi?")
        if onay != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        for g_id in list(self.grup_secili_idler):
            c.execute("DELETE FROM sales_history WHERE product_name IN (SELECT name FROM products WHERE group_id = ?)", (g_id,))
            c.execute("DELETE FROM products WHERE group_id = ?", (g_id,))
            c.execute("DELETE FROM product_groups WHERE id = ?", (g_id,))
        conn.commit()
        conn.close()
        self.grup_secili_idler.clear()
        self.load_groups_page_list()
        if hasattr(self, "load_product_groups"):
            self.load_product_groups()
        if hasattr(self, "load_products"):
            self.load_products()

    def grup_hepsini_sil(self):
        onay = QMessageBox.question(self, "Onay", "TÜM gruplar ve ürünler silinsin mi? Bu işlem geri alınamaz!")
        if onay != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM sales_history")
        c.execute("DELETE FROM products")
        c.execute("DELETE FROM product_groups")
        conn.commit()
        conn.close()
        self.grup_secili_idler.clear()
        self.load_groups_page_list()
        if hasattr(self, "load_product_groups"):
            self.load_product_groups()
        if hasattr(self, "load_products"):
            self.load_products()

    def grup_sayfa_degistir(self, yeni_sayfa):
        self.grup_aktif_sayfa = max(1, yeni_sayfa)
        self.load_groups_page_list()

    def _sayfa_dilimi(self, kayitlar, aktif_attr, toplam_kayit):
        toplam_sayfa = max(1, (toplam_kayit + self.sayfa_boyutu - 1) // self.sayfa_boyutu)
        aktif = getattr(self, aktif_attr)
        if aktif > toplam_sayfa:
            aktif = toplam_sayfa
            setattr(self, aktif_attr, aktif)
        baslangic = (aktif - 1) * self.sayfa_boyutu
        bitis = min(baslangic + self.sayfa_boyutu, toplam_kayit)
        return kayitlar[baslangic:bitis], toplam_sayfa, baslangic, bitis

    def _sayfalama_ciz(self, kutu, aktif_sayfa, toplam_sayfa, degistir_fn):
        while kutu.count():
            item = kutu.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if toplam_sayfa <= 1:
            return

        grup_araligi = 20
        mevcut_grup = (aktif_sayfa - 1) // grup_araligi
        baslangic_sayfa = mevcut_grup * grup_araligi + 1
        bitis_sayfa = min(baslangic_sayfa + grup_araligi - 1, toplam_sayfa)
        nav_stil = "border: 1px solid #cbd5e1; background: white; color: #475569; font-size: 11px; font-weight: bold; padding: 0 8px; border-radius: 4px;"

        if baslangic_sayfa > 1:
            btn_first = QPushButton("« İlk")
            btn_first.setFixedHeight(30)
            btn_first.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_first.setStyleSheet(nav_stil)
            btn_first.clicked.connect(lambda: degistir_fn(1))
            kutu.addWidget(btn_first)

        btn_prev = QPushButton("ÖNCEKİ")
        btn_prev.setFixedHeight(30)
        btn_prev.setEnabled(aktif_sayfa > 1)
        btn_prev.setStyleSheet("border: 1px solid #cbd5e1; background: white; color: #475569; font-size: 11px; font-weight: bold; padding: 0 10px; border-radius: 4px;")
        btn_prev.clicked.connect(lambda: degistir_fn(aktif_sayfa - 1))
        kutu.addWidget(btn_prev)

        for s in range(baslangic_sayfa, bitis_sayfa + 1):
            btn_page = QPushButton(str(s))
            btn_page.setFixedSize(32, 30)
            btn_page.setCursor(Qt.CursorShape.PointingHandCursor)
            if s == aktif_sayfa:
                btn_page.setStyleSheet("border: none; background: #0284c7; color: white; font-size: 12px; font-weight: bold; border-radius: 4px;")
            else:
                btn_page.setStyleSheet("border: 1px solid #cbd5e1; background: white; color: #475569; font-size: 12px; font-weight: bold; border-radius: 4px;")
            btn_page.clicked.connect(lambda _, p=s: degistir_fn(p))
            kutu.addWidget(btn_page)

        btn_next = QPushButton("SONRAKİ")
        btn_next.setFixedHeight(30)
        btn_next.setEnabled(aktif_sayfa < toplam_sayfa)
        btn_next.setStyleSheet("border: 1px solid #cbd5e1; background: white; color: #475569; font-size: 11px; font-weight: bold; padding: 0 10px; border-radius: 4px;")
        btn_next.clicked.connect(lambda: degistir_fn(aktif_sayfa + 1))
        kutu.addWidget(btn_next)

        if bitis_sayfa < toplam_sayfa:
            btn_last = QPushButton("Son »")
            btn_last.setFixedHeight(30)
            btn_last.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_last.setStyleSheet(nav_stil)
            btn_last.clicked.connect(lambda: degistir_fn(toplam_sayfa))
            kutu.addWidget(btn_last)

    def _islem_butonu(self, simge, tooltip, renk, tiklama):
        stiller = {
            "mavi": (
                "background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 4px;"
                " color: #0284c7; font-size: 15px; font-weight: bold; padding: 0;",
                "background: #e0f2fe;",
            ),
            "kirmizi": (
                "background: #fef2f2; border: 1px solid #fecaca; border-radius: 4px;"
                " color: #dc2626; font-size: 15px; font-weight: bold; padding: 0;",
                "background: #fee2e2;",
            ),
            "yesil": (
                "background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 4px;"
                " font-size: 15px; padding: 0;",
                "background: #dcfce7;",
            ),
        }
        normal, hover = stiller[renk]
        btn = QPushButton(simge)
        btn.setFixedSize(34, 32)
        btn.setToolTip(tooltip)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(f"QPushButton {{ {normal} }} QPushButton:hover {{ {hover} }}")
        btn.clicked.connect(tiklama)
        return btn

    def popup_grup_duzenle(self, g_id, g_name):
        text, ok = QInputDialog.getText(
            self, "GRUP DÜZENLE", "GRUP ADINI GÜNCELLE:",
            QLineEdit.EchoMode.Normal, g_name
        )
        if not ok or not text.strip():
            return
        yeni = buyuk_harf(text.strip())
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        try:
            c.execute("UPDATE product_groups SET name = ? WHERE id = ?", (yeni, g_id))
            conn.commit()
            self.load_groups_page_list()
            self.load_products()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Hata", "BU GRUP ADI ZATEN MEVCUT!")
        finally:
            conn.close()

    def popup_yeni_grup_ekle(self):
        """+ YENİ GRUP EKLE butonuna basınca açılan şık giriş kutusu"""
        text, ok = QInputDialog.getText(self, "YENİ GRUP EKLE", "YENİ HAMMADDE GRUP ADI GİRİN:\n(Örn: POM, ABS, PP MOBLEN)")
        if ok and text.strip():
            grup_adi = buyuk_harf(text.strip())
            if grup_adi.strip().upper() in ["WHATSAPP", "GENEL POLİMER", "TÜMÜ"]:
                return
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("INSERT INTO product_groups (name) VALUES (?)", (grup_adi,))
                conn.commit()
                self.load_groups_page_list()
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "Hata", "BU GRUP ADI ZATEN MEVCUT!")
            finally:
                conn.close()

    def grup_icerigini_goster(self, grup_adi):
        """Tıklanan ürün grubunun altındaki tüm ürünleri ve fiyatlarını liste halinde ekrana basar."""
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            if grup_adi in ["GRUPSUZLAR", "DİĞER"] or not grup_adi:
                c.execute("""
                    SELECT p.name, p.sell_price
                    FROM products p
                    LEFT JOIN product_groups g ON p.group_id = g.id
                    WHERE g.name IN ('GRUPSUZLAR', 'DİĞER') OR g.name IS NULL OR g.name = ''
                    ORDER BY p.name
                """)
            else:
                c.execute("""
                    SELECT p.name, p.sell_price
                    FROM products p
                    JOIN product_groups g ON p.group_id = g.id
                    WHERE g.name = ? COLLATE NOCASE
                    ORDER BY p.name
                """, (grup_adi,))
            urunler = c.fetchall()
            conn.close()
            if not urunler:
                QMessageBox.information(self, "Bilgi", f"'{grup_adi}' grubunda kayıtlı ürün bulunmuyor.")
                return
            dialog = QDialog(self)
            dialog.setWindowTitle(f"📁 '{grup_adi}' İçindeki Ürünler ({len(urunler)} Adet)")
            dialog.setFixedSize(450, 550)
            dialog.setStyleSheet("background-color: #f8fafc;")
            layout = QVBoxLayout(dialog)
            baslik = QLabel(f"{str(grup_adi).upper()} GRUBU ÜRÜNLERİ")
            baslik.setStyleSheet("font-weight: bold; color: #334155; font-size: 14px;")
            layout.addWidget(baslik)
            liste = QListWidget()
            liste.setStyleSheet("""
                QListWidget {
                    background: white; border: 1px solid #cbd5e1; border-radius: 6px; padding: 5px; font-size: 13px;
                }
                QListWidget::item { border-bottom: 1px solid #f1f5f9; padding: 8px; }
            """)
            for ad, fiyat in urunler:
                liste.addItem(f"📌 {ad}   —   {float(fiyat or 0):,.2f} ₺")
            layout.addWidget(liste)
            btn_kapat = QPushButton("Kapat")
            btn_kapat.setFixedHeight(35)
            btn_kapat.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_kapat.setStyleSheet("background-color: #64748b; color: white; font-weight: bold; border-radius: 6px;")
            btn_kapat.clicked.connect(dialog.close)
            layout.addWidget(btn_kapat)
            dialog.exec()
        except Exception as e:
            print(f"Grup içeriği gösterme hatası: {e}")

    def grup_sil(self, grup_id, grup_adi):
        """Grubu ve gruba bağlı ürünleri veritabanından kilitlemeden siler"""
        reply = QMessageBox.question(
            self,
            "GRUP SİLME ONAYI",
            f"'{grup_adi}' grubunu silmek istediğinize emin misiniz?\n\n"
            f"DİKKAT: Bu gruba bağlı tüm ürünler de silinecektir!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            # 1. Önce bu gruba bağlı ürünlerin satış hareketlerini temizle
            c.execute("""
                DELETE FROM sales_history
                WHERE product_name IN (SELECT name FROM products WHERE group_id = ?)
            """, (grup_id,))
            # 2. Gruba bağlı ürünleri sil (Foreign key engelini kaldırır)
            c.execute("DELETE FROM products WHERE group_id = ?", (grup_id,))
            # 3. Grubun kendisini sil
            c.execute("DELETE FROM product_groups WHERE id = ?", (grup_id,))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Başarılı", f"'{grup_adi}' grubu ve bağlı ürünleri silindi.")
            if hasattr(self, 'load_groups_page_list'): self.load_groups_page_list()
            if hasattr(self, 'load_product_groups_table'): self.load_product_groups_table()
            if hasattr(self, 'load_product_groups'): self.load_product_groups()
            if hasattr(self, 'load_products'): self.load_products()
        except Exception as e:
            QMessageBox.critical(self, "Silme Hatası", f"Grup silinirken hata oluştu:\n{str(e)}")

    def tum_satis_kayitlarini_temizle(self):
        """Tüm satış hareketlerini, sahte test kayıtlarını ve cari borçları sıfırlar"""
        reply = QMessageBox.warning(
            self,
            "DİKKAT: TÜM SATIŞLARI SİL",
            "Sistemde kayıtlı olan TÜM satış geçmişi silinecek,\n"
            "Müşterilerin borç ve tahsilat bakiyeleri 0.00 ₺ olarak sıfırlanacaktır.\n\n"
            "Bu işlem geri alınamaz! Devam etmek istiyor musunuz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            conn = sqlite3.connect(DB_NAME, timeout=20)
            c = conn.cursor()
            # 1. Satış geçmişi tablosunu tamamen boşalt
            c.execute("DELETE FROM sales_history")
            # 2. Müşteri carilerini ve borçlarını sıfırla
            c.execute("""
                UPDATE customers
                SET debt = 0.0, payment = 0.0, remaining_debt = 0.0, shopping_count = 0
            """)
            # 3. Ürün stoklarını isteğe bağlı başlangıç haline (0 KG) çek
            c.execute("UPDATE products SET stock_kg = 0.0")
            conn.commit()
            conn.close()
            QMessageBox.information(
                self,
                "Başarılı",
                "Tüm satış kayıtları başarıyla silindi ve cari hesaplar sıfırlandı."
            )
            # Ekranları ve tabloları anında tazele
            if hasattr(self, 'load_products'): self.load_products()
            if hasattr(self, 'load_customers'): self.load_customers()
            if hasattr(self, 'load_customer_details'): self.load_customer_details()
            if hasattr(self, 'table_wp_onizleme'): self.table_wp_onizleme.setRowCount(0)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kayıtlar silinirken bir hata oluştu:\n{str(e)}")

    def tum_urunleri_tamamen_sil(self):
        """Tüm ürün kartlarını ve bağlı verileri tek seferde sıfırlar"""
        reply = QMessageBox.warning(
            self,
            "TÜM ÜRÜNLERİ SİL",
            "Sistemdeki TÜM ürün kayıtları tamamen silinecektir.\n\n"
            "Ürün listesini sıfırdan temiz yüklemek istiyorsanız onaylayın.\n"
            "Bu işlem geri alınamaz!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            c.execute("DELETE FROM sales_history")
            c.execute("DELETE FROM products")
            c.execute("UPDATE customers SET debt = 0.0, payment = 0.0, remaining_debt = 0.0, shopping_count = 0")
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Başarılı", "Tüm ürünler başarıyla silindi. Liste tertemiz oldu.")
            if hasattr(self, 'load_products'): self.load_products()
            if hasattr(self, 'load_groups_page_list'): self.load_groups_page_list()
            if hasattr(self, 'load_product_groups_table'): self.load_product_groups_table()
            if hasattr(self, 'load_product_groups'): self.load_product_groups()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Ürünler silinirken sorun çıktı:\n{str(e)}")

    def popup_yeni_urun_ekle(self):
        """+ YENİ ÜRÜN EKLE tıklandığında açılan pencere"""
        dlg = UrunDuzenleDialog(self)
        if dlg.exec():
            self.load_products()
            self.load_groups_page_list()

    def popup_urun_duzenle(self, product_id):
        """Düzenle (Kalem) butonuna tıklandığında açılan pencere"""
        dlg = UrunDuzenleDialog(self, product_id=product_id)
        if dlg.exec():
            self.load_products()
            self.load_groups_page_list()

    def load_products(self):
        """Ürünleri listeler ve her satıra 3 adet şık işlem butonu ekler"""
        if not hasattr(self, "table"):
            return

        search = buyuk_harf(self.search_input.text().strip())
        self.table.setRowCount(0)

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        query = """
            SELECT p.id, g.name, p.name, p.unit, p.stock_kg,
                   p.buy_price, p.buy_currency, p.sell_price, p.sell_currency
            FROM products p
            JOIN product_groups g ON p.group_id = g.id
        """
        params = []
        if search:
            query += " WHERE p.name LIKE ? OR g.name LIKE ?"
            params.extend([f"%{search}%", f"%{search}%"])

        query += " ORDER BY p.id DESC"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        toplam_kayit = len(rows)
        sayfa_verisi, toplam_sayfa, baslangic, bitis = self._sayfa_dilimi(
            rows, "urun_aktif_sayfa", toplam_kayit
        )

        self.lbl_toplam_urun_bilgi.setText(f"{toplam_kayit} ÜRÜN")
        if toplam_kayit > 0:
            self.lbl_urun_alt_bilgi.setText(f"{toplam_kayit} kayıttan {baslangic + 1} ile {bitis} arasındakiler")
        else:
            self.lbl_urun_alt_bilgi.setText("Kayıt bulunamadı")

        self.table.setRowCount(len(sayfa_verisi))

        for r_idx, (p_id, g_name, p_name, unit, stock, buy, buy_kur, sell, sell_kur) in enumerate(sayfa_verisi):
            al_sim = CURRENCY_SYMBOLS.get(buy_kur, buy_kur or "")
            sat_sim = CURRENCY_SYMBOLS.get(sell_kur, sell_kur or "")

            # Grup (Büyük & Kalın)
            it_grp = QTableWidgetItem(g_name)
            it_grp.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            self.table.setItem(r_idx, 0, it_grp)

            # Hammadde Adı
            it_name = QTableWidgetItem(p_name)
            it_name.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            self.table.setItem(r_idx, 1, it_name)

            # Birim
            it_unit = QTableWidgetItem(unit if unit else DEFAULT_UNIT)
            it_unit.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r_idx, 2, it_unit)

            # Stok
            it_stk = QTableWidgetItem(f"{stock:,.2f}")
            it_stk.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(r_idx, 3, it_stk)

            # Satış Fiyatı (punto ve kalınlık stok/alış ile aynı)
            it_sell = QTableWidgetItem(f"{sell:,.2f} {sat_sim}")
            it_sell.setForeground(QColor("#0284c7"))
            it_sell.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(r_idx, 4, it_sell)

            # Alış Fiyatı
            it_buy = QTableWidgetItem(f"{buy:,.2f} {al_sim}")
            it_buy.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(r_idx, 5, it_buy)

            # Kâr Marjı (iki fiyat farklı para birimindeyse oran anlamsız olur)
            if buy_kur != sell_kur:
                it_marj = QTableWidgetItem("—")
                it_marj.setToolTip("Alış ve satış farklı para biriminde; kâr marjı hesaplanamaz.")
            elif buy > 0:
                marj = ((sell - buy) / buy) * 100
                it_marj = QTableWidgetItem(f"% {marj:.1f}")
                if marj > 0:
                    it_marj.setForeground(QColor("#16a34a"))
                else:
                    it_marj.setForeground(QColor("#dc2626"))
            else:
                it_marj = QTableWidgetItem("% 0.0")
            it_marj.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r_idx, 6, it_marj)

            # --- İŞLEMLER BUTONLARI (3 ADET, 34x32) ---
            action_box = QWidget()
            a_layout = QHBoxLayout(action_box)
            a_layout.setContentsMargins(0, 0, 0, 0)
            a_layout.setSpacing(6)
            a_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            a_layout.addWidget(self._islem_butonu(
                "🏷", "Barkod Yazdır", "mavi",
                lambda pid=p_id: QMessageBox.information(self, "Barkod", f"Ürün ID {pid} için barkod yazdırma hazırlanıyor...")
            ))
            a_layout.addWidget(self._islem_butonu(
                "✏", "Düzenle", "mavi",
                lambda pid=p_id: self.popup_urun_duzenle(pid)
            ))
            a_layout.addWidget(self._islem_butonu(
                "🗑", "Sil", "kirmizi",
                lambda _, pid=p_id, nm=p_name: self.urun_sil(pid, nm)
            ))
            self.table.setCellWidget(r_idx, 7, action_box)

        self._sayfalama_ciz(self.urun_pagination_box, self.urun_aktif_sayfa, toplam_sayfa, self.urun_sayfa_degistir)

    def urun_sayfa_degistir(self, yeni_sayfa):
        self.urun_aktif_sayfa = max(1, yeni_sayfa)
        self.load_products()

    def urun_sil(self, urun_id, urun_adi):
        reply = QMessageBox.question(
            self,
            "Ürün Sil",
            f"'{urun_adi}' ürününü silmek istediğinize emin misiniz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            conn = sqlite3.connect(DB_NAME, timeout=15)
            c = conn.cursor()
            c.execute("DELETE FROM sales_history WHERE product_name = ?", (urun_adi,))
            c.execute("DELETE FROM products WHERE id = ?", (urun_id,))
            conn.commit()
            conn.close()
            if hasattr(self, 'load_products'): self.load_products()
            if hasattr(self, 'load_groups_page_list'): self.load_groups_page_list()
            if hasattr(self, 'load_product_groups_table'): self.load_product_groups_table()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Ürün silinemedi: {str(e)}")

def bozuk_musteri_isimlerini_onar():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM customers")
    musteriler = c.fetchall()

    cop_desenler = [
        r'\bPM\b', r'\bAM\b', r'\bBRO AHMET\b', r'\bBABA\b', r'\bBABA1\b',
        r'\bSİZ\b', r'\bSIZ\b', r'\bKIRMIZI\b', r'\bKRIMIZI\b', r'\bBEYAZ\b',
        r'\bSİYAH\b', r'\bMAVİ\b', r'\bYEŞİL\b', r'\bGRİ\b', r'\bTURUNCU\b',
        r'\bNTR\b', r'\bNAT\b'
    ]

    for mid, name in musteriler:
        yeni_ad = name
        for pat in cop_desenler:
            yeni_ad = re.sub(pat, '', yeni_ad, flags=re.IGNORECASE)
        yeni_ad = re.sub(r'\s+', ' ', yeni_ad).strip()

        if yeni_ad and yeni_ad != name:
            c.execute("UPDATE customers SET name = ? WHERE id = ?", (yeni_ad, mid))
            c.execute("UPDATE sales_history SET customer_name = ? WHERE customer_id = ?", (yeni_ad, mid))

    conn.commit()
    conn.close()


def musteri_isimlerini_kapsamli_temizle_ve_birlestir():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
    musteriler = c.fetchall()

    yeni_isim_haritasi = {}
    for mid, ad, debt, payment, remaining, sayim in musteriler:
        _, temiz_ad = WhatsAppSatisAyristirici.musteri_adini_ayikla(ad or "", [])
        if not temiz_ad or temiz_ad == "MÜŞTERİSİZ SATIŞ":
            temiz_ad = (ad or "").strip().upper() or "MÜŞTERİSİZ SATIŞ"
        yeni_isim_haritasi.setdefault(temiz_ad, []).append(
            (mid, debt or 0.0, payment or 0.0, remaining or 0.0, sayim or 0)
        )

    for temiz_ad, kayitlar in yeni_isim_haritasi.items():
        if len(kayitlar) == 1:
            mid = kayitlar[0][0]
            c.execute("UPDATE customers SET name = ? WHERE id = ?", (temiz_ad, mid))
            c.execute("UPDATE sales_history SET customer_name = ? WHERE customer_id = ?", (temiz_ad, mid))
            continue

        ana_id = min(k[0] for k in kayitlar)
        toplam_debt = sum(k[1] for k in kayitlar)
        toplam_payment = sum(k[2] for k in kayitlar)
        toplam_remaining = sum(k[3] for k in kayitlar)
        toplam_sayim = sum(k[4] for k in kayitlar)

        c.execute("""UPDATE customers
                     SET name = ?, debt = ?, payment = ?, remaining_debt = ?, shopping_count = ?
                     WHERE id = ?""",
                  (temiz_ad, toplam_debt, toplam_payment, toplam_remaining, toplam_sayim, ana_id))

        for mid, *_ in kayitlar:
            if mid == ana_id:
                continue
            c.execute("UPDATE sales_history SET customer_id = ?, customer_name = ? WHERE customer_id = ?",
                      (ana_id, temiz_ad, mid))
            c.execute("DELETE FROM customers WHERE id = ?", (mid,))

    conn.commit()
    conn.close()


def standart_urun_ve_grup_belirle(ham_metin, grup_ipucu=""):
    if not ham_metin:
        return "GRUPSUZLAR", "DİĞER"

    s = str(ham_metin).upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
    kelimeler = s.split()

    tespit_edilen_grup = None
    for ana_grup, varyasyonlar in BILINEN_HAMMADDELER.items():
        for v in sorted(varyasyonlar, key=len, reverse=True):
            v_up = str(v).upper().replace("İ", "I")
            if len(v_up) <= 2:
                if v_up in kelimeler:
                    tespit_edilen_grup = ana_grup
                    break
            elif v_up in kelimeler or v_up in s:
                tespit_edilen_grup = ana_grup
                break
        if tespit_edilen_grup:
            break

    if not tespit_edilen_grup:
        if grup_ipucu and grup_ipucu.upper() not in ["GRUPSUZLAR", "TUMU", "DIGER", ""]:
            tespit_edilen_grup = grup_ipucu.upper()
        else:
            tespit_edilen_grup = "GRUPSUZLAR"

    nitelik = "NATUREL"
    if any(k in s for k in ["ELYAFLI", "ELYAF", "CE"]):
        nitelik = "ELYAFLI"
    elif any(k in s for k in ["KALSITLI", "KALSIT"]):
        nitelik = "KALSİTLİ"
    elif any(k in s for k in ["ORJINAL", "ORJ", "ORIJINAL"]):
        nitelik = "ORJİNAL"
    elif any(k in s for k in ["BEYAZ", "BEYZ"]):
        nitelik = "BEYAZ"
    elif any(k in s for k in ["SIYAH", "SYH"]):
        nitelik = "SİYAH"
    elif any(k in s for k in ["GRI", "FUME"]):
        nitelik = "GRİ"
    elif any(k in s for k in ["RENKLI", "MAVI", "KIRMIZI", "YESIL", "SARI"]):
        nitelik = "RENKLİ"
    elif any(k in s for k in ["CAPAK", "KIRIM"]):
        nitelik = "ÇAPAK"
    elif any(k in s for k in ["SEFFAF", "OPAK"]):
        nitelik = "ŞEFFAF"

    standart_ad = f"{nitelik} {tespit_edilen_grup}" if tespit_edilen_grup != "GRUPSUZLAR" else ham_metin.strip().upper()
    return tespit_edilen_grup, standart_ad


def urunleri_temizle_ve_birlestir_maks10():
    """Veritabanındaki 38 POM, 31 Antişok gibi şişmiş tüm ürünleri temizler ve birleştirir"""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        SELECT p.id, p.name, COALESCE(g.name, ''), p.sell_price
        FROM products p
        LEFT JOIN product_groups g ON p.group_id = g.id
    """)
    mevcut_urunler = c.fetchall()

    kanonik_urunler = {}
    for uid, ad, grup, fiyat in mevcut_urunler:
        temiz_grup, temiz_ad = standart_urun_ve_grup_belirle(ad or "", grup or "")
        anahtar = (temiz_grup, temiz_ad)
        kanonik_urunler.setdefault(anahtar, []).append((uid, fiyat or 0.0, ad))

    for (temiz_grup, temiz_ad), kayitlar in kanonik_urunler.items():
        ana_id = kayitlar[0][0]
        ana_fiyat = max(k[1] for k in kayitlar)
        if temiz_grup == "PP":
            grup_adi = "PP MOBLEN"
        elif temiz_grup in ["DİĞER", ""]:
            grup_adi = "GRUPSUZLAR"
        else:
            grup_adi = temiz_grup
        c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (grup_adi,))
        g_row = c.fetchone()
        if g_row:
            gid = g_row[0]
        else:
            c.execute("INSERT INTO product_groups (name) VALUES (?)", (grup_adi,))
            gid = c.lastrowid
        c.execute("UPDATE products SET name = ?, group_id = ?, sell_price = ? WHERE id = ?",
                  (temiz_ad, gid, ana_fiyat, ana_id))
        for k_id, _, eski_ad in kayitlar[1:]:
            try:
                c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (temiz_ad, eski_ad))
            except Exception:
                pass
            c.execute("DELETE FROM products WHERE id = ?", (k_id,))

    conn.commit()
    conn.close()


def veritabanindaki_yanlis_gruplari_duzelt():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("DELETE FROM products WHERE name LIKE '%GENEL POLİMER%'")
    # 2. Ürünleri adındaki polimere göre doğru gruplarına dağıt
    c.execute("SELECT id, name FROM products")
    urunler = c.fetchall()
    for pid, name in urunler:
        n_up = name.upper()
        if "POM" in n_up: g_name = "POM"
        elif "ABS" in n_up: g_name = "ABS"
        elif any(k in n_up for k in ["HDPE", "ELTEKS", "I20"]): g_name = "HDPE"
        elif any(k in n_up for k in ["PP", "MOBLEN"]): g_name = "PP MOBLEN"
        elif "PA66" in n_up: g_name = "PA66"
        elif "PA6" in n_up: g_name = "PA6"
        elif "PVC" in n_up: g_name = "PVC"
        else: g_name = "GRUPSUZLAR"
        c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (g_name,))
        g_row = c.fetchone()
        if g_row:
            gid = g_row[0]
            c.execute("UPDATE products SET group_id = ? WHERE id = ?", (gid, pid))
    conn.commit()
    conn.close()


def veritabanindaki_bozuk_musterileri_birlestir_ve_temizle():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name, debt, payment, remaining_debt FROM customers")
    musteriler = c.fetchall()
    kume_haritasi = defaultdict(list)
    for m in musteriler:
        m_id, m_name, debt, payment, rem_debt = m

        kelimeler = [w for w in m_name.split() if MusteriKumeleyici.tr_temizle_ve_koke_in(w) not in MusteriKumeleyici.YASAKLI_KELIMELER]
        temiz_ad = " ".join(kelimeler).strip() if kelimeler else m_name

        kok = MusteriKumeleyici.musteri_kok_olustur(temiz_ad)
        if not kok or len(temiz_ad) < 3:
            kok = "MUSTERISIZ"
        kume_haritasi[kok].append((m_id, temiz_ad, debt or 0.0, payment or 0.0, rem_debt or 0.0))
    for kok, m_listesi in kume_haritasi.items():
        if kok == "MUSTERISIZ":
            continue
        lider = max(m_listesi, key=lambda x: (len(x[1]), x[0]))
        lider_id, lider_ad = lider[0], lider[1]
        toplam_debt = sum(x[2] for x in m_listesi)
        toplam_payment = sum(x[3] for x in m_listesi)
        toplam_rem = sum(x[4] for x in m_listesi)
        c.execute("""
            UPDATE customers
            SET name = ?, debt = ?, payment = ?, remaining_debt = ?
            WHERE id = ?
        """, (lider_ad, toplam_debt, toplam_payment, toplam_rem, lider_id))
        for diger in m_listesi:
            diger_id = diger[0]
            if diger_id != lider_id:
                c.execute("UPDATE sales_history SET customer_id = ?, customer_name = ? WHERE customer_id = ?", (lider_id, lider_ad, diger_id))
                c.execute("DELETE FROM customers WHERE id = ?", (diger_id,))
    conn.commit()
    conn.close()


def hatali_malzeme_carilerini_temizle():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id FROM customers WHERE name = 'MÜŞTERİSİZ SATIŞ' LIMIT 1")
    row = c.fetchone()
    if row:
        m_bos_id = row[0]
    else:
        c.execute("INSERT INTO customers (name, debt, remaining_debt) VALUES ('MÜŞTERİSİZ SATIŞ', 0.0, 0.0)")
        m_bos_id = c.lastrowid
    hatali_isimler = [
        "RENGE GİDER ANT", "RENGE GİDER", "ANT", "SİYAH", "KALSİT",
        "NUMUNE", "KIRIM", "POM ÇAPAK", "BEYAZ POM"
    ]
    for h_ad in hatali_isimler:
        c.execute("SELECT id, debt, payment, remaining_debt FROM customers WHERE name = ? COLLATE NOCASE", (h_ad,))
        rows = c.fetchall()
        for rid, d, p, rem in rows:
            if rid == m_bos_id:
                continue
            c.execute("UPDATE sales_history SET customer_id = ?, customer_name = 'MÜŞTERİSİZ SATIŞ' WHERE customer_id = ?", (m_bos_id, rid))
            c.execute("DELETE FROM customers WHERE id = ?", (rid,))
    conn.commit()
    conn.close()


def musteri_isimlerinde_konumlari_basa_al():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM customers")
    musteriler = c.fetchall()
    for mid, name in musteriler:
        yeni_ad = MusteriKumeleyici.konumu_basa_al_ve_formatla(name)
        if yeni_ad != name:
            c.execute("UPDATE customers SET name = ? WHERE id = ?", (yeni_ad, mid))
            c.execute("UPDATE sales_history SET customer_name = ? WHERE customer_id = ?", (yeni_ad, mid))
    conn.commit()
    conn.close()


def veritabanindaki_pom_harici_granulleri_temizle():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM products WHERE name NOT LIKE '%POM%' AND (name LIKE '%GRANÜL%' OR name LIKE '%GRANUR%')")
    hatali_urunler = c.fetchall()
    for pid, name in hatali_urunler:
        yeni_ad = re.sub(r'\b(GRANÜL|GRANUR|GRANUL|GRANÜR)\b', '', name, flags=re.IGNORECASE)
        yeni_ad = re.sub(r'\s+', ' ', yeni_ad).strip()
        u_up = yeni_ad.upper()
        if "ABS" in u_up: hedef_grup = "ABS"
        elif "PA66" in u_up: hedef_grup = "PA66"
        elif "PA6" in u_up: hedef_grup = "PA6"
        elif any(k in u_up for k in ["HDPE", "ELTEKS", "I20"]): hedef_grup = "HDPE"
        elif any(k in u_up for k in ["PP", "MOBLEN"]): hedef_grup = "PP MOBLEN"
        elif "ANTİŞOK" in u_up or "HIPS" in u_up: hedef_grup = "ANTİŞOK"
        else: hedef_grup = "GRUPSUZLAR"
        c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (hedef_grup,))
        g_row = c.fetchone()
        if g_row:
            gid = g_row[0]
        else:
            c.execute("INSERT INTO product_groups (name) VALUES (?)", (hedef_grup,))
            gid = c.lastrowid
        c.execute("UPDATE products SET name = ?, group_id = ? WHERE id = ?", (yeni_ad, gid, pid))
        c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (yeni_ad, name))
    conn.commit()
    conn.close()


def whatsapp_grubunu_tamamen_temizle():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    c.execute("SELECT id FROM product_groups WHERE name = 'GRUPSUZLAR' LIMIT 1")
    g_row = c.fetchone()
    if g_row:
        grupsuz_id = g_row[0]
    else:
        c.execute("INSERT INTO product_groups (name) VALUES ('GRUPSUZLAR')")
        grupsuz_id = c.lastrowid

    c.execute("SELECT id FROM product_groups WHERE name = 'WHATSAPP' COLLATE NOCASE LIMIT 1")
    wp_row = c.fetchone()

    if wp_row:
        wp_grup_id = wp_row[0]
        c.execute("UPDATE products SET group_id = ? WHERE group_id = ?", (grupsuz_id, wp_grup_id))
        c.execute("DELETE FROM product_groups WHERE id = ?", (wp_grup_id,))

    c.execute("SELECT id FROM product_groups WHERE name LIKE '%GENEL POLİMER%' LIMIT 1")
    gp_row = c.fetchone()
    if gp_row:
        c.execute("UPDATE products SET group_id = ? WHERE group_id = ?", (grupsuz_id, gp_row[0]))
        c.execute("DELETE FROM product_groups WHERE id = ?", (gp_row[0],))

    conn.commit()
    conn.close()


def derin_veritabani_temizligi():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM customers")
    musteriler = c.fetchall()
    TR_MAP = WhatsAppSatisAyristirici.TR_ISIM_DUZELTME
    for mid, name in musteriler:
        yeni_ad = re.sub(r'\b[a-zA-ZğüşıöçĞÜŞİÖÇ]\b', '', name)
        yeni_ad = re.sub(r'\b(DOLAR|BOYA|TL|USD)\b', '', yeni_ad, flags=re.IGNORECASE)
        yeni_ad = re.sub(r'\s+', ' ', yeni_ad).strip()
        kelimeler = [TR_MAP.get(k, k) for k in yeni_ad.split()]
        yeni_ad = " ".join(kelimeler).strip()
        if not yeni_ad or yeni_ad in ["MÜŞTERİ", "MUSTERI"]:
            yeni_ad = "MÜŞTERİSİZ SATIŞ"
        if yeni_ad != name:
            c.execute("UPDATE customers SET name = ? WHERE id = ?", (yeni_ad, mid))
            c.execute("UPDATE sales_history SET customer_name = ? WHERE customer_id = ?", (yeni_ad, mid))
    c.execute("DELETE FROM products WHERE name LIKE '%GENEL POLİMER%'")
    c.execute("DELETE FROM sales_history WHERE product_name LIKE '%GENEL POLİMER%'")
    c.execute("SELECT id FROM product_groups WHERE name = 'GRUPSUZLAR' LIMIT 1")
    g_row = c.fetchone()
    grupsuz_id = g_row[0] if g_row else None
    c.execute("SELECT id, name FROM products WHERE name NOT LIKE '%POM%' AND name LIKE '%GRANÜL%'")
    for pid, pname in c.fetchall():
        temiz_pname = re.sub(r'\bGRANÜL\b', '', pname).strip()
        temiz_pname = re.sub(r'\s+', ' ', temiz_pname)
        gid = grupsuz_id
        if "ABS" in temiz_pname:
            c.execute("SELECT id FROM product_groups WHERE name = 'ABS' LIMIT 1")
            r = c.fetchone()
            if r: gid = r[0]
        elif "HDPE" in temiz_pname:
            c.execute("SELECT id FROM product_groups WHERE name = 'HDPE' LIMIT 1")
            r = c.fetchone()
            if r: gid = r[0]
        elif "PA66" in temiz_pname:
            c.execute("SELECT id FROM product_groups WHERE name = 'PA66' LIMIT 1")
            r = c.fetchone()
            if r: gid = r[0]
        elif "PA6" in temiz_pname:
            c.execute("SELECT id FROM product_groups WHERE name = 'PA6' LIMIT 1")
            r = c.fetchone()
            if r: gid = r[0]
        elif "PP" in temiz_pname or "MOBLEN" in temiz_pname:
            c.execute("SELECT id FROM product_groups WHERE name = 'PP MOBLEN' LIMIT 1")
            r = c.fetchone()
            if r: gid = r[0]
        if gid is None:
            continue
        c.execute("UPDATE products SET name = ?, group_id = ? WHERE id = ?", (temiz_pname, gid, pid))
        c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (temiz_pname, pname))
    conn.commit()
    conn.close()


def pom_harici_tum_granulleri_kazi():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM products WHERE name NOT LIKE '%POM%' AND (name LIKE '%GRANÜL%' OR name LIKE '%GRANUR%')")
    hatalilar = c.fetchall()
    for pid, eski_ad in hatalilar:
        yeni_ad = re.sub(r'\b(GRANÜL|GRANUR|GRANUL|GRANÜR)\b', '', eski_ad, flags=re.IGNORECASE)
        yeni_ad = re.sub(r'\s+', ' ', yeni_ad).strip()
        c.execute("UPDATE products SET name = ? WHERE id = ?", (yeni_ad, pid))
        c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (yeni_ad, eski_ad))
    conn.commit()
    conn.close()


def pom_istilasini_kurtar_ve_gruplari_duzelt():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, name FROM product_groups")
    gruplar = {row[1].upper(): row[0] for row in c.fetchall()}
    def grup_id_al(g_ad):
        if g_ad in gruplar:
            return gruplar[g_ad]
        c.execute("INSERT INTO product_groups (name) VALUES (?)", (g_ad,))
        gid = c.lastrowid
        gruplar[g_ad] = gid
        return gid
    c.execute("SELECT id, name FROM products")
    urunler = c.fetchall()
    for pid, name in urunler:
        eski_ad = name
        yeni_ad = name
        if "POM" not in yeni_ad.upper():
            yeni_ad = re.sub(r'\b(GRANÜL|GRANUR|GRANUL|GRANÜR)\b', '', yeni_ad, flags=re.IGNORECASE)
            yeni_ad = re.sub(r'\s+', ' ', yeni_ad).strip()
        if "GENEL POLİMER" in yeni_ad.upper():
            yeni_ad = re.sub(r'\bGENEL\s+POLİMER\b', '', yeni_ad, flags=re.IGNORECASE).strip()
            if not yeni_ad:
                yeni_ad = "KALSİTLİ PP"
        u_up = yeni_ad.upper()
        if "POM" in u_up:
            dogru_grup = "POM"
        elif "ABS" in u_up:
            dogru_grup = "ABS"
        elif any(k in u_up for k in ["HDPE", "ELTEKS", "I20", "PE100"]):
            dogru_grup = "HDPE"
        elif any(k in u_up for k in ["PP", "MOBLEN"]):
            dogru_grup = "PP MOBLEN"
        elif "PA66" in u_up or "N66" in u_up:
            dogru_grup = "PA66"
        elif "PA6" in u_up or "N6" in u_up:
            dogru_grup = "PA6"
        elif "PVC" in u_up:
            dogru_grup = "PVC"
        elif any(k in u_up for k in ["ANTİŞOK", "HIPS", "ANT"]):
            dogru_grup = "ANTİŞOK"
        else:
            dogru_grup = "GRUPSUZLAR"
        gid = grup_id_al(dogru_grup)
        c.execute("UPDATE products SET name = ?, group_id = ? WHERE id = ?", (yeni_ad, gid, pid))
        if eski_ad != yeni_ad:
            c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (yeni_ad, eski_ad))
    conn.commit()
    conn.close()


def alattin_mustafa_varyasyonlarini_birlestir():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    c = conn.cursor()
    c.execute("""
        SELECT id, name, debt, payment, remaining_debt, shopping_count
        FROM customers
        WHERE name LIKE '%ALAT%' OR name LIKE '%ALAAT%' OR name LIKE '%MUSTAF%'
    """)
    rows = c.fetchall()
    if not rows:
        conn.close()
        return
    gruplar = defaultdict(list)
    for r in rows:
        mid, name, debt, pay, rem, shop = r
        kanonik = MusteriKumeleyici.kanonik_musteri_anahtari(name)
        if not kanonik or kanonik == "musterisiz":
            continue
        gruplar[kanonik].append(r)
    for kanonik, musteri_listesi in gruplar.items():
        if len(musteri_listesi) <= 1:
            continue
        lider = max(musteri_listesi, key=lambda x: (len(x[1]), x[5]))
        lider_id = lider[0]
        lider_ad = "ALAATTİN MUSTAFA" if "mustafa" in kanonik and "alaattin" in kanonik else lider[1]
        toplam_debt = sum(float(x[2] or 0.0) for x in musteri_listesi)
        toplam_pay = sum(float(x[3] or 0.0) for x in musteri_listesi)
        toplam_rem = sum(float(x[4] or 0.0) for x in musteri_listesi)
        toplam_shop = sum(int(x[5] or 0) for x in musteri_listesi)
        c.execute("""
            UPDATE customers
            SET name = ?, debt = ?, payment = ?, remaining_debt = ?, shopping_count = ?
            WHERE id = ?
        """, (lider_ad, toplam_debt, toplam_pay, toplam_rem, toplam_shop, lider_id))
        for diger in musteri_listesi:
            diger_id = diger[0]
            if diger_id != lider_id:
                c.execute("""
                    UPDATE sales_history
                    SET customer_id = ?, customer_name = ?
                    WHERE customer_id = ?
                """, (lider_id, lider_ad, diger_id))
                c.execute("DELETE FROM customers WHERE id = ?", (diger_id,))
    conn.commit()
    conn.close()


def veritabanindaki_personel_ve_tarihleri_kurtar():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    c = conn.cursor()
    c.execute("""
        SELECT id, personnel_name, created_at, note
        FROM sales_history
        WHERE personnel_name LIKE '%[%' OR personnel_name LIKE '%/%' OR personnel_name LIKE '%:%'
    """)
    bozuk_kayitlar = c.fetchall()
    for s_id, bozuk_personel, mevcut_tarih, not_metni in bozuk_kayitlar:
        t_match = re.search(r'(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})', str(bozuk_personel or ""))
        yeni_tarih = mevcut_tarih
        if t_match:
            ham_t = t_match.group(1)
            parcalar = re.split(r'[./-]', ham_t)
            if len(parcalar) == 3:
                p1, p2, p3 = parcalar
                p3 = "20" + p3 if len(p3) == 2 else p3
                if int(p1) <= 12 and int(p2) > 12:
                    yeni_tarih = f"{p2.zfill(2)}.{p1.zfill(2)}.{p3}"
                else:
                    yeni_tarih = f"{p1.zfill(2)}.{p2.zfill(2)}.{p3}"
        yeni_personel = "AHMET"
        if not_metni:
            n_up = str(not_metni).upper()
            if "BABA" in n_up:
                yeni_personel = "BABA"
            elif "SİZ" in n_up or "SIZ" in n_up:
                yeni_personel = "MEHMET"
            elif "AHMET" in n_up:
                yeni_personel = "AHMET"
        c.execute("""
            UPDATE sales_history
            SET created_at = ?, personnel_name = ?
            WHERE id = ?
        """, (yeni_tarih, yeni_personel, s_id))
    conn.commit()
    conn.close()


def ucuk_fiyatlari_tavana_cek():
    conn = sqlite3.connect(DB_NAME, timeout=15)
    c = conn.cursor()
    c.execute("SELECT id, qty, price FROM sales_history WHERE price > 200.0")
    for sid, kg, eski_fiyat in c.fetchall():
        yeni_fiyat = 200.0
        yeni_tutar = float(kg or 0.0) * yeni_fiyat
        c.execute(
            "UPDATE sales_history SET price = ?, total_amount = ? WHERE id = ?",
            (yeni_fiyat, yeni_tutar, sid),
        )
    c.execute("SELECT id FROM customers")
    for (cid,) in c.fetchall():
        c.execute(
            "SELECT SUM(total_amount), SUM(payment_received) FROM sales_history WHERE customer_id = ?",
            (cid,),
        )
        row = c.fetchone()
        t_debt = row[0] or 0.0
        t_pay = row[1] or 0.0
        t_rem = max(0.0, t_debt - t_pay)
        c.execute(
            "UPDATE customers SET debt = ?, payment = ?, remaining_debt = ? WHERE id = ?",
            (t_debt, t_pay, t_rem, cid),
        )
    conn.commit()
    conn.close()


def gecmis_satis_ve_raporlari_sifirla():
    """Tarihsel ve ürünsel raporda kalan tüm eski satış kayıtlarını temizler"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM sales_history")
        try:
            c.execute("DELETE FROM satislar")
        except Exception:
            pass
        conn.commit()
        conn.close()
        print("Satış geçmişi ve raporlar tamamen sıfırlandı.")
    except Exception as e:
        print(f"Rapor sıfırlama hatası: {e}")


def tertemiz_musteri_listesine_gecis():
    """Çöp olan tüm müşterileri siler ve sadece kullanıcının doğru XLSX dosyasını içeri aktarır."""
    excel_yolu = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Müşteriler-BenimPOS_2.xlsx")
    if not os.path.exists(excel_yolu):
        print(f"Excel aktarım hatası: dosya yok ({excel_yolu})")
        return
    try:
        temiz_isimler = []
        if OPENPYXL_VAR:
            wb = openpyxl.load_workbook(excel_yolu, read_only=True, data_only=True)
            ws = wb.active
            basliklar = [str(h.value).strip() if h.value is not None else "" for h in next(ws.iter_rows(min_row=2, max_row=2))]
            musteri_idx = None
            for i, h in enumerate(basliklar):
                if h.lower() in ("müşteri", "musteri", "ad", "adı", "name"):
                    musteri_idx = i
                    break
            if musteri_idx is None:
                musteri_idx = 0
            for row in ws.iter_rows(min_row=3, values_only=True):
                if not row or musteri_idx >= len(row) or row[musteri_idx] is None:
                    continue
                isim = str(row[musteri_idx]).strip()
                if isim:
                    temiz_isimler.append(isim)
            wb.close()
        else:
            import pandas as pd
            df_temiz = pd.read_excel(excel_yolu, header=1)
            col = "Müşteri" if "Müşteri" in df_temiz.columns else df_temiz.columns[0]
            temiz_isimler = df_temiz[col].dropna().tolist()
        if not temiz_isimler:
            print("Excel aktarım hatası: müşteri sütunu boş, silme yapılmadı.")
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM customers")
        for isim in temiz_isimler:
            isim = str(isim).strip().upper()
            c.execute(
                "INSERT INTO customers (name, debt, payment, remaining_debt, shopping_count) VALUES (?, 0, 0, 0, 0)",
                (isim,),
            )
        conn.commit()
        conn.close()
        print(f"BÜYÜK TEMİZLİK TAMAMLANDI! Eski müşteriler silindi, {len(temiz_isimler)} gerçek müşteri eklendi.")
    except Exception as e:
        print(f"Excel aktarım hatası: {e}")


def super_agresif_musteri_birlestir():
    """Takoz Alaattin, Alaattin Dolar, Baypen Kalemcilik gibi parçalanmış carileri kök isimlerinde birleştirir."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
        musteriler = c.fetchall()
        gruplar = {}
        for m in musteriler:
            mid, ad, debt, payment, remaining, sayim = m
            ad_up = str(ad or "").upper().replace("İ", "I")
            if "ALAT" in ad_up or "ALAAT" in ad_up or "ALLAT" in ad_up:
                lider = "ALAATTİN"
            elif "BAYP" in ad_up or "RAYP" in ad_up:
                lider = "BAYPEN"
            elif "CEYH" in ad_up:
                lider = "CEYHAN"
            else:
                lider = ad
            gruplar.setdefault(lider, []).append(m)
        for lider_ad, kayitlar in gruplar.items():
            if len(kayitlar) == 1 and kayitlar[0][1] == lider_ad:
                continue
            ana_id = min(k[0] for k in kayitlar)
            top_debt = sum(float(k[2] or 0.0) for k in kayitlar)
            top_pay = sum(float(k[3] or 0.0) for k in kayitlar)
            top_rem = sum(float(k[4] or 0.0) for k in kayitlar)
            top_shop = sum(int(k[5] or 0) for k in kayitlar)
            c.execute(
                """UPDATE customers SET name=?, debt=?, payment=?, remaining_debt=?, shopping_count=? WHERE id=?""",
                (lider_ad, top_debt, top_pay, top_rem, top_shop, ana_id),
            )
            for k in kayitlar:
                diger_id = k[0]
                if diger_id != ana_id:
                    c.execute(
                        "UPDATE sales_history SET customer_id=?, customer_name=? WHERE customer_id=?",
                        (ana_id, lider_ad, diger_id),
                    )
                    try:
                        c.execute(
                            "UPDATE finans_evraklar SET customer_id=?, sahibi_kesideci=? WHERE customer_id=?",
                            (ana_id, lider_ad, diger_id),
                        )
                    except Exception:
                        pass
                    c.execute("DELETE FROM customers WHERE id=?", (diger_id,))
        conn.commit()
        conn.close()
        print("Mükerrer Alaattin, Baypen ve Ceyhan müşterileri tek hesapta birleştirildi.")
    except Exception as e:
        print(f"Birleştirme hatası oluştu: {e}")


def ramazan_ve_orhan_birlestir():
    """Dağılmış Ramazan ve Orhan Bahçe kayıtlarını doğru ana carilerde birleştirir."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
        musteriler = c.fetchall()

        gruplar = {}
        for m in musteriler:
            mid, ad, debt, payment, remaining, sayim = m
            ad_up = str(ad or "").upper().replace("İ", "I")

            hedef = None
            # 1. ORHAN BAHÇE
            if "ORHAN" in ad_up and "BAH" in ad_up:
                hedef = "ORHAN BAHÇE"

            # 2. RAMAZANLAR
            elif "RAMAZAN" in ad_up:
                if "EFOR" in ad_up:
                    hedef = "EFOR PLASTİK RAMAZAN"
                elif "ELEKT" in ad_up:
                    hedef = "RAMAZAN ELEKTRİKÇİ"
                elif "IKITELLI" in ad_up or "BAY" in ad_up:
                    hedef = "İKİTELLİ RAMAZAN"
                else:
                    # Komşu, Üstkat, Üstkomşu, Orjinal Ramazan Komşu vb. hepsi ana Ramazan'a
                    hedef = "RAMAZAN"

            if hedef:
                gruplar.setdefault(hedef, []).append(m)

        for lider_ad, kayitlar in gruplar.items():
            if len(kayitlar) <= 1 and kayitlar[0][1] == lider_ad:
                continue

            # En küçük ID ana hesap kabul edilir
            ana_id = min(k[0] for k in kayitlar)
            top_debt = sum(float(k[2] or 0.0) for k in kayitlar)
            top_pay = sum(float(k[3] or 0.0) for k in kayitlar)
            top_rem = sum(float(k[4] or 0.0) for k in kayitlar)
            top_shop = sum(int(k[5] or 0) for k in kayitlar)

            # Ana hesabı güncelle
            c.execute("""UPDATE customers SET name=?, debt=?, payment=?, remaining_debt=?, shopping_count=? WHERE id=?""",
                      (lider_ad, top_debt, top_pay, top_rem, top_shop, ana_id))

            # Geçmiş satışları ve finans evraklarını ana hesaba bağla, kopyaları sil
            for k in kayitlar:
                diger_id = k[0]
                if diger_id != ana_id:
                    c.execute("UPDATE sales_history SET customer_id=?, customer_name=? WHERE customer_id=?", (ana_id, lider_ad, diger_id))
                    try:
                        c.execute("UPDATE finans_evraklar SET customer_id=?, sahibi_kesideci=? WHERE customer_id=?", (ana_id, lider_ad, diger_id))
                    except Exception:
                        pass
                    c.execute("DELETE FROM customers WHERE id=?", (diger_id,))

        conn.commit()
        conn.close()
        print("✅ Ramazan ve Orhan Bahçe kayıtları başarıyla teke düşürüldü!")
    except Exception as e:
        print(f"Birleştirme hatası: {e}")


def musteri_birlestir(ana_musteri_id, birlestirilecek_id_listesi):
    """
    Kopya müşterilerin tüm satış ve kasa hareketlerini ana müşteriye aktarır,
    ardından kopya kayıtları temizler.
    """
    conn = sqlite3.connect(DB_NAME, timeout=15)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT name FROM customers WHERE id = ?", (ana_musteri_id,))
        ana = cursor.fetchone()
        if not ana:
            return False, "Ana müşteri bulunamadı."
        ana_ad = ana[0]

        for kopya_id in birlestirilecek_id_listesi:
            if kopya_id == ana_musteri_id:
                continue
            cursor.execute(
                "UPDATE sales_history SET customer_id = ?, customer_name = ? WHERE customer_id = ?",
                (ana_musteri_id, ana_ad, kopya_id),
            )
            try:
                cursor.execute(
                    "UPDATE finans_evraklar SET customer_id = ?, sahibi_kesideci = ? WHERE customer_id = ?",
                    (ana_musteri_id, ana_ad, kopya_id),
                )
            except Exception:
                pass
            cursor.execute("DELETE FROM customers WHERE id = ?", (kopya_id,))

        cursor.execute(
            "SELECT COUNT(*), COALESCE(SUM(total_amount), 0), COALESCE(SUM(payment_received), 0) FROM sales_history WHERE customer_id = ?",
            (ana_musteri_id,),
        )
        sayim, t_debt, t_pay = cursor.fetchone()
        try:
            cursor.execute(
                "SELECT COALESCE(SUM(tutar), 0) FROM finans_evraklar WHERE customer_id = ?",
                (ana_musteri_id,),
            )
            t_pay = float(t_pay or 0) + float(cursor.fetchone()[0] or 0)
        except Exception:
            t_pay = float(t_pay or 0)
        t_debt = float(t_debt or 0)
        t_rem = max(0.0, t_debt - t_pay)
        cursor.execute(
            """UPDATE customers
               SET shopping_count = ?, debt = ?, payment = ?, remaining_debt = ?
               WHERE id = ?""",
            (int(sayim or 0), t_debt, t_pay, t_rem, ana_musteri_id),
        )

        conn.commit()
        return True, "Müşteriler başarıyla birleştirildi."
    except Exception as e:
        conn.rollback()
        return False, f"Hata oluştu: {str(e)}"
    finally:
        conn.close()


def kristalleri_gruba_tasi_ve_birlestir():
    """Grupsuzlar'daki kristalleri KRİSTAL grubuna aktarır, Çapak Kristal'i Renkli Kristal yapar."""
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()

        c.execute("SELECT id FROM product_groups WHERE name = 'KRİSTAL' COLLATE NOCASE LIMIT 1")
        row = c.fetchone()
        if row:
            gid = row[0]
        else:
            c.execute("INSERT INTO product_groups (name) VALUES ('KRİSTAL')")
            gid = c.lastrowid

        c.execute("""
            UPDATE sales_history
            SET product_name = 'RENKLİ KRİSTAL'
            WHERE product_name LIKE '%ÇAPAK%KRİSTAL%'
               OR product_name LIKE '%ÇAPAK%KRISTAL%'
               OR REPLACE(UPPER(IFNULL(product_name,'')), 'İ', 'I') LIKE '%CAPAK%KRISTAL%'
        """)

        c.execute("SELECT id FROM products WHERE name = 'RENKLİ KRİSTAL' LIMIT 1")
        renkli = c.fetchone()
        c.execute("""
            SELECT id FROM products
            WHERE name LIKE '%ÇAPAK%KRİSTAL%'
               OR name LIKE '%ÇAPAK%KRISTAL%'
               OR REPLACE(UPPER(IFNULL(name,'')), 'İ', 'I') LIKE '%CAPAK%KRISTAL%'
        """)
        for (pid,) in c.fetchall():
            if renkli and pid != renkli[0]:
                c.execute("DELETE FROM products WHERE id = ?", (pid,))
            else:
                c.execute("UPDATE products SET name = 'RENKLİ KRİSTAL', group_id = ? WHERE id = ?", (gid, pid))
                renkli = (pid,)

        c.execute("""
            UPDATE products
            SET group_id = ?
            WHERE name LIKE '%KRİSTAL%' OR name LIKE '%KRISTAL%'
        """, (gid,))

        try:
            c.execute("""
                UPDATE sales_history
                SET group_name = 'KRİSTAL'
                WHERE product_name LIKE '%KRİSTAL%' OR product_name LIKE '%KRISTAL%'
            """)
        except Exception:
            pass

        c.execute("SELECT id, name FROM products WHERE name LIKE '%KRİSTAL%' OR name LIKE '%KRISTAL%'")
        isimler = {}
        for pid, ad in c.fetchall():
            isimler.setdefault(ad, []).append(pid)
        for ids in isimler.values():
            for diger in ids[1:]:
                c.execute("DELETE FROM products WHERE id = ?", (diger,))

        conn.commit()
        conn.close()
        print("✅ Kristaller başarıyla KRİSTAL grubuna taşındı, Çapak Kristal -> Renkli Kristal yapıldı!")
    except Exception as e:
        print(f"Kristal düzeltme hatası: {e}")


def mutasan_tek_hesap_yap():
    """Bölünmüş tüm Mutasan kayıtlarını satış ve ödeme geçmişiyle birlikte tek hesaba toplar."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers WHERE name LIKE '%MUTASAN%'")
        kayitlar = c.fetchall()

        if len(kayitlar) <= 1:
            conn.close()
            return

        ana_id = min(k[0] for k in kayitlar)
        lider_ad = "MUTASAN"

        top_debt = sum(float(k[2] or 0.0) for k in kayitlar)
        top_pay = sum(float(k[3] or 0.0) for k in kayitlar)
        top_rem = sum(float(k[4] or 0.0) for k in kayitlar)
        top_shop = sum(int(k[5] or 0) for k in kayitlar)

        c.execute("""UPDATE customers SET name=?, debt=?, payment=?, remaining_debt=?, shopping_count=? WHERE id=?""",
                  (lider_ad, top_debt, top_pay, top_rem, top_shop, ana_id))

        for k in kayitlar:
            diger_id = k[0]
            if diger_id != ana_id:
                c.execute("UPDATE sales_history SET customer_id=?, customer_name=? WHERE customer_id=?", (ana_id, lider_ad, diger_id))
                try:
                    c.execute("UPDATE finans_evraklar SET customer_id=?, sahibi_kesideci=? WHERE customer_id=?", (ana_id, lider_ad, diger_id))
                except Exception:
                    pass
                c.execute("DELETE FROM customers WHERE id=?", (diger_id,))

        conn.commit()
        conn.close()
        print("✅ Başarılı! Tüm Mutasanlar tek bir kurumsal hesapta birleştirildi.")
    except Exception as e:
        print(f"Mutasan birleştirme hatası: {e}")


def evrensel_yapay_zeka_birlestirici():
    """
    Hiçbir firma ismine veya sözlüğe ihtiyaç duymadan;
    'Bursa Can Plastik', 'Can Pls', 'Can Plast' gibi benzer isimleri matematiksel olarak
    tespit eder ve en uzun/en düzgün ismin altında tüm hesapları birleştirir.
    """
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
        musteriler = c.fetchall()

        if not musteriler:
            conn.close()
            return

        cop_kelimeler = {"PLASTIK", "PLASTİK", "PLS", "PLAST", "SANAYI", "TICARET", "LTD", "STI",
                         "ABI", "BEY", "USTA", "KARDES", "ADANA", "BURSA", "IKITELLI", "TOPCULAR", "MERTER", "SEYHAN"}

        def saf_koku_bul(isim):
            isim = str(isim or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
            kelimeler = [w for w in isim.split() if w not in cop_kelimeler and len(w) > 2]
            return " ".join(kelimeler)

        kumeler = defaultdict(list)
        islenen_id_ler = set()

        for m1 in musteriler:
            m1_id, m1_ad = m1[0], m1[1]
            if m1_id in islenen_id_ler:
                continue

            m1_kok = saf_koku_bul(m1_ad)
            grup = [m1]
            islenen_id_ler.add(m1_id)

            if m1_kok:
                for m2 in musteriler:
                    m2_id, m2_ad = m2[0], m2[1]
                    if m2_id in islenen_id_ler:
                        continue

                    m2_kok = saf_koku_bul(m2_ad)
                    if not m2_kok:
                        continue

                    if m1_kok == m2_kok or (len(m1_kok) > 3 and len(m2_kok) > 3 and difflib.SequenceMatcher(None, m1_kok, m2_kok).ratio() > 0.75):
                        grup.append(m2)
                        islenen_id_ler.add(m2_id)

            kumeler[m1_ad] = grup

        birlestirilen_grup_sayisi = 0
        for _, kayitlar in kumeler.items():
            if len(kayitlar) <= 1:
                continue

            birlestirilen_grup_sayisi += 1
            lider = max(kayitlar, key=lambda k: len(str(k[1])))
            lider_id, lider_ad = lider[0], lider[1]

            top_debt = sum(float(k[2] or 0.0) for k in kayitlar)
            top_pay = sum(float(k[3] or 0.0) for k in kayitlar)
            top_rem = sum(float(k[4] or 0.0) for k in kayitlar)
            top_shop = sum(int(k[5] or 0) for k in kayitlar)

            c.execute("""UPDATE customers SET name=?, debt=?, payment=?, remaining_debt=?, shopping_count=? WHERE id=?""",
                      (lider_ad, top_debt, top_pay, top_rem, top_shop, lider_id))

            for k in kayitlar:
                diger_id = k[0]
                if diger_id != lider_id:
                    c.execute("UPDATE sales_history SET customer_id=?, customer_name=? WHERE customer_id=?", (lider_id, lider_ad, diger_id))
                    try:
                        c.execute("UPDATE finans_evraklar SET customer_id=?, sahibi_kesideci=? WHERE customer_id=?", (lider_id, lider_ad, diger_id))
                    except Exception:
                        pass
                    c.execute("DELETE FROM customers WHERE id=?", (diger_id,))

        conn.commit()
        conn.close()
        print(f"✅ MUHTEŞEM! Hiçbir isme ihtiyaç duymadan {birlestirilen_grup_sayisi} farklı kopya müşteri grubu yapay zeka ile teke düşürüldü!")
    except Exception as e:
        print(f"Evrensel birleştirme hatası: {e}")


def temizle_ve_normalizasyon_yap():
    """Müşteri isimlerindeki fiil/cümle kalıntılarını temizler; TURUNCU gibi hatalı kayıtları siler."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name FROM customers")
        musteriler = c.fetchall()

        fiil_ve_cumle_ekleri = [
            "ATICAKMIS", "ATACAKMIS", "ATICAK", "ATACAK", "MISMIS", "DIK", "DILER",
            "BEGENMIS", "GOTURCEKMISSIN", "ORTEGI", "ORTAGIYDI", "YAZDIGINI", "ISIM",
            "CIRO", "TOPLAM", "TUTAR", "FIYAT", "METIN", "DUZENLENDI",
        ]

        def ascii_up(metin):
            return str(metin or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")

        for mid, ad in musteriler:
            if not ad:
                continue
            orijinal_ad = str(ad).strip()
            ad_up = ascii_up(orijinal_ad)

            if ad_up in ["TURUNCU", "KIRMIZI", "MAVI", "YEDEK", "TEST", "MUSTERI"]:
                c.execute("UPDATE sales_history SET customer_id=NULL, customer_name='MÜŞTERİSİZ SATIŞ' WHERE customer_id=?", (mid,))
                c.execute("DELETE FROM customers WHERE id = ?", (mid,))
                continue

            temiz_kelimeler = []
            for kelime in orijinal_ad.split():
                k_up = ascii_up(kelime)
                if any(ek in k_up for ek in fiil_ve_cumle_ekleri):
                    continue
                if len(k_up) < 2 and not k_up.isdigit():
                    continue
                temiz_kelimeler.append(kelime)

            yeni_ad = " ".join(temiz_kelimeler).strip()
            if not yeni_ad or len(yeni_ad) < 2:
                c.execute("UPDATE sales_history SET customer_id=NULL, customer_name='MÜŞTERİSİZ SATIŞ' WHERE customer_id=?", (mid,))
                c.execute("DELETE FROM customers WHERE id = ?", (mid,))
            elif yeni_ad != orijinal_ad:
                c.execute("UPDATE customers SET name = ? WHERE id = ?", (yeni_ad, mid))
                c.execute("UPDATE sales_history SET customer_name = ? WHERE customer_id = ?", (yeni_ad, mid))

        conn.commit()
        conn.close()
        print("✅ Türkçe isim kurallarına uymayan cümle kalıntıları ve hatalı 'Turuncu' gibi kayıtlar temizlendi!")
    except Exception as e:
        print(f"İsim temizleme hatası: {e}")


BILINEN_HAMMADDELER = {
    "POM": ["POM", "DELRIN", "DERLIN", "DELDIM", "ASETAL", "KEPITAL", "KOCETAL"],
    "PP MOBLEN": ["PP", "MOBLEN", "POLIPROPILEN"],
    "ABS": ["ABS", "TERLURAN", "MAGNUM"],
    "ANTİŞOK": ["ANTISOK", "ANTİŞOK", "HIPS", "ANT"],
    "PA66": ["PA66", "PA 66", "N66", "NAYLON 66", "ULTRAMID A", "ZYTEL 101"],
    "PA6": ["PA6", "PA 6", "N6", "NAYLON 6", "ULTRAMID", "AKROMID", "DURETHAN", "POLIAMID", "NAYLON"],
    "HDPE": ["HDPE", "PE", "POLIETILEN", "ELTEKS", "I20", "PE100"],
    "LDPE": ["LDPE", "ALCAK YOGUNLUK", "F220", "G03", "SERA NAYLONU"],
    "PVC": ["PVC", "POLIVINIL"],
    "KRİSTAL": ["KRISTAL", "KRİSTAL", "PS", "GPPS", "KRS"],
    "PC": ["PC", "POLIKARBON", "POLICARBON", "LEXAN", "MAKROLON"],
    "PET": ["PET", "PETE", "POLIESTER", "POLYESTER", "PET-G", "PETG", "PREFORM"],
    "ASA": ["ASA", "LURAN S", "GELOY"],
    "PBT": ["PBT", "VALOX", "POCAN", "ULTRADUR"],
    "SAN": ["SAN", "LURAN", "TYRIL"],
    "PMMA": ["PMMA", "PLEKSI", "PLEKSIGLAS", "AKRILIK", "ALTUGLAS"],
    "TPU": ["TPU", "TERMOPLASTIK POLIURETAN", "POLIURETAN"],
    "TPE": ["TPE", "TERMOPLASTIK ELASTOMER", "ELASTOMER"],
    "EVA": ["EVA", "ETILEN VINIL ASETAT"],
    "KAUÇUK": ["KAUCUK", "KAUÇUK", "EPDM", "NBR", "SBR"],
    "PA12": ["PA12", "PA 12", "GRILAMID", "RILSAN"],
    "PPS": ["PPS", "FORTRON", "RYTON"],
    "PEEK": ["PEEK", "VICTREX"],
    "PLA": ["PLA", "POLILAKTIK ASIT"],
}


def hammadde_grubu_ve_adi_coz(metin):
    ham = str(metin or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
    tespit_edilen_grup = None
    for ana_grup, varyasyonlar in BILINEN_HAMMADDELER.items():
        if any(v.replace("İ", "I") in ham.split() or v.replace("İ", "I") in ham for v in varyasyonlar):
            tespit_edilen_grup = ana_grup
            break
    if tespit_edilen_grup:
        renk = "NATUREL"
        if "SIYAH" in ham:
            renk = "SİYAH"
        elif "BEYAZ" in ham:
            renk = "BEYAZ"
        elif "RENKLI" in ham or "CAPAK" in ham:
            renk = "RENKLİ"
        elif "ORJ" in ham:
            renk = "ORJİNAL"
        return tespit_edilen_grup, f"{renk} {tespit_edilen_grup}"
    return "GRUPSUZLAR", str(metin or "").strip().upper()


def yeni_hammadde_gruplarini_tanimla():
    """PET, ASA, PBT, TPU, EVA vb. grupları eksiksiz olarak veritabanına açar."""
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()
        for grup in BILINEN_HAMMADDELER.keys():
            c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (grup,))
            if not c.fetchone():
                c.execute("INSERT INTO product_groups (name) VALUES (?)", (grup,))
        conn.commit()
        conn.close()
        print("✅ 24 hammadde grubu başarıyla sisteme tanımlandı.")
    except Exception as e:
        print(f"Grup tanımlama hatası: {e}")


def grupsuzlar_ve_kristal_mantigini_kesin_duzelt():
    """
    1. KRİSTAL grubunu açar, tüm kristalleri Grupsuzlar'dan çıkarıp KRİSTAL grubuna taşır.
    2. 'BEYAZ DİĞER', 'ÇAPAK DİĞER' gibi saçma isimleri temizler.
    3. Grupsuzlar'a yalnızca tekil hammadde adı girilmiş (Örn: 'PET', 'KAUÇUK') kayıtları bırakır.
    """
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()

        c.execute("SELECT id FROM product_groups WHERE UPPER(name) IN ('KRİSTAL', 'KRISTAL') LIMIT 1")
        k_row = c.fetchone()
        if not k_row:
            c.execute("INSERT INTO product_groups (name) VALUES ('KRİSTAL')")
            kristal_gid = c.lastrowid
        else:
            kristal_gid = k_row[0]

        c.execute("SELECT id FROM product_groups WHERE UPPER(name) = 'GRUPSUZLAR' LIMIT 1")
        g_row = c.fetchone()
        if not g_row:
            c.execute("INSERT INTO product_groups (name) VALUES ('GRUPSUZLAR')")
            grupsuz_gid = c.lastrowid
        else:
            grupsuz_gid = g_row[0]

        c.execute("""
            UPDATE products
            SET group_id = ?
            WHERE UPPER(name) LIKE '%KRİSTAL%'
               OR UPPER(name) LIKE '%KRISTAL%'
               OR UPPER(name) LIKE '% GPPS%'
        """, (kristal_gid,))

        c.execute("SELECT id, name FROM products WHERE UPPER(name) LIKE '%DİĞER%' OR UPPER(name) LIKE '%DIGER%'")
        diger_urunler = c.fetchall()
        for pid, ad in diger_urunler:
            ad_temiz = ad.upper().replace("DİĞER", "").replace("DIGER", "").strip()
            if ad_temiz in ["ÇAPAK", "CAPAK", "NATUREL", "BEYAZ", "SİYAH", "SIYAH", "RENKLİ", "RENKLI", "GRİ", "GRI", "ORJİNAL", "ORJ"]:
                yeni_ad = f"{ad_temiz} (HAMMADDESİZ)"
                c.execute("UPDATE products SET name = ?, group_id = ? WHERE id = ?", (yeni_ad, grupsuz_gid, pid))
                c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (yeni_ad, ad))
            else:
                c.execute("UPDATE products SET name = ? WHERE id = ?", (ad_temiz, pid))
                c.execute("UPDATE sales_history SET product_name = ? WHERE product_name = ?", (ad_temiz, ad))

        TEKIL_HAMMADDELER = ["PET", "KAUÇUK", "KAUCUK", "EVA", "TPU", "TPE", "PLA", "SEK"]
        for tekil in TEKIL_HAMMADDELER:
            c.execute("""
                UPDATE products
                SET group_id = ?
                WHERE UPPER(TRIM(name)) = ?
            """, (grupsuz_gid, tekil))

        conn.commit()
        conn.close()
        print("✅ Kristaller kendi grubuna ayrıldı, '... DİĞER' saçmalığı temizlendi!")
    except Exception as e:
        print(f"Kristal/DİĞER düzeltme hatası: {e}")


def veritabanini_grupsuz_sacmaligindan_kurtar():
    """Çapak Naturel gibi hatalı ürünleri gerçek gruplarına aktarır, Grupsuzlar'ı sadece tekil istisnalara bırakır."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name FROM products")
        for uid, ad in c.fetchall():
            grup, temiz_ad = hammadde_grubu_ve_adi_coz(ad)
            c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (grup,))
            row = c.fetchone()
            if row:
                gid = row[0]
            else:
                c.execute("INSERT INTO product_groups (name) VALUES (?)", (grup,))
                gid = c.lastrowid
            c.execute("UPDATE products SET name = ?, group_id = ? WHERE id = ?", (temiz_ad, gid, uid))

        c.execute("SELECT id, product_name FROM sales_history")
        for sid, pad in c.fetchall():
            grup, temiz_ad = hammadde_grubu_ve_adi_coz(pad)
            c.execute("UPDATE sales_history SET product_name = ? WHERE id = ?", (temiz_ad, sid))
            try:
                c.execute("UPDATE sales_history SET group_name = ? WHERE id = ?", (grup, sid))
            except Exception:
                pass

        c.execute("SELECT name, MIN(id) FROM products GROUP BY name HAVING COUNT(*) > 1")
        for ad, ana_id in c.fetchall():
            c.execute("DELETE FROM products WHERE name = ? AND id != ?", (ad, ana_id))

        conn.commit()
        conn.close()
        print("✅ Hammadde türleri ayrıştırıldı: Çapak/Naturel ekleri ait oldukları gruba bağlandı, Grupsuzlar temizlendi!")
    except Exception as e:
        print(f"Hata: {e}")


def guvenli_dinamik_birlestir():
    """Eski düzeltilmiş hesaplara dokunmadan, sadece harf hatası yapılmış düşük işlemli çöp kayıtları ana hesaplara bağlar."""
    if os.path.exists(DB_NAME):
        shutil.copyfile(DB_NAME, DB_NAME + ".guvenli_yedek")
        print("🛡️ Mevcut veritabanının yedeği alındı (.guvenli_yedek).")

    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
        musteriler = c.fetchall()

        DOKUNULMAZ_FARKLI_ISIMLER = {
            "ALICAN", "ALİCAN", "SINCANCI", "SİNCANCI", "SINCAN", "SİNCAN",
            "ORHAN", "BURHAN", "AYHAN", "CEYHAN", "ERHAN", "ILHAN", "İLHAN",
        }

        def duzenleme_farki(s1, s2):
            if len(s1) < len(s2):
                return duzenleme_farki(s2, s1)
            if len(s2) == 0:
                return len(s1)
            prev = range(len(s2) + 1)
            for i, c1 in enumerate(s1):
                cur = [i + 1]
                for j, c2 in enumerate(s2):
                    cur.append(min(prev[j + 1] + 1, cur[j] + 1, prev[j] + (c1 != c2)))
                prev = cur
            return prev[-1]

        def temiz_kokler(ad):
            ad = str(ad or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
            cop = {"DELDIM", "DELRIN", "POM", "PLASTIK", "PLS", "SAN", "TIC", "LTD", "STI", "ABI", "BEY", "ADANA", "BURSA"}
            return [k for k in ad.split() if k not in cop and len(k) >= 3]

        islenenler = set()
        birlestirilen_grup_sayisi = 0
        sirali_musteriler = sorted(musteriler, key=lambda x: int(x[5] or 0), reverse=True)

        for ana in sirali_musteriler:
            ana_id, ana_ad, ana_debt, ana_pay, ana_rem, ana_shop = ana
            if ana_id in islenenler:
                continue
            ana_kelimeler = temiz_kokler(ana_ad)
            baglanacaklar = []
            for aday in sirali_musteriler:
                aday_id, aday_ad, a_debt, a_pay, a_rem, a_shop = aday
                if aday_id == ana_id or aday_id in islenenler:
                    continue
                if int(a_shop or 0) > 4 and int(ana_shop or 0) > 4:
                    continue
                aday_kelimeler = temiz_kokler(aday_ad)
                if any(k in DOKUNULMAZ_FARKLI_ISIMLER for k in aday_kelimeler) and any(k in DOKUNULMAZ_FARKLI_ISIMLER for k in ana_kelimeler):
                    if set(aday_kelimeler) != set(ana_kelimeler):
                        continue
                eslesme = False
                for w1 in ana_kelimeler:
                    for w2 in aday_kelimeler:
                        maks_u = max(len(w1), len(w2))
                        if maks_u < 4:
                            continue
                        fark = duzenleme_farki(w1, w2)
                        if maks_u >= 8 and fark <= 3:
                            eslesme = True
                        elif maks_u >= 6 and fark <= 2:
                            eslesme = True
                        elif maks_u >= 4 and fark <= 1:
                            eslesme = True
                        if len(w1) >= 5 and (w2.startswith(w1[:5]) or w1.startswith(w2[:5])) and fark <= 3:
                            eslesme = True
                        if eslesme:
                            break
                    if eslesme:
                        break
                if eslesme:
                    baglanacaklar.append(aday)
                    islenenler.add(aday_id)

            if baglanacaklar:
                birlestirilen_grup_sayisi += 1
                islenenler.add(ana_id)
                ek_debt = sum(float(k[2] or 0.0) for k in baglanacaklar)
                ek_pay = sum(float(k[3] or 0.0) for k in baglanacaklar)
                ek_rem = sum(float(k[4] or 0.0) for k in baglanacaklar)
                ek_shop = sum(int(k[5] or 0) for k in baglanacaklar)
                c.execute(
                    """UPDATE customers SET debt=?, payment=?, remaining_debt=?, shopping_count=? WHERE id=?""",
                    (
                        float(ana_debt or 0.0) + ek_debt,
                        float(ana_pay or 0.0) + ek_pay,
                        float(ana_rem or 0.0) + ek_rem,
                        int(ana_shop or 0) + ek_shop,
                        ana_id,
                    ),
                )
                for k in baglanacaklar:
                    d_id = k[0]
                    c.execute("UPDATE sales_history SET customer_id=?, customer_name=? WHERE customer_id=?", (ana_id, ana_ad, d_id))
                    try:
                        c.execute("UPDATE finans_evraklar SET customer_id=?, sahibi_kesideci=? WHERE customer_id=?", (ana_id, ana_ad, d_id))
                    except Exception:
                        pass
                    c.execute("DELETE FROM customers WHERE id=?", (d_id,))

        conn.commit()
        conn.close()
        print(f"✅ İşlem tamam! Eski cariler korunarak {birlestirilen_grup_sayisi} adet hatalı yazım ana hesaplara birleştirildi.")
    except Exception as e:
        print(f"Hata: {e}")


def tum_kopuk_satislari_gercek_musteriye_bagla():
    """Silinip yeniden yüklenen müşteriler yüzünden kopan bağları KESİN olarak onarır."""
    try:
        conn = sqlite3.connect(DB_NAME, timeout=30)
        c = conn.cursor()

        c.execute("DELETE FROM customers WHERE name IN ('SATI', 'BEN', 'YOK', 'ADAM')")

        c.execute("SELECT id FROM customers WHERE name = 'MÜŞTERİSİZ SATIŞ' LIMIT 1")
        row = c.fetchone()
        if not row:
            c.execute(
                "INSERT INTO customers (name, debt, payment, remaining_debt, shopping_count) VALUES ('MÜŞTERİSİZ SATIŞ', 0.0, 0.0, 0.0, 0)"
            )
            m_bos_id = c.lastrowid
        else:
            m_bos_id = row[0]

        c.execute("SELECT id, name FROM customers")
        musteriler = c.fetchall()

        c.execute(
            "SELECT id, customer_name FROM sales_history WHERE customer_id IS NULL OR customer_id NOT IN (SELECT id FROM customers)"
        )
        kopuk_satislar = c.fetchall()

        duzeltilen_satis = 0
        for sid, cname in kopuk_satislar:
            if not cname or cname == "MÜŞTERİSİZ SATIŞ":
                c.execute("UPDATE sales_history SET customer_id = ? WHERE id = ?", (m_bos_id, sid))
                duzeltilen_satis += 1
                continue

            cname_up = str(cname).upper().replace("İ", "I")
            hedef_id = m_bos_id
            hedef_ad = "MÜŞTERİSİZ SATIŞ"

            if "ALAT" in cname_up or "ALAAT" in cname_up or "ALLAT" in cname_up or "ALAD" in cname_up or "ALAAD" in cname_up:
                hedef_ad = "ALAATTİN"
            elif "BAYP" in cname_up or "RAYP" in cname_up:
                hedef_ad = "BAYPEN"
            elif "CEYH" in cname_up:
                hedef_ad = "CEYHAN"
            elif "MUTASAN" in cname_up:
                hedef_ad = "MUTASAN"
            else:
                hedef_ad = str(cname).strip().upper()

            eslesen = [m for m in musteriler if (m[1] or "").upper() == hedef_ad.upper()]
            if not eslesen:
                eslesen = [m for m in musteriler if hedef_ad.upper() in (m[1] or "").upper()]
            if eslesen:
                hedef_id = eslesen[0][0]
                hedef_ad = eslesen[0][1]
            elif hedef_ad != "MÜŞTERİSİZ SATIŞ":
                c.execute(
                    "INSERT INTO customers (name, debt, payment, remaining_debt, shopping_count) VALUES (?, 0.0, 0.0, 0.0, 0)",
                    (hedef_ad,),
                )
                hedef_id = c.lastrowid
                musteriler.append((hedef_id, hedef_ad))

            c.execute(
                "UPDATE sales_history SET customer_id = ?, customer_name = ? WHERE id = ?",
                (hedef_id, hedef_ad, sid),
            )
            duzeltilen_satis += 1

        c.execute(
            """
            UPDATE sales_history
            SET payment_received = 0.0, payment_type = 'AÇIK HESAP'
            WHERE customer_id = ? AND (payment_received IS NULL OR payment_type IS NULL OR payment_type = '')
            """,
            (m_bos_id,),
        )

        c.execute("SELECT id, name FROM customers")
        for m_id, m_name in c.fetchall():
            c.execute(
                """
                SELECT COUNT(*), COALESCE(SUM(total_amount), 0), COALESCE(SUM(payment_received), 0)
                FROM sales_history WHERE customer_id = ?
                """,
                (m_id,),
            )
            cnt, t_amount, t_pay = c.fetchone()
            kalan_borc = max(0.0, float(t_amount or 0) - float(t_pay or 0))
            c.execute(
                """
                UPDATE customers
                SET shopping_count = ?, debt = ?, payment = ?, remaining_debt = ?
                WHERE id = ?
                """,
                (cnt, t_amount, t_pay, kalan_borc, m_id),
            )

        conn.commit()
        conn.close()
        print(f"✅ BÜYÜK KURTARMA TAMAMLANDI! {duzeltilen_satis} adet kopuk satış gerçek sahiplerine bağlandı.")
    except Exception as e:
        print(f"Kurtarma Hatası: {e}")


def bosluksuz_ve_birlesik_isimleri_birlestir():
    """
    AZAYMETAL <-> AZAY METAL gibi aradaki boşluk/tire farkından dolayı
    bölünmüş tüm carileri ana hesap altında birleştirir.
    """
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()
        c.execute("SELECT id, name, debt, payment, remaining_debt, shopping_count FROM customers")
        musteriler = c.fetchall()
        if not musteriler:
            conn.close()
            return

        def kompakt_ad(ad):
            metin = str(ad or "").upper().replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").replace("Ü", "U")
            return "".join(ch for ch in metin if ch.isalnum())

        gruplar = {}
        for m in musteriler:
            k_ad = kompakt_ad(m[1])
            if not k_ad or k_ad in ["MUSTERISIZSATIS", "MUSTERI", "SATI"]:
                continue
            gruplar.setdefault(k_ad, []).append(m)

        birlestirilen_sayisi = 0
        for k_ad, kayitlar in gruplar.items():
            if len(kayitlar) <= 1:
                continue
            lider = max(kayitlar, key=lambda x: int(x[5] or 0))
            ana_id, ana_ad = lider[0], lider[1]
            top_debt = sum(float(k[2] or 0.0) for k in kayitlar)
            top_pay = sum(float(k[3] or 0.0) for k in kayitlar)
            top_rem = sum(float(k[4] or 0.0) for k in kayitlar)
            top_shop = sum(int(k[5] or 0) for k in kayitlar)
            c.execute("""
                UPDATE customers
                SET debt = ?, payment = ?, remaining_debt = ?, shopping_count = ?
                WHERE id = ?
            """, (top_debt, top_pay, top_rem, top_shop, ana_id))
            for k in kayitlar:
                diger_id = k[0]
                if diger_id != ana_id:
                    c.execute("UPDATE sales_history SET customer_id = ?, customer_name = ? WHERE customer_id = ?", (ana_id, ana_ad, diger_id))
                    try:
                        c.execute("UPDATE finans_evraklar SET customer_id = ?, sahibi_kesideci = ? WHERE customer_id = ?", (ana_id, ana_ad, diger_id))
                    except Exception:
                        pass
                    c.execute("DELETE FROM customers WHERE id = ?", (diger_id,))
                    birlestirilen_sayisi += 1

        conn.commit()
        conn.close()
        print(f"✅ İşlem tamamlandı: {birlestirilen_sayisi} adet birleşik/boşluksuz cari ana hesabına devredildi!")
    except Exception as e:
        print(f"Boşluksuz birleştirme hatası: {e}")


def fiil_ve_cop_carileri_temizle():
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()
        c.execute("SELECT id FROM customers WHERE name = 'MÜŞTERİSİZ SATIŞ' LIMIT 1")
        row = c.fetchone()
        if not row:
            c.execute("INSERT INTO customers (name, shopping_count, debt, payment, remaining_debt) VALUES (?, 0, 0, 0, 0)", ("MÜŞTERİSİZ SATIŞ",))
            m_bos_id = c.lastrowid
        else:
            m_bos_id = row[0]

        c.execute("SELECT id, name FROM customers")
        musteriler = c.fetchall()
        silinenler = 0
        for mid, name in musteriler:
            if mid == m_bos_id:
                continue
            if gecersiz_cari_mi(name):
                c.execute("UPDATE sales_history SET customer_id = ?, customer_name = 'MÜŞTERİSİZ SATIŞ' WHERE customer_id = ?", (m_bos_id, mid))
                try:
                    c.execute("UPDATE finans_evraklar SET customer_id = ?, sahibi_kesideci = 'MÜŞTERİSİZ SATIŞ' WHERE customer_id = ?", (m_bos_id, mid))
                except Exception:
                    pass
                c.execute("DELETE FROM customers WHERE id = ?", (mid,))
                silinenler += 1

        c.execute("SELECT COUNT(*), COALESCE(SUM(total_amount), 0), COALESCE(SUM(payment_received), 0) FROM sales_history WHERE customer_id = ?", (m_bos_id,))
        cnt, t_amt, t_pay = c.fetchone()
        c.execute(
            "UPDATE customers SET shopping_count = ?, debt = ?, payment = ?, remaining_debt = ? WHERE id = ?",
            (cnt, t_amt, t_pay, max(0.0, float(t_amt) - float(t_pay)), m_bos_id)
        )
        conn.commit()
        conn.close()
        print(f"✅ {silinenler} adet hatalı fiil/cümle carisi silindi ve satışları müşterisiz satışa aktarıldı.")
    except Exception as e:
        print(f"Fiil/çöp cari temizleme hatası: {e}")


def kayip_hammadde_kartlarini_kurtar():
    """sales_history'de satışı olan ancak products'ta olmayan hammaddelerin kartlarını oluşturur."""
    try:
        conn = sqlite3.connect(DB_NAME, timeout=15)
        c = conn.cursor()
        c.execute("SELECT DISTINCT product_name FROM sales_history WHERE product_name IS NOT NULL AND product_name != ''")
        satistaki_urunler = [r[0].strip().upper() for r in c.fetchall()]

        for urun_adi in satistaki_urunler:
            hedef_grup = "GRUPSUZLAR"
            en_len = 0
            for grup, anahtarlar in BILINEN_HAMMADDELER.items():
                for k in anahtarlar:
                    k_up = str(k).upper().replace("İ", "I")
                    if k_up and k_up in urun_adi.replace("İ", "I") and len(k_up) >= en_len:
                        if len(k_up) < 3 and k_up not in urun_adi.replace("İ", "I").split():
                            continue
                        if len(k_up) > en_len or (len(k_up) == en_len and hedef_grup == "GRUPSUZLAR"):
                            en_len = len(k_up)
                            hedef_grup = grup

            c.execute("SELECT id FROM product_groups WHERE name = ? COLLATE NOCASE LIMIT 1", (hedef_grup,))
            g_row = c.fetchone()
            if g_row:
                gid = g_row[0]
            else:
                c.execute("INSERT INTO product_groups (name) VALUES (?)", (hedef_grup,))
                gid = c.lastrowid

            c.execute("SELECT price FROM sales_history WHERE UPPER(product_name) = ? AND price > 0 ORDER BY id DESC LIMIT 1", (urun_adi,))
            p_row = c.fetchone()
            fiyat = float(p_row[0]) if p_row else 40.0

            c.execute("SELECT id FROM products WHERE UPPER(name) = ? LIMIT 1", (urun_adi,))
            mevcut = c.fetchone()
            if not mevcut:
                c.execute("""
                    INSERT INTO products (group_id, name, stock_kg, buy_price, sell_price, unit)
                    VALUES (?, ?, 0.0, 0.0, ?, 'KG')
                """, (gid, urun_adi, fiyat))
            else:
                c.execute("UPDATE products SET group_id = ? WHERE id = ?", (gid, mevcut[0]))

        conn.commit()
        conn.close()
        print("✅ Satış geçmişindeki tüm PC, PMMA, PET ürün kartları oluşturuldu ve gruplarına bağlandı!")
    except Exception as e:
        print(f"Kayıp hammadde kartı kurtarma hatası: {e}")


def wp_kayitlarini_denetle(dry_run=True):
    """Eski WP satışlarını eşleştiriciyle raporlar. Varsayılan dry-run; yedek alınmadan yazmaz."""
    conn = sqlite3.connect(DB_NAME, timeout=30)
    try:
        if not dry_run:
            yedek = DB_NAME + ".wp_onarim_yedek"
            with sqlite3.connect(yedek) as yedek_conn:
                conn.backup(yedek_conn)
            print(f"Yedek alındı: {yedek}")
        c = conn.cursor()
        c.execute("SELECT id, name, remaining_debt FROM customers")
        musteriler = c.fetchall()
        esleyici = MusteriEsleyici(musteriler)
        c.execute("SELECT id, customer_id, customer_name, product_name, qty, total_amount FROM sales_history")
        rapor = []
        for sid, cid, cad, urun, qty, tutar in c.fetchall():
            es = esleyici.coz(cad)
            if es.kesin and es.musteri_id != cid:
                rapor.append({
                    "sales_history.id": sid,
                    "eski": (cid, cad),
                    "onerilen": (es.musteri_id, es.musteri_ad),
                    "gerekce": es.gerekce,
                    "durum": es.durum,
                    "tutar": tutar,
                    "miktar": qty,
                    "urun": urun,
                })
            elif es.durum == "INCELE":
                rapor.append({
                    "sales_history.id": sid,
                    "eski": (cid, cad),
                    "onerilen": None,
                    "gerekce": es.gerekce,
                    "durum": "INCELE",
                    "adaylar": es.adaylar,
                    "tutar": tutar,
                })
        print(f"WP denetim (dry_run={dry_run}): {len(rapor)} satır incelenecek. Finansal düzeltme yapılmadı.")
        return rapor
    finally:
        conn.close()


if __name__ == "__main__":
    init_db()
    yeni_hammadde_gruplarini_tanimla()
    grupsuzlar_ve_kristal_mantigini_kesin_duzelt()
    tum_kopuk_satislari_gercek_musteriye_bagla()
    bosluksuz_ve_birlesik_isimleri_birlestir()
    fiil_ve_cop_carileri_temizle()
    kayip_hammadde_kartlarini_kurtar()
    app = QApplication(sys.argv)
    font = app.font()
    font.setPointSize(12)
    app.setFont(font)
    window = BenimPOSPlastik()
    window.show()
    sys.exit(app.exec())
