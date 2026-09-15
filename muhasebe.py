"""Basit ön muhasebe; mevcut satış ve çek/senet kayıtlarını değiştirmez."""
import sqlite3
import re
import json
from contextlib import closing
from decimal import Decimal, InvalidOperation

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTabWidget, QFormLayout, QComboBox, QLineEdit, QDateEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QAbstractItemView)


def hazirla(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS on_muhasebe (
        id INTEGER PRIMARY KEY, tarih TEXT NOT NULL,
        tur TEXT NOT NULL CHECK(tur IN ('Gelir', 'Gider', 'Açılış')),
        hesap TEXT NOT NULL CHECK(hesap IN ('Kasa', 'Banka')),
        aciklama TEXT NOT NULL, kurus INTEGER NOT NULL CHECK(kurus > 0),
        iptal INTEGER NOT NULL DEFAULT 0 CHECK(iptal IN (0,1)))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS on_muhasebe_gecmis (
        id INTEGER PRIMARY KEY, hareket_id INTEGER NOT NULL,
        eski_kayit TEXT NOT NULL, zaman TEXT DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()


def tutar_coz(text):
    try:
        if not re.fullmatch(r'(?:[0-9]+|[0-9]{1,3}(?:\.[0-9]{3})+)(?:,[0-9]{1,2})?', text.strip()):
            raise ValueError
        value = Decimal(text.strip().replace(' ', '').replace('.', '').replace(',', '.'))
        if not value.is_finite() or value <= 0 or value > Decimal('9999999999.99'):
            raise ValueError
        if value != value.quantize(Decimal('0.01')):
            raise ValueError
        return int(value * 100)
    except (InvalidOperation, ValueError):
        raise ValueError('Pozitif bir TL tutarı yazın. Örnek: 1.250,50') from None


def hareket_ekle(conn, tarih, tur, hesap, aciklama, tutar, kayit_id=None):
    kurus = tutar_coz(tutar)
    if not aciklama.strip():
        raise ValueError('Açıklama boş bırakılamaz.')
    if not QDate.fromString(tarih, 'yyyy-MM-dd').isValid():
        raise ValueError('Geçerli bir tarih seçin.')
    with conn:
        values = (tarih, tur, hesap, aciklama.strip(), kurus)
        if kayit_id is None:
            conn.execute('INSERT INTO on_muhasebe(tarih,tur,hesap,aciklama,kurus) VALUES(?,?,?,?,?)', values)
        else:
            old = conn.execute('SELECT * FROM on_muhasebe WHERE id=? AND iptal=0', (kayit_id,)).fetchone()
            if old is None:
                raise ValueError('Kayıt bulunamadı veya iptal edilmiş.')
            conn.execute('INSERT INTO on_muhasebe_gecmis(hareket_id,eski_kayit) VALUES(?,?)',
                         (kayit_id,json.dumps(old,ensure_ascii=False)))
            conn.execute('UPDATE on_muhasebe SET tarih=?,tur=?,hesap=?,aciklama=?,kurus=? WHERE id=?',values+(kayit_id,))


def ozet(conn):
    # Satışlar ve saf tahsilatlar ayrıdır; ciro ödeme toplamı değildir.
    ciro, tahsilat = conn.execute('''SELECT
        COALESCE(SUM(CASE WHEN COALESCE(qty,0)>0 OR
            (items_json IS NOT NULL AND items_json NOT IN ('','[]'))
            THEN total_amount ELSE 0 END),0),
        COALESCE(SUM(payment_received),0) FROM sales_history''').fetchone()
    alacak, cari_avans = conn.execute('''SELECT
        COALESCE(SUM(MAX(COALESCE(remaining_debt,0),0)),0),
        COALESCE(SUM(MAX(-COALESCE(remaining_debt,0),0)),0) FROM customers''').fetchone()
    gelir = gider = 0
    hesaplar = {'Kasa': 0, 'Banka': 0}
    for tur, hesap, kurus in conn.execute('SELECT tur,hesap,kurus FROM on_muhasebe WHERE iptal=0'):
        hesaplar[hesap] += -kurus if tur == 'Gider' else kurus
        gelir += kurus if tur == 'Gelir' else 0
        gider += kurus if tur == 'Gider' else 0
    return dict(ciro=ciro, tahsilat=tahsilat, alacak=alacak, avans=cari_avans,
                gelir=gelir/100, gider=gider/100,
                kasa=hesaplar['Kasa']/100, banka=hesaplar['Banka']/100)


def tl(value):
    return f'{value:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.') + ' TL'


class MuhasebeSayfasi(QWidget):
    def __init__(self, db_path, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.edit_id = None
        with closing(sqlite3.connect(db_path)) as conn:
            hazirla(conn)
        self.setStyleSheet('''
            QWidget { background: #101a29; color: #edf3fc; }
            QTabWidget::pane, QLineEdit, QComboBox, QDateEdit {
                border: 1px solid #364b65; border-radius: 8px; padding: 6px; }
            QTabBar::tab { padding: 10px 16px; background: #192b42; }
            QTabBar::tab:selected { background: #245287; }
            QPushButton { background: #245287; padding: 8px 14px; border-radius: 8px; }
            QPushButton:hover { background: #3069a8; }
            QLabel#kart { background: #192b42; padding: 12px; border-radius: 10px; }
        ''')
        root = QVBoxLayout(self)
        title = QLabel('MUHASEBE — İşletmenizin genel durumu')
        root.addWidget(title)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs)
        self.kartlar = {}
        summary = QWidget()
        layout = QVBoxLayout(summary)
        for key, title in [('ciro','Ciro · kayıtlı satış toplamı'),
                ('tahsilat','Satış geçmişindeki tahsilat'), ('alacak','Müşterilerden alacak'),
                ('gelir','Diğer gelirler · satış hariç'), ('gider','Kaydedilen giderler')]:
            label = QLabel()
            label.setObjectName('kart')
            label.setWordWrap(True)
            layout.addWidget(label)
            self.kartlar[key] = (label, title)
        note = QLabel('Tüm zamanlar • Yalnızca mevcut kayıtlara dayanır.\n'
            'Ciro kâr değildir. Satılan ürünlerin tarihsel maliyeti eksik olduğundan net kâr hesaplanmaz.\n'
            'Eski WhatsApp tutarları ve döviz kayıtları doğrulanmadan sonuçları kesin kabul etmeyin.')
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addStretch()
        self.tabs.addTab(summary, 'Genel Durum')
        movements = QWidget()
        ml = QVBoxLayout(movements)
        form = QFormLayout()
        self.tarih = QDateEdit(QDate.currentDate())
        self.tarih.setCalendarPopup(True)
        self.tarih.setDisplayFormat('dd.MM.yyyy')
        self.tur = QComboBox(); self.tur.addItems(['Gider','Gelir','Açılış'])
        self.hesap = QComboBox(); self.hesap.addItems(['Kasa','Banka'])
        self.aciklama = QLineEdit()
        self.tutar = QLineEdit(); self.tutar.setPlaceholderText('Örnek: 1.250,50')
        for name, widget in [('Tarih',self.tarih),('İşlem',self.tur),('Hesap',self.hesap),
                             ('Açıklama',self.aciklama),('Tutar (TL)',self.tutar)]:
            form.addRow(name,widget)
        ml.addLayout(form)
        warning = QLabel('Satış tahsilatlarını burada tekrar gelir yazmayın. Açılış: ilk kasa/banka tutarı; gelir sayılmaz.')
        warning.setWordWrap(True); ml.addWidget(warning)
        actions = QHBoxLayout()
        self.save = QPushButton('Kaydı ekle'); self.save.clicked.connect(self.kaydet)
        edit = QPushButton('Seçileni düzenle'); edit.clicked.connect(self.duzenle)
        reset = QPushButton('Yeni kayıt / Vazgeç'); reset.clicked.connect(self.form_temizle)
        cancel = QPushButton('Seçili kaydı iptal et'); cancel.clicked.connect(self.iptal)
        actions.addWidget(self.save); actions.addWidget(edit); actions.addWidget(reset)
        actions.addWidget(cancel); actions.addStretch()
        ml.addLayout(actions)
        self.table = QTableWidget(0,7)
        self.table.setHorizontalHeaderLabels(['No','Tarih','İşlem','Hesap','Açıklama','Tutar','Durum'])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4,QHeaderView.ResizeMode.Stretch)
        ml.addWidget(self.table)
        self.tabs.addTab(movements,'Gelir–Gider / Kasa–Banka')
        balance = QWidget(); bl = QVBoxLayout(balance)
        self.bilanco = QLabel(); self.bilanco.setWordWrap(True); bl.addWidget(self.bilanco); bl.addStretch()
        self.tabs.addTab(balance,'Varlık–Borç Özeti')
        from muhasebe_donem import DonemSayfasi
        self.tabs.addTab(DonemSayfasi(db_path),'Dönem / Stok / Maliyet')
        refresh = QPushButton('Bilgileri yenile'); refresh.clicked.connect(self.yenile); root.addWidget(refresh)

    def showEvent(self, event):
        super().showEvent(event)
        self.yenile()

    def yenile(self):
        with closing(sqlite3.connect(self.db_path)) as conn:
            data = ozet(conn)
            rows = conn.execute('SELECT id,tarih,tur,hesap,aciklama,kurus,iptal FROM on_muhasebe ORDER BY tarih DESC,id DESC').fetchall()
        for key,(label,title) in self.kartlar.items():
            label.setText(title + '\n' + tl(data[key]))
        self.bilanco.setText('VARLIK–BORÇ ÖZETİ (resmî bilanço değildir)\n\n'
            f"Kasa · yalnızca bu bölümdeki kayıtlar: {tl(data['kasa'])}\n"
            f"Banka · yalnızca bu bölümdeki kayıtlar: {tl(data['banka'])}\n"
            f"Cari alacak: {tl(data['alacak'])}\n"
            f"Negatif cari bakiyeler / müşteri avansları: {tl(data['avans'])}\n\n"
            'Satış tahsilatları hesap dağılımı bilinmediği için kasa/bankaya otomatik eklenmez.\n'
            'Tedarikçi borçları, krediler, stok maliyeti ve sabit varlıklar henüz bu özete dahil değildir.\n'
            'Bu nedenle toplam varlık, toplam borç ve özkaynak hesaplanmaz. Çek–Senet mevcut sayfasındadır.')
        self.table.setRowCount(len(rows))
        for i,row in enumerate(rows):
            values = list(row[:5]) + [tl(row[5]/100),'İptal' if row[6] else 'Aktif']
            for j,value in enumerate(values):
                self.table.setItem(i,j,QTableWidgetItem(str(value)))

    def kaydet(self):
        try:
            with closing(sqlite3.connect(self.db_path)) as conn:
                hareket_ekle(conn,self.tarih.date().toString('yyyy-MM-dd'),self.tur.currentText(),
                            self.hesap.currentText(),self.aciklama.text(),self.tutar.text(),self.edit_id)
        except (ValueError,sqlite3.Error) as exc:
            QMessageBox.warning(self,'Kayıt eklenemedi',str(exc)); return
        self.form_temizle(); self.yenile()

    def form_temizle(self):
        self.edit_id = None
        self.save.setText('Kaydı ekle')
        self.aciklama.clear(); self.tutar.clear()

    def duzenle(self):
        row = self.table.currentRow()
        if row < 0 or self.table.item(row,6).text() == 'İptal':
            QMessageBox.information(self,'Kayıt seçin','Önce tablodan aktif bir kayıt seçin.'); return
        self.edit_id = int(self.table.item(row,0).text())
        self.tarih.setDate(QDate.fromString(self.table.item(row,1).text(),'yyyy-MM-dd'))
        self.tur.setCurrentText(self.table.item(row,2).text())
        self.hesap.setCurrentText(self.table.item(row,3).text())
        self.aciklama.setText(self.table.item(row,4).text())
        self.tutar.setText(self.table.item(row,5).text().removesuffix(' TL'))
        self.save.setText('Değişikliği kaydet')

    def iptal(self):
        row = self.table.currentRow()
        if row < 0:
            return
        if QMessageBox.question(self,'Kaydı iptal et','Seçili hareket iptal edilsin mi? Kayıt silinmez.',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        try:
            with closing(sqlite3.connect(self.db_path)) as conn, conn:
                conn.execute('UPDATE on_muhasebe SET iptal=1 WHERE id=?',(int(self.table.item(row,0).text()),))
        except sqlite3.Error as exc:
            QMessageBox.warning(self,'İptal edilemedi',str(exc)); return
        self.form_temizle()
        self.yenile()
