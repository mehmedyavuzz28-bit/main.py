"""Dönem raporu: geçmiş satışlarla güncel stok bakiyesini karıştırmaz."""
import sqlite3
from contextlib import closing
from datetime import datetime
from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (QWidget,QVBoxLayout,QHBoxLayout,QFormLayout,QLabel,
    QPushButton,QDateEdit,QLineEdit,QComboBox,QMessageBox,QDoubleSpinBox)
from muhasebe import tl, tutar_coz


def tarih_coz(value):
    value = str(value or '').strip()
    for fmt in ('%Y-%m-%d','%d.%m.%Y','%d/%m/%Y','%d-%m-%Y'):
        try:
            return datetime.strptime(value[:10],fmt).date().isoformat()
        except ValueError:
            pass
    return None


def hazirla(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS muhasebe_ek_bakiye (
        tarih TEXT PRIMARY KEY, varlik INTEGER, tedarikci INTEGER, kredi INTEGER,
        aciklama TEXT NOT NULL)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS muhasebe_donem_girdi (
        bas TEXT, son TEXT, maliyet INTEGER NOT NULL, aciklama TEXT NOT NULL,
        PRIMARY KEY(bas,son))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS muhasebe_stok_duzeltme (
        id INTEGER PRIMARY KEY, urun_id INTEGER, eski REAL, yeni REAL,
        neden TEXT NOT NULL, zaman TEXT DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()


def rapor(conn, bas, son):
    if bas > son:
        raise ValueError('Başlangıç tarihi bitişten sonra olamaz.')
    ciro = tahsilat = gelir = gider = 0
    bilinmeyen = 0
    for tarih,qty,items,total,paid in conn.execute(
            'SELECT created_at,qty,items_json,total_amount,payment_received FROM sales_history'):
        date = tarih_coz(tarih)
        if date is None:
            bilinmeyen += 1
            continue
        if bas <= date <= son:
            if (qty or 0)>0 or (items and items not in ('','[]')):
                ciro += total or 0
            tahsilat += paid or 0
    for tur,kurus in conn.execute('SELECT tur,kurus FROM on_muhasebe WHERE iptal=0 AND tarih BETWEEN ? AND ?', (bas,son)):
        gelir += kurus/100 if tur=='Gelir' else 0
        gider += kurus/100 if tur=='Gider' else 0
    manual = conn.execute('SELECT maliyet FROM muhasebe_donem_girdi WHERE bas=? AND son=?',(bas,son)).fetchone()
    # Bugünkü alış fiyatı geçmiş satışın maliyeti değildir.
    maliyet = manual[0]/100 if manual else None
    stok_degeri = 0
    eksik_stok = 0
    for qty,price,currency in conn.execute('SELECT stock_kg,buy_price,buy_currency FROM products'):
        if not qty:
            continue
        if qty<0 or not price or price<0 or currency not in ('TRY','TL'):
            eksik_stok += 1
        else:
            stok_degeri += qty*price
    return dict(ciro=ciro,tahsilat=tahsilat,gelir=gelir,gider=gider,maliyet=maliyet,
        kar=None if maliyet is None else ciro-maliyet+gelir-gider,
        stok_degeri=stok_degeri,eksik_stok=eksik_stok,bilinmeyen=bilinmeyen)


def stok_duzelt(conn, pid, eski, yeni, neden):
    if not neden.strip():
        raise ValueError('Stok değişikliği için açıklama yazın.')
    with conn:
        cur = conn.execute('UPDATE products SET stock_kg=? WHERE id=? AND stock_kg=?',(yeni,pid,eski))
        if cur.rowcount != 1:
            raise ValueError('Stok başka işlemde değişti. Listeyi yenileyip tekrar deneyin.')
        conn.execute('INSERT INTO muhasebe_stok_duzeltme(urun_id,eski,yeni,neden) VALUES(?,?,?,?)',
                     (pid,eski,yeni,neden.strip()))


class DonemSayfasi(QWidget):
    def __init__(self,path,parent=None):
        super().__init__(parent)
        self.path=path
        with closing(sqlite3.connect(path)) as conn:
            hazirla(conn)
        layout=QVBoxLayout(self)
        dates=QHBoxLayout()
        today=QDate.currentDate()
        self.bas=QDateEdit(QDate(today.year(),today.month(),1))
        self.son=QDateEdit(today)
        for label,widget in [('Başlangıç',self.bas),('Bitiş',self.son)]:
            widget.setCalendarPopup(True); widget.setDisplayFormat('dd.MM.yyyy')
            dates.addWidget(QLabel(label)); dates.addWidget(widget)
        for name,fn in [('Bu ay',self.ay),('Bu yıl',self.yil),('Raporla',self.yenile)]:
            button=QPushButton(name); button.clicked.connect(fn); dates.addWidget(button)
        layout.addLayout(dates)
        self.result=QLabel(); self.result.setWordWrap(True); layout.addWidget(self.result)
        form=QFormLayout()
        self.cost=QLineEdit(); self.cost.setPlaceholderText('Seçilen dönemde SATILAN malların toplam maliyeti (TL)')
        self.note=QLineEdit(); self.note.setPlaceholderText('Örnek: muhasebecinin Haziran maliyet hesabı')
        form.addRow('Satılan mal maliyeti',self.cost); form.addRow('Kaynak / açıklama',self.note)
        layout.addLayout(form)
        save=QPushButton('Seçilen dönemin maliyetini kaydet / güncelle'); save.clicked.connect(self.kaydet)
        layout.addWidget(save)
        extra=QFormLayout()
        self.bakiyeler={}
        for key,title in [('varlik','Diğer varlıklar (net TL)'),('tedarikci','Tedarikçi borcu (TL)'),('kredi','Kalan kredi borcu (TL)')]:
            field=QLineEdit(); field.setPlaceholderText('Bitiş tarihi itibarıyla bakiye; bilinmiyorsa boş bırakın')
            self.bakiyeler[key]=field; extra.addRow(title,field)
        self.bakiye_not=QLineEdit(); extra.addRow('Bakiye açıklaması',self.bakiye_not)
        layout.addLayout(extra)
        button=QPushButton('Bitiş tarihi bakiyelerini kaydet'); button.clicked.connect(self.bakiye_kaydet); layout.addWidget(button)
        notice=QLabel('Bu alan tüm mal alışları değil, yalnızca satılan malın maliyetidir.\n'
            'Maliyet tam seçilen tarih aralığına aittir; farklı aralıkta otomatik birleştirilmez.\n'
            'Sonuç tahmini faaliyet sonucudur; vergi, faiz ve amortisman dahil edilmediyse net kâr değildir.\n'
            'Stok aşağıda GÜNCELDİR; seçilen geçmiş tarihin stoğu değildir.')
        notice.setWordWrap(True); layout.addWidget(notice)
        stock=QFormLayout()
        self.product=QComboBox(); self.qty=QDoubleSpinBox()
        self.qty.setRange(-1e9,1e9); self.qty.setDecimals(3)
        self.reason=QLineEdit(); self.reason.setPlaceholderText('Örnek: depo sayımı, tartım farkı')
        stock.addRow('Mevcut ürün',self.product); stock.addRow('Yeni stok miktarı',self.qty)
        stock.addRow('Değişiklik nedeni',self.reason); layout.addLayout(stock)
        self.product.currentIndexChanged.connect(self.sec)
        button=QPushButton('Stok miktarını düzelt'); button.clicked.connect(self.stok_kaydet); layout.addWidget(button)
        layout.addStretch()
        self.bas.dateChanged.connect(self.temizle); self.son.dateChanged.connect(self.temizle)

    def aralik(self):
        return self.bas.date().toString('yyyy-MM-dd'),self.son.date().toString('yyyy-MM-dd')

    def temizle(self):
        self.cost.clear(); self.note.clear(); self.result.setText('Tarih değişti. Raporla düğmesine basın.')
        for field in self.bakiyeler.values():field.clear()
        self.bakiye_not.clear()

    def ay(self):
        d=self.bas.date(); self.bas.setDate(QDate(d.year(),d.month(),1))
        self.son.setDate(QDate(d.year(),d.month(),d.daysInMonth())); self.yenile()

    def yil(self):
        y=self.bas.date().year(); self.bas.setDate(QDate(y,1,1)); self.son.setDate(QDate(y,12,31)); self.yenile()

    def showEvent(self,event):
        super().showEvent(event); self.yenile()

    def yenile(self):
        try:
            with closing(sqlite3.connect(self.path)) as conn:
                d=rapor(conn,*self.aralik())
                row=conn.execute('SELECT maliyet,aciklama FROM muhasebe_donem_girdi WHERE bas=? AND son=?',self.aralik()).fetchone()
                products=conn.execute('SELECT id,name,stock_kg,unit FROM products ORDER BY name').fetchall()
                bakiye=conn.execute('SELECT varlik,tedarikci,kredi,aciklama FROM muhasebe_ek_bakiye WHERE tarih=?',(self.aralik()[1],)).fetchone()
            for i,field in enumerate(self.bakiyeler.values()):
                field.setText(tl(bakiye[i]/100).removesuffix(' TL') if bakiye and bakiye[i] is not None else '')
            self.bakiye_not.setText(bakiye[3] if bakiye else '')
            self.cost.setText(tl(row[0]/100).removesuffix(' TL') if row else '')
            self.note.setText(row[1] if row else '')
            self.result.setText(f"Dönem cirosu: {tl(d['ciro'])} | Tahsilat: {tl(d['tahsilat'])}\n"
                f"Diğer gelir: {tl(d['gelir'])} | Gider: {tl(d['gider'])}\n"
                + ('Tahmini faaliyet sonucu: '+tl(d['kar']) if d['kar'] is not None else 'Kâr hesaplanmadı: dönem maliyetini girin.')
                + f"\nGüncel stok değeri (kayıtlı TL alış fiyatıyla): {tl(d['stok_degeri'])}\n"
                f"Değere dahil edilemeyen ürün: {d['eksik_stok']} | Tarihi okunamayan satış: {d['bilinmeyen']}")
            self.product.clear()
            for pid,name,qty,unit in products:
                self.product.addItem(f'{name} ({unit}) — mevcut: {qty}',(pid,qty or 0))
            self.sec()
        except (sqlite3.Error,ValueError) as exc:
            QMessageBox.warning(self,'Rapor alınamadı',str(exc))

    def sec(self):
        data=self.product.currentData()
        if data: self.qty.setValue(data[1])

    def kaydet(self):
        try:
            bas,son=self.aralik()
            if bas>son or not self.note.text().strip(): raise ValueError('Tarih aralığını kontrol edin ve açıklama yazın.')
            cost=0 if self.cost.text().strip() in ('0','0,00') else tutar_coz(self.cost.text())
            with closing(sqlite3.connect(self.path)) as conn,conn:
                conn.execute('INSERT INTO muhasebe_donem_girdi VALUES(?,?,?,?) ON CONFLICT(bas,son) DO UPDATE SET maliyet=excluded.maliyet,aciklama=excluded.aciklama',
                             (bas,son,cost,self.note.text().strip()))
            self.yenile()
        except (ValueError,sqlite3.Error) as exc: QMessageBox.warning(self,'Kaydedilemedi',str(exc))

    def stok_kaydet(self):
        data=self.product.currentData()
        if not data:return
        if QMessageBox.question(self,'Stok düzeltmesi',f'Mevcut {data[1]} → yeni {self.qty.value()}. Onaylıyor musunuz?',
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        try:
            with closing(sqlite3.connect(self.path)) as conn:
                stok_duzelt(conn,*data,self.qty.value(),self.reason.text())
            self.reason.clear(); self.yenile()
        except (ValueError,sqlite3.Error) as exc: QMessageBox.warning(self,'Kaydedilemedi',str(exc))

    def bakiye_kaydet(self):
        try:
            if not self.bakiye_not.text().strip():raise ValueError('Bakiyelerin kaynağını açıklamaya yazın.')
            values=[]
            for field in self.bakiyeler.values():
                text=field.text().strip()
                values.append(None if not text else 0 if text in ('0','0,00') else tutar_coz(text))
            with closing(sqlite3.connect(self.path)) as conn,conn:
                conn.execute('INSERT INTO muhasebe_ek_bakiye VALUES(?,?,?,?,?) ON CONFLICT(tarih) DO UPDATE SET varlik=excluded.varlik,tedarikci=excluded.tedarikci,kredi=excluded.kredi,aciklama=excluded.aciklama',
                    (self.aralik()[1],*values,self.bakiye_not.text().strip()))
            self.yenile()
        except (ValueError,sqlite3.Error) as exc:QMessageBox.warning(self,'Kaydedilemedi',str(exc))
