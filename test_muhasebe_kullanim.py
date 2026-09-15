"""Gerçek veriye dokunmadan CRUD ve ciro kontrolü.

Referans: Apple Q4 2024 yayımlanan ciro 94,9 milyar USD.
https://www.apple.com/newsroom/2024/10/apple-reports-fourth-quarter-results/
94.900 + 0 test girdisi milyon USD ölçeğindedir; TL dönüşümü değildir.
Gelir/gider senaryoları tamamen kurgusaldır, Apple kayıtları değildir.
"""
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from PyQt6.QtWidgets import QApplication, QMessageBox
from muhasebe import MuhasebeSayfasi, ozet, tutar_coz


class KullanimTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_ekle_duzenle_iptal(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder,'test.db')
            with sqlite3.connect(path) as conn:
                conn.execute('CREATE TABLE sales_history(qty,items_json,total_amount,payment_received)')
                conn.execute('CREATE TABLE customers(remaining_debt)')
            page = MuhasebeSayfasi(path)
            page.tur.setCurrentText('Gider')
            page.aciklama.setText('Kurgusal kira')
            page.tutar.setText('1.250,50')
            page.kaydet()
            page.table.selectRow(0); page.duzenle()
            page.tutar.setText('1.500,75'); page.kaydet()
            with sqlite3.connect(path) as conn:
                self.assertEqual(ozet(conn)['gider'],1500.75)
                self.assertEqual(conn.execute('SELECT COUNT(*) FROM on_muhasebe').fetchone()[0],1)
                self.assertEqual(conn.execute('SELECT COUNT(*) FROM on_muhasebe_gecmis').fetchone()[0],1)
            page.table.selectRow(0)
            with patch.object(QMessageBox,'question',return_value=QMessageBox.StandardButton.Yes):
                page.iptal()
            with sqlite3.connect(path) as conn:
                self.assertEqual(ozet(conn)['gider'],0)
                self.assertEqual(conn.execute('SELECT COUNT(*) FROM on_muhasebe').fetchone()[0],1)
            page.close()

    def test_kaynak_ciro_ve_tutar(self):
        from muhasebe import hazirla
        with sqlite3.connect(':memory:') as conn:
            hazirla(conn)
            conn.execute('CREATE TABLE sales_history(qty,items_json,total_amount,payment_received)')
            conn.execute('CREATE TABLE customers(remaining_debt)')
            conn.execute('INSERT INTO sales_history VALUES(1,NULL,94900,0)')
            self.assertEqual(ozet(conn)['ciro'],94900)
            conn.execute('UPDATE sales_history SET total_amount=95000')
            self.assertEqual(ozet(conn)['ciro'],95000)
        for value in ['12.50','1.2.3','1,234','-50']:
            with self.assertRaises(ValueError):
                tutar_coz(value)


if __name__ == '__main__':
    unittest.main()
