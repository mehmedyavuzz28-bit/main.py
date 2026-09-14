import sqlite3
import unittest
from muhasebe_donem import hazirla,rapor,stok_duzelt,tarih_coz

class DonemTest(unittest.TestCase):
    def test_donem_stok_maliyet(self):
        with sqlite3.connect(':memory:') as c:
            hazirla(c)
            c.execute('CREATE TABLE sales_history(created_at,qty,items_json,total_amount,payment_received)')
            c.execute('CREATE TABLE on_muhasebe(tarih,tur,kurus,iptal)')
            c.execute('CREATE TABLE products(id,name,stock_kg,buy_price,buy_currency,unit)')
            c.executemany('INSERT INTO sales_history VALUES(?,1,NULL,?,?)',
                [('2026-01-01 12:00:00',1000,500),('31.01.2026 12:00',2000,1000),('2026-02-01',9000,0)])
            c.execute("INSERT INTO products VALUES(1,'POM',10,20,'TRY','KG')")
            c.execute("INSERT INTO on_muhasebe VALUES('2026-01-02','Gider',10000,0)")
            d=rapor(c,'2026-01-01','2026-01-31')
            self.assertEqual(d['ciro'],3000); self.assertIsNone(d['kar'])
            c.execute("INSERT INTO muhasebe_donem_girdi VALUES('2026-01-01','2026-01-31',150000,'örnek')")
            self.assertEqual(rapor(c,'2026-01-01','2026-01-31')['kar'],1400)
            self.assertEqual(d['stok_degeri'],200)
            stok_duzelt(c,1,10,7,'sayım')
            self.assertEqual(rapor(c,'2026-01-01','2026-01-31')['stok_degeri'],140)
            with self.assertRaises(ValueError): stok_duzelt(c,1,10,8,'eski bakiye')
            self.assertEqual(c.execute('SELECT COUNT(*) FROM muhasebe_stok_duzeltme').fetchone()[0],1)
            self.assertIsNone(tarih_coz('bilinmeyen'))

if __name__=='__main__':unittest.main()
