import sqlite3
import unittest
from muhasebe import hazirla, hareket_ekle, ozet, tutar_coz


class MuhasebeTest(unittest.TestCase):
    def test_tutar(self):
        self.assertEqual(tutar_coz('1.250,50'),125050)
        for text in ['0','-5','NaN','Infinity','1,123','abc']:
            with self.assertRaises(ValueError):
                tutar_coz(text)

    def test_ozet_ve_iptal(self):
        with sqlite3.connect(':memory:') as conn:
            hazirla(conn)
            conn.execute('CREATE TABLE sales_history(qty,items_json,total_amount,payment_received)')
            conn.execute('CREATE TABLE customers(remaining_debt)')
            conn.executemany('INSERT INTO sales_history VALUES(?,?,?,?)',
                             [(2,None,100,50),(0,None,900,30),(0,'[{"qty":1}]',200,200)])
            conn.executemany('INSERT INTO customers VALUES(?)',[(120,),(-20,)])
            hareket_ekle(conn,'2026-09-14','Açılış','Kasa','İlk bakiye','1.000')
            hareket_ekle(conn,'2026-09-14','Gider','Kasa','Kira','100')
            hareket_ekle(conn,'2026-09-14','Gelir','Banka','Diğer gelir','50')
            data=ozet(conn)
            self.assertEqual(data['ciro'],300)
            self.assertEqual(data['tahsilat'],280)
            self.assertEqual(data['gelir'],50)
            self.assertEqual(data['kasa'],900)
            self.assertEqual(data['alacak'],120)
            self.assertEqual(data['avans'],20)
            conn.execute("UPDATE on_muhasebe SET iptal=1 WHERE tur='Gider'")
            self.assertEqual(ozet(conn)['kasa'],1000)
            self.assertEqual(ozet(conn)['gider'],0)
            self.assertEqual(conn.execute('SELECT COUNT(*) FROM on_muhasebe').fetchone()[0],3)


if __name__ == '__main__':
    unittest.main()
