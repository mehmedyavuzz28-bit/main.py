import sqlite3
import tempfile
import os
import unittest

from wp_musteri_esleme import MusteriEsleyici, MUSTERISIZ, wp_ek_borc


class MusteriEslemeKabulTest(unittest.TestCase):
    def setUp(self):
        self.kayitlar = [
            (1, "ALİ USTA", 0),
            (2, "EFOR PLASTİK RAMAZAN", 0),
            (3, "OLGUN", 0),
            (4, "HAS RULO", 0),
            (5, "ÇOLAK PLASTİK", 0),
            (6, "UĞUR PLASTİK", 0),
            (7, "ALAATTİN AKSESUAR", 0),
            (8, "BURHAN", 0),
            (9, "AKRİLİK BURHAN", 0),
            (10, "EFE PLASTİK", 0),
            (11, "ADANA CAN PLASTİK", 0),
            (12, "ALAATTİN MUSTAFA", 0),
            (13, "TAHSİN", 0),
            (14, "ERSİN", 0),
        ]
        self.es = MusteriEsleyici(self.kayitlar)

    def test_ali_usta_kisaltilmaz(self):
        e = self.es.coz("ali usta")
        self.assertTrue(e.kesin)
        self.assertEqual(e.musteri_id, 1)
        self.assertIn("USTA", e.musteri_ad)

    def test_alaattin_aliaslar(self):
        for ham in ["Alaattin", "Alattin", "Alaaddin", "Aladdin", "Alddinn"]:
            e = self.es.coz(ham)
            self.assertTrue(e.kesin, ham)
            self.assertEqual(e.musteri_id, 7, ham)

    def test_burhan_akrilik_ayni_kisi(self):
        e = self.es.coz("akrilik burhan")
        self.assertTrue(e.kesin)
        self.assertEqual(e.musteri_id, 8)
        self.assertIn("BURHAN", e.musteri_ad)

    def test_plastik_ortak_kelime_ayirir(self):
        e1 = self.es.coz("EFE PLASTİK")
        e2 = self.es.coz("ADANA CAN PLASTİK")
        self.assertEqual(e1.musteri_id, 10)
        self.assertEqual(e2.musteri_id, 11)
        ters = MusteriEsleyici(list(reversed(self.kayitlar)))
        self.assertEqual(ters.coz("EFE PLASTİK").musteri_id, 10)
        self.assertEqual(ters.coz("ADANA CAN PLASTİK").musteri_id, 11)

    def test_mustafa_alaattin_diger_hesaba_gitmez(self):
        e = self.es.coz("MUSTAFA ALAATTİN")
        self.assertTrue(e.kesin)
        self.assertEqual(e.musteri_id, 12)

    def test_tahsin_ersin_fiil_degil(self):
        self.assertEqual(self.es.coz("TAHSİN").musteri_id, 13)
        self.assertEqual(self.es.coz("ERSİN").musteri_id, 14)

    def test_musteri_kelimesi_musterisiz(self):
        e = self.es.coz("müşteri")
        self.assertFalse(e.kesin)
        self.assertEqual(e.musteri_ad, MUSTERISIZ)

    def test_devir_borc_ornek_19(self):
        d, p = 1000.0, 900.0
        p += 200.0
        ek1 = wp_ek_borc(d, 900.0, 200.0)
        ek2 = wp_ek_borc(d, 1100.0, 100.0)
        self.assertEqual(ek1, 100.0)
        self.assertEqual(ek2, 100.0)
        self.assertEqual(ek1 + ek2, 200.0)


class KayitYoluTekrarTest(unittest.TestCase):
    def test_ayni_kaynak_ikinci_kez_yazilmaz(self):
        from main import wp_satis_kayit_servisi, wp_kaynak_anahtari_uret

        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            conn = sqlite3.connect(path)
            c = conn.cursor()
            c.execute("CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT, shopping_count INTEGER DEFAULT 0, debt REAL DEFAULT 0, payment REAL DEFAULT 0, remaining_debt REAL DEFAULT 0)")
            c.execute("CREATE TABLE product_groups (id INTEGER PRIMARY KEY, name TEXT)")
            c.execute("INSERT INTO product_groups (name) VALUES ('POM')")
            c.execute("CREATE TABLE products (id INTEGER PRIMARY KEY, group_id INTEGER, name TEXT, stock_kg REAL DEFAULT 0, buy_price REAL DEFAULT 0, sell_price REAL DEFAULT 0)")
            c.execute("""CREATE TABLE sales_history (
                id INTEGER PRIMARY KEY, personnel_name TEXT, customer_id INTEGER, customer_name TEXT,
                product_name TEXT, qty REAL, price REAL, total_amount REAL, payment_type TEXT,
                payment_received REAL, created_at TEXT, note TEXT)""")
            c.execute("""CREATE TABLE wp_kaynak_olaylar (
                id INTEGER PRIMARY KEY, kaynak_anahtar TEXT UNIQUE, tarih TEXT, personel TEXT,
                ham_govde TEXT, ham_musteri_aday TEXT, belirsizlik TEXT, kalem_no INTEGER,
                sales_history_id INTEGER, tutar REAL, tahsilat REAL, durum TEXT)""")
            c.execute("INSERT INTO customers (name, remaining_debt) VALUES ('OLGUN', 0)")
            conn.commit()

            kayit = [{
                "musteri_id": 1, "musteri_ad": "OLGUN", "urun_ad": "POM",
                "miktar": 126, "fiyat": 35, "odeme_alindi": 4410,
                "tarih": "05.04.2023 18:09", "personel": "BABA",
                "ham_govde": "126 kğ granür derlin 35 tl ödemesi alındı Olgun",
                "kalem_no": 0, "durum": "TAMAM",
            }]

            def gid(cur, u):
                return 1

            o1 = wp_satis_kayit_servisi(c, conn, kayit, {}, {}, gid)
            o2 = wp_satis_kayit_servisi(c, conn, kayit, {}, {}, gid)
            c.execute("SELECT COUNT(*) FROM sales_history")
            n = c.fetchone()[0]
            c.execute("SELECT remaining_debt, payment, shopping_count FROM customers WHERE id=1")
            rem, pay, shop = c.fetchone()
            self.assertEqual(o1["kaydedilen"], 1)
            self.assertEqual(o2["mevcut"], 1)
            self.assertEqual(n, 1)
            self.assertEqual(shop, 1)
            self.assertEqual(pay, 4410)
            conn.close()
        finally:
            os.remove(path)

    def test_kaynak_anahtari_parser_surumune_bagli_degil(self):
        from main import wp_kaynak_anahtari_uret
        a = wp_kaynak_anahtari_uret("1", "A", "govde", 0)
        b = wp_kaynak_anahtari_uret("1", "A", "govde", 0)
        c = wp_kaynak_anahtari_uret("1", "A", "govde", 1)
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)


class ParserKabulTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from main import WhatsAppSatisAyristirici
        cls.A = WhatsAppSatisAyristirici
        cls.musteriler = [
            (1, "ALİ USTA"),
            (2, "EFOR PLASTİK RAMAZAN"),
            (3, "OLGUN"),
            (4, "HAS RULO"),
            (5, "ÇOLAK PLASTİK"),
            (6, "UĞUR PLASTİK"),
        ]
        cls.urunler = [(1, "POM", 40.0), (2, "PP MOBLEN", 30.0), (3, "ABS", 50.0)]

    def test_ali_usta_311(self):
        s = self.A.bloklari_ayristir("311 kg pp ali usta", self.musteriler, self.urunler)
        self.assertTrue(s)
        self.assertEqual(s[0]["miktar"], 311)
        self.assertIn("USTA", s[0]["musteri_ad"].upper())
        self.assertNotEqual(s[0]["musteri_ad"].upper(), "ALI")

    def test_sonraki_mesajda_musteri(self):
        ham = (
            "[28.03.2023 16:37:32] yavuz: 444 kg tıruncu delrim\n"
            "[28.03.2023 16:37:51] yavuz: efor plastik ramazan abi\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0]["miktar"], 444)
        self.assertIn("RAMAZAN", s[0]["musteri_ad"].upper())
        self.assertNotIn("DELRIM", s[0]["musteri_ad"].upper())

    def test_olgun_tahsilat(self):
        ham = (
            "[5.04.2023 18:09:14] Baba1: 126 kğ granür derlin 35 tl ödemesi alındı\n"
            "[5.04.2023 18:09:40] Baba1: Olgun\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0]["miktar"], 126)
        self.assertEqual(s[0]["fiyat"], 35)
        self.assertEqual(s[0]["tutar"], 4410)
        self.assertEqual(s[0]["odeme_alindi"], 4410)
        self.assertIn("OLGUN", s[0]["musteri_ad"].upper())

    def test_iki_musteri_iki_kalem(self):
        ham = "ali usta 340 kg moblen\n497.5 kg abs orjinal uğur plastik"
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertEqual(len(s), 2)
        adlar = {x["musteri_ad"].upper() for x in s}
        self.assertTrue(any("USTA" in a for a in adlar))
        self.assertTrue(any("UĞUR" in a or "UGUR" in a for a in adlar))

    def test_colak_yazarsin_82_kg(self):
        ham = (
            "[7.04.2023 14:27:59] yavuz: 56 kg çolak plastik moblen\n"
            "[7.04.2023 14:28:04] yavuz: 26 da sabah vardı\n"
            "[7.04.2023 14:28:13] yavuz: 82 kg yazarsın adama baba\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        toplam = sum(x["miktar"] for x in s)
        self.assertEqual(toplam, 82)

    def test_net_satis_rapor_degil(self):
        s = self.A.bloklari_ayristir("nisanın ilk haftası 352.322 tl net satış", self.musteriler, self.urunler)
        self.assertEqual(s, [])

    def test_iki_urun_tek_tahsilat(self):
        ham = "100 kg pom ali usta 50 tl\n80 kg abs uğur plastik 40 tl 10000 tl odeme"
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertGreaterEqual(len(s), 1)
        self.assertEqual(sum(x.get("odeme_alindi") or 0 for x in s), 10000)

    def test_has_rulo_doviz_fiyat_degil(self):
        ham = (
            "[12.04.2023 21:48:21] yavuz: 125 kg granür delrin verdim\n"
            "[12.04.2023 21:48:43] yavuz: has rulaya\n"
            "[12.04.2023 21:48:52] yavuz: 200 dolar 50 euro aldım\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0]["miktar"], 125)
        self.assertEqual(s[0].get("odeme_usd"), 200)
        self.assertEqual(s[0].get("odeme_eur"), 50)
        self.assertNotAlmostEqual(s[0]["fiyat"], 200 * 38.5)

    def test_colak_sadece_tahsilat(self):
        s = self.A.bloklari_ayristir("Çolak plastikten 12.500 TL ödeme geldi", self.musteriler, self.urunler)
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0].get("olay_tipi"), "TAHSILAT")
        self.assertEqual(s[0]["miktar"], 0)
        self.assertEqual(s[0]["odeme_alindi"], 12500)
        self.assertEqual(s[0]["musteri_id"], 5)
        self.assertIn("COLAK", s[0]["musteri_ad"].upper().replace("Ç", "C").replace("İ", "I"))

    def test_bin_celiskisi_incele(self):
        s = self.A.bloklari_ayristir("müslüm 7000 bin ödeme verdi", self.musteriler, self.urunler)
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0]["durum"], "INCELE")
        self.assertNotEqual(s[0]["odeme_alindi"], 7000000)
        self.assertEqual(s[0]["odeme_alindi"], 0)

    def test_iki_ton_sekiz_yuz_tek_satis(self):
        ham = (
            "İki ton 800 kg Orjinal abiyez uğur plastik\n"
            "bir ton 800 kg ödeme alındı\n"
            "bir ton ödeme cuma günü haftaya\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        satis = [x for x in s if x.get("olay_tipi") != "TAHSILAT"]
        self.assertEqual(len(satis), 1)
        self.assertEqual(satis[0]["miktar"], 2800)
        self.assertEqual(satis[0]["odeme_alindi"], 0)
        ad = satis[0]["musteri_ad"].upper().replace("İ", "I")
        self.assertIn("UGUR", ad)

    def test_ayni_kilo_farkli_saat_mukerrer_aday(self):
        ham = (
            "[7.04.2023 09:00:00] yavuz: 497.5 kg abs orjinal uğur plastik\n"
            "[7.04.2023 18:00:00] yavuz: 497.5 kg abs orjinal uğur plastik\n"
        )
        s = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        self.assertEqual(len(s), 2)
        self.assertTrue(all(x.get("olasi_mukerrer") for x in s))
        self.assertEqual(sum(x["miktar"] for x in s), 995.0)

    def test_kaynak_kimligi_tek_metin_ayni(self):
        ham = (
            "[28.03.2023 16:37:32] yavuz: 444 kg tıruncu delrim\n"
            "[28.03.2023 16:37:51] yavuz: efor plastik ramazan abi\n"
        )
        a = self.A.bloklari_ayristir(ham, self.musteriler, self.urunler)
        b = self.A.bloklari_ayristir(ham.replace("\n", "\n"), self.musteriler, self.urunler)
        self.assertEqual(len(a), len(b))
        self.assertEqual(a[0]["miktar"], b[0]["miktar"])
        self.assertEqual(a[0]["ham_govde"], b[0]["ham_govde"])


class KayitRollbackTest(unittest.TestCase):
    def test_hata_olunca_yarim_kalmaz(self):
        from main import wp_satis_kayit_servisi
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            conn = sqlite3.connect(path)
            c = conn.cursor()
            c.execute("CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT, shopping_count INTEGER DEFAULT 0, debt REAL DEFAULT 0, payment REAL DEFAULT 0, remaining_debt REAL DEFAULT 0)")
            c.execute("CREATE TABLE product_groups (id INTEGER PRIMARY KEY, name TEXT)")
            c.execute("INSERT INTO product_groups (name) VALUES ('POM')")
            c.execute("CREATE TABLE products (id INTEGER PRIMARY KEY, group_id INTEGER, name TEXT, stock_kg REAL DEFAULT 1000, buy_price REAL DEFAULT 0, sell_price REAL DEFAULT 0)")
            c.execute("INSERT INTO products (group_id, name, stock_kg) VALUES (1,'POM',1000)")
            c.execute("""CREATE TABLE sales_history (
                id INTEGER PRIMARY KEY, personnel_name TEXT, customer_id INTEGER, customer_name TEXT,
                product_name TEXT, qty REAL, price REAL, total_amount REAL, payment_type TEXT,
                payment_received REAL, created_at TEXT, note TEXT)""")
            c.execute("""CREATE TABLE wp_kaynak_olaylar (
                id INTEGER PRIMARY KEY, kaynak_anahtar TEXT UNIQUE, tarih TEXT, personel TEXT,
                ham_govde TEXT, ham_musteri_aday TEXT, belirsizlik TEXT, kalem_no INTEGER,
                sales_history_id INTEGER, tutar REAL, tahsilat REAL, durum TEXT)""")
            c.execute("INSERT INTO customers (name, remaining_debt, payment, shopping_count) VALUES ('OLGUN', 0, 0, 0)")
            conn.commit()
            kayitlar = [
                {"musteri_id": 1, "musteri_ad": "OLGUN", "urun_ad": "POM", "miktar": 10, "fiyat": 10,
                 "odeme_alindi": 0, "tarih": "1", "personel": "A", "ham_govde": "a", "kalem_no": 0, "durum": "TAMAM"},
                {"musteri_id": 1, "musteri_ad": "OLGUN", "urun_ad": "ABS", "miktar": 10, "fiyat": 10,
                 "odeme_alindi": 0, "tarih": "2", "personel": "A", "ham_govde": "b", "kalem_no": 0, "durum": "TAMAM"},
            ]

            def gid(cur, u):
                if "ABS" in str(u or "").upper():
                    raise RuntimeError("kontrollu hata")
                return 1

            with self.assertRaises(RuntimeError):
                wp_satis_kayit_servisi(c, conn, kayitlar, {}, {}, gid)
            c.execute("SELECT COUNT(*) FROM sales_history")
            self.assertEqual(c.fetchone()[0], 0)
            c.execute("SELECT stock_kg FROM products WHERE id=1")
            self.assertEqual(c.fetchone()[0], 1000)
            c.execute("SELECT remaining_debt, shopping_count FROM customers WHERE id=1")
            rem, shop = c.fetchone()
            self.assertEqual(rem, 0)
            self.assertEqual(shop, 0)
            conn.close()
        finally:
            os.remove(path)


if __name__ == "__main__":
    unittest.main()
