"""Targeted tests of real main1.py functions without importing Qt or missing modules.

AST extraction avoids running the desktop application or its startup migrations.
The matcher and debt helper are explicit test doubles; these tests do NOT claim
to test the missing wp_musteri_esleme.py implementation or the complete parser.
"""
import ast
import hashlib
from pathlib import Path
import sqlite3
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

SOURCE = Path(__file__).resolve().parents[1] / "main1.py"
TREE = ast.parse(SOURCE.read_text(encoding="utf-8"), filename=str(SOURCE))
NAMES = {
    "init_db", "mevcut_musterilerle_esle", "wp_musterisiz_id_al",
    "wp_kaynak_anahtari_uret", "wp_satis_kayit_servisi",
}

def norm(value):
    return str(value or "").translate(str.maketrans(
        "İIıŞşĞğÜüÖöÇç", "iii ss gg uu oo cc".replace(" ", "")
    )).lower().strip()

def namespace():
    ns = {
        "sqlite3": sqlite3, "hashlib": hashlib,
        "MUSTERISIZ": "MÜŞTERİSİZ SATIŞ",
        "turkce_toleransli_metin": norm,
    }
    nodes = [n for n in TREE.body
             if isinstance(n, ast.FunctionDef) and n.name in NAMES]
    if {n.name for n in nodes} != NAMES:
        raise AssertionError("Expected application functions missing")
    exec(compile(ast.Module(body=nodes, type_ignores=[]),
                 str(SOURCE), "exec"), ns)
    return ns

def result(cid=None, ad="MÜŞTERİSİZ SATIŞ", durum="BULUNAMADI"):
    return SimpleNamespace(
        kesin=cid is not None, musteri_id=cid,
        musteri_ad=ad, durum=durum,
    )

class MatchingContractTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace()

    def configure(self, answer):
        self.ns["MusteriEsleyici"] = lambda rows: SimpleNamespace(
            coz=lambda name: answer
        )

    def test_unknown_name_is_not_returned_as_new_customer(self):
        self.configure(result())
        self.assertEqual(
            self.ns["mevcut_musterilerle_esle"]("YAZARSIN ADAMA", []),
            (None, "MÜŞTERİSİZ SATIŞ", 0.0),
        )

    def test_ambiguous_name_stays_unassigned(self):
        self.configure(result(durum="BELIRSIZ"))
        self.assertIsNone(self.ns["mevcut_musterilerle_esle"](
            "ALİ", [(1, "ALİ USTA", 5), (2, "ALİ KALIP", 9)]
        )[0])

    def test_verified_id_uses_its_actual_reference_balance(self):
        self.configure(result(2, "ALİ KALIP", "TAM_ESLESME"))
        rows = [(1, "ALİ USTA", 5), (2, "ALİ KALIP", 9)]
        for data in (rows, list(reversed(rows))):
            self.assertEqual(
                self.ns["mevcut_musterilerle_esle"]("ALİ KALIP", data),
                (2, "ALİ KALIP", 9.0),
            )

    def test_missing_reference_id_is_not_accepted(self):
        self.configure(result(999, "ALİ KALIP", "TAM_ESLESME"))
        self.assertEqual(self.ns["mevcut_musterilerle_esle"](
            "ALİ KALIP", [(1, "ALİ USTA", 5)]
        ), (None, "MÜŞTERİSİZ SATIŞ", 0.0))

    def test_matcher_failure_is_visible(self):
        self.ns["MusteriEsleyici"] = Mock(side_effect=RuntimeError("matcher"))
        with self.assertRaisesRegex(RuntimeError, "matcher"):
            self.ns["mevcut_musterilerle_esle"]("ALİ", [])

class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = str(Path(self.temp.name) / "test.db")
        self.ns = namespace()
        self.ns["DB_NAME"] = self.path

    def initialize(self):
        self.ns["init_db"]()
        conn = sqlite3.connect(self.path)
        self.addCleanup(conn.close)
        return conn

    def test_paid_balance_is_not_resurrected_on_restart(self):
        conn = self.initialize()
        conn.execute("ALTER TABLE customers ADD COLUMN balance REAL")
        conn.execute(
            "INSERT INTO customers(name, remaining_debt, balance) VALUES (?,0,700)",
            ("TAHSİN",),
        )
        conn.commit()
        self.ns["init_db"]()
        self.ns["init_db"]()
        self.assertEqual(conn.execute(
            "SELECT remaining_debt FROM customers WHERE name='TAHSİN'"
        ).fetchone()[0], 0)

    def test_legacy_balance_is_migrated_only_when_column_is_added(self):
        with sqlite3.connect(self.path) as conn:
            conn.execute(
                "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT, balance REAL)"
            )
            conn.execute("INSERT INTO customers VALUES (1,'ALİ USTA',700)")
        self.ns["init_db"]()
        with sqlite3.connect(self.path) as conn:
            self.assertEqual(conn.execute(
                "SELECT remaining_debt FROM customers WHERE id=1"
            ).fetchone()[0], 700)
            conn.execute("UPDATE customers SET remaining_debt=0 WHERE id=1")
        self.ns["init_db"]()
        with sqlite3.connect(self.path) as conn:
            self.assertEqual(conn.execute(
                "SELECT remaining_debt FROM customers WHERE id=1"
            ).fetchone()[0], 0)

    def prepare_writer(self):
        conn = self.initialize()
        self.ns["MusteriEsleyici"] = lambda rows: SimpleNamespace(
            coz=lambda name: result()
        )
        self.ns["standart_urun_ve_grup_belirle"] = lambda ad, _: ("POM", ad)
        # Out-of-scope missing module: spy rather than a substitute debt formula.
        self.ns["wp_musteri_borcunu_guncelle"] = Mock(return_value=0)
        return conn

    def record(self, **changes):
        d = {
            "olay_tipi": "SATIS", "musteri_ad": "YAZARSIN ADAMA",
            "urun_ad": "TEST POM", "miktar": 10, "fiyat": 30,
            "tarih": "27.03.2023 16:39:44", "personel": "YAVUZ",
            "ham_govde": "10 kg TEST POM 30 tl", "durum": "TAMAM",
        }
        d.update(changes)
        return d

    def write(self, conn, records, products=None):
        return self.ns["wp_satis_kayit_servisi"](
            conn.cursor(), conn, records, {},
            products if products is not None else {},
            lambda cur, name: cur.execute(
                "SELECT id FROM product_groups WHERE name='POM'"
            ).fetchone()[0],
        )

    def test_unknown_customer_keeps_sale_without_creating_garbage_account(self):
        conn = self.prepare_writer()
        self.write(conn, [self.record()])
        self.assertEqual(conn.execute("SELECT name FROM customers").fetchall(),
                         [("MÜŞTERİSİZ SATIŞ",)])
        self.assertEqual(conn.execute(
            "SELECT customer_name, qty FROM sales_history"
        ).fetchall(), [("MÜŞTERİSİZ SATIŞ", 10.0)])

    def test_stale_parser_id_does_not_create_dangling_sale(self):
        conn = self.prepare_writer()
        self.write(conn, [self.record(musteri_id=99999)])
        self.assertEqual(conn.execute(
            "SELECT COUNT(*) FROM sales_history s JOIN customers c ON c.id=s.customer_id"
        ).fetchone()[0], 1)

    def test_valid_id_uses_stored_customer_name(self):
        conn = self.prepare_writer()
        cur = conn.execute("INSERT INTO customers(name) VALUES ('ALİ USTA')")
        cid = cur.lastrowid
        conn.commit()
        self.write(conn, [self.record(musteri_id=cid, musteri_ad="YANLIŞ AD")])
        self.assertEqual(conn.execute(
            "SELECT customer_id, customer_name FROM sales_history"
        ).fetchone(), (cid, "ALİ USTA"))

    def test_rollback_restores_database_and_product_cache(self):
        conn = self.prepare_writer()
        self.ns["wp_musteri_borcunu_guncelle"].side_effect = RuntimeError("stop")
        before = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        products = {}
        with self.assertRaisesRegex(RuntimeError, "stop"):
            self.write(conn, [self.record()], products)
        self.assertEqual(products, {})
        self.assertEqual(conn.execute(
            "SELECT COUNT(*) FROM products"
        ).fetchone()[0], before)
        for table in ("customers", "sales_history", "wp_kaynak_olaylar"):
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM " + table
            ).fetchone()[0], 0)

class StartupAndPreviewTests(unittest.TestCase):
    def test_startup_does_not_run_customer_repair(self):
        guards = [n for n in TREE.body if isinstance(n, ast.If)
                  and "__name__" in ast.unparse(n.test)]
        self.assertEqual(len(guards), 1)
        called = {n.func.id for n in ast.walk(guards[0])
                  if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)}
        forbidden = {
            "tum_kopuk_satislari_gercek_musteriye_bagla",
            "bosluksuz_ve_birlesik_isimleri_birlestir",
            "fiil_ve_cop_carileri_temizle",
        }
        self.assertFalse(called & forbidden)

    def test_editing_customer_cell_invalidates_previous_id(self):
        app = next(n for n in TREE.body if isinstance(n, ast.ClassDef)
                   and n.name == "BenimPOSPlastik")
        method = next(n for n in app.body if isinstance(n, ast.FunctionDef)
                      and n.name == "wp_tablo_hucre_degisti")
        ns = {"buyuk_harf": lambda s: s.upper(),
              "WhatsAppSatisAyristirici": SimpleNamespace(
                  tur_tespit_et=lambda s: "BİREYSEL")}
        exec(compile(ast.Module(body=[method], type_ignores=[]),
                     str(SOURCE), "exec"), ns)
        class Cell:
            def text(self):
                return "YENİ MÜŞTERİ"
        class Table:
            def item(self, row, col):
                return Cell() if col == 3 else None
            def blockSignals(self, value):
                pass
        view = SimpleNamespace(
            table_wp_onizleme=Table(),
            wp_cozumlenen_veriler=[{
                "musteri_id": 123, "musteri_ad": "ESKİ MÜŞTERİ"
            }],
            lbl_wp_durum_ozet=SimpleNamespace(setText=lambda text: None),
        )
        ns["wp_tablo_hucre_degisti"](view, 0, 3)
        self.assertIsNone(view.wp_cozumlenen_veriler[0]["musteri_id"])
        self.assertEqual(view.wp_cozumlenen_veriler[0]["musteri_ad"],
                         "YENİ MÜŞTERİ")

if __name__ == "__main__":
    unittest.main()
