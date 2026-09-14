from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
import re
import unicodedata


MUSTERISIZ = "MÜŞTERİSİZ SATIŞ"


def normalize(metin):
    """Türkçe karşılaştırma anahtarı; kayıt adının yerine kullanılmaz."""
    s = str(metin or "").translate(str.maketrans({
        "İ": "i", "I": "i", "ı": "i",
        "Ş": "s", "ş": "s",
        "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u",
        "Ö": "o", "ö": "o",
        "Ç": "c", "ç": "c",
    })).lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", s))


def turkce_buyuk(metin):
    return str(metin or "").translate(
        str.maketrans({"i": "İ", "ı": "I"})
    ).upper()


HAMMADDE = frozenset("""
abs akrilik pmma pleksi pom delrin derlin asetal
pp moblen hdpe ldpe lldpe elteks i20 pe100
pa6 pa66 n6 n66 naylon6 naylon66
pc polikarbon pvc pet petg pbt tpu tpe eva hips
antisok kristal gpps
""".split())

KELIME_ESADLARI = {
    "pls": "plastik",
    "plstik": "plastik",
}

GENEL_KELIMELER = frozenset("""
plastik metal kalip sanayi ticaret limited sirketi
ltd sti as anonim geri donusum
abi abla bey hanim usta aksesuar
""".split())

BOS_ADLAR = {
    "",
    "musterisiz",
    "musterisiz satis",
    "genel polimer",
    "musteri",
}

VARSAYILAN_ALIAS = {
    "alaattin": "ALAATTİN AKSESUAR",
    "alattin": "ALAATTİN AKSESUAR",
    "alaaddin": "ALAATTİN AKSESUAR",
    "aladdin": "ALAATTİN AKSESUAR",
    "alddinn": "ALAATTİN AKSESUAR",
    "burhan": "BURHAN",
    "akrilik burhan": "BURHAN",
    "abs burhan": "BURHAN",
}


def ad_kelimeleri(ad):
    return tuple(
        KELIME_ESADLARI.get(w, w)
        for w in normalize(ad).split()
        if w not in HAMMADDE
    )


def kanonik_anahtar(ad):
    return " ".join(sorted(ad_kelimeleri(ad)))


def temiz_gorunen_ad(ad):
    kelimeler = [
        w for w in str(ad or "").split()
        if normalize(w) not in HAMMADDE
    ]
    return turkce_buyuk(" ".join(kelimeler).strip())


def wp_ek_borc(devir, onceki_wp_toplami, yeni_tutar):
    """D=devir, P=onceki WP toplamı, T=yeni tutar. remaining_debt yalnızca ek_borc kadar artar."""
    d = float(devir or 0.0)
    p = float(onceki_wp_toplami or 0.0)
    t = float(yeni_tutar or 0.0)
    return max(0.0, p + t - d) - max(0.0, p - d)


@dataclass(frozen=True)
class Eslesme:
    musteri_id: object
    musteri_ad: str
    durum: str
    gerekce: str
    adaylar: tuple = ()

    @property
    def kesin(self):
        return self.musteri_id is not None


class MusteriEsleyici:
    """
    kayitlar:
        [(id, name), ...] veya [(id, name, remaining_debt), ...]

    Yalnızca güvenilir referans / doğrulanmış müşteriler verilmeli.
    Benzerlik sonuçları adaydır; tek başına hesap ataması yapmaz.
    """

    def __init__(self, kayitlar, aliaslar=None):
        self.kayitlar = {}
        self.devirler = {}
        self.indeks = defaultdict(list)

        for row in kayitlar:
            cid, ad = row[0], str(row[1] or "").strip()
            if cid is None:
                continue
            if normalize(ad) in BOS_ADLAR:
                continue

            key = kanonik_anahtar(ad)
            if not key:
                continue
            if not set(key.split()) - GENEL_KELIMELER:
                continue

            if cid in self.kayitlar:
                if self.kayitlar[cid] != ad:
                    raise ValueError(
                        f"Aynı müşteri id'si farklı adlarla geldi: {cid}"
                    )
                continue

            self.kayitlar[cid] = ad
            self.devirler[cid] = float(row[2] or 0.0) if len(row) > 2 else 0.0
            self.indeks[key].append(cid)

        birlesik = dict(VARSAYILAN_ALIAS)
        if aliaslar:
            birlesik.update(aliaslar)

        self.aliaslar = {}
        for kaynak, hedef in birlesik.items():
            key = kanonik_anahtar(kaynak)
            onceki = self.aliaslar.get(key)
            if (
                onceki is not None
                and kanonik_anahtar(onceki) != kanonik_anahtar(hedef)
            ):
                raise ValueError(
                    f"Çakışan müşteri aliası: {kaynak}"
                )
            self.aliaslar[key] = hedef

    def _sec(self, kimlikler, tercih, durum):
        kimlikler = sorted(set(kimlikler))

        birebir = [
            cid for cid in kimlikler
            if normalize(self.kayitlar[cid]) == normalize(tercih)
        ]
        secilebilir = birebir if birebir else kimlikler

        if len(secilebilir) == 1:
            cid = secilebilir[0]
            return Eslesme(
                cid,
                temiz_gorunen_ad(self.kayitlar[cid]),
                durum,
                "Tek referans müşteri bulundu.",
            )

        return Eslesme(
            None,
            MUSTERISIZ,
            "BELIRSIZ",
            "Aynı anahtar birden fazla müşteri hesabına ait.",
            tuple(
                (cid, self.kayitlar[cid], 1.0)
                for cid in kimlikler
            ),
        )

    def coz(self, ham_ad):
        if normalize(ham_ad) in BOS_ADLAR:
            return Eslesme(
                None, MUSTERISIZ, "BULUNAMADI",
                "Müşteri adı belirtilmemiş.",
            )

        key = kanonik_anahtar(ham_ad)
        if not key or not set(key.split()) - GENEL_KELIMELER:
            return Eslesme(
                None, MUSTERISIZ, "BULUNAMADI",
                "Müşteriyi ayıran bir isim bulunamadı.",
            )

        if key in self.aliaslar:
            hedef = self.aliaslar[key]
            ids = self.indeks.get(kanonik_anahtar(hedef), [])
            if ids:
                return self._sec(ids, hedef, "ALIAS")

            return Eslesme(
                None, MUSTERISIZ, "REFERANS_EKSIK",
                f"Alias hedefi referans listesinde bulunamadı: {hedef}",
            )

        ids = self.indeks.get(key, [])
        if ids:
            return self._sec(ids, ham_ad, "TAM_ESLESME")

        adaylar = []
        sorgu = key.split()
        ayirt_edici = set(sorgu) - GENEL_KELIMELER

        for aday_key, aday_ids in self.indeks.items():
            hedef = aday_key.split()
            if len(sorgu) != len(hedef):
                continue
            if not set(hedef) - GENEL_KELIMELER:
                continue
            if not ayirt_edici:
                continue

            ayirt_skor = max(
                SequenceMatcher(None, a, b).ratio()
                for a in ayirt_edici
                for b in set(hedef) - GENEL_KELIMELER
            )
            skor = SequenceMatcher(None, key, aday_key).ratio()

            if skor >= 0.78 and ayirt_skor >= 0.78:
                for cid in aday_ids:
                    adaylar.append((
                        cid, self.kayitlar[cid], round(skor, 4)
                    ))

        adaylar.sort(
            key=lambda x: (-x[2], normalize(x[1]), str(x[0]))
        )

        return Eslesme(
            None,
            MUSTERISIZ,
            "INCELE" if adaylar else "BULUNAMADI",
            (
                "Benzer adlar var; kimlik eşitliği doğrulanmalı."
                if adaylar else
                "Referans müşteri eşleşmesi bulunamadı."
            ),
            tuple(adaylar[:5]),
        )

    def devir(self, musteri_id):
        return float(self.devirler.get(musteri_id, 0.0) or 0.0)
