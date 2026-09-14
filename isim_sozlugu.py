# isim_sozlugu.py

ERKEK_ISIMLERI = {
    "MEHMET", "MUSTAFA", "ALİ", "AHMET", "HASAN", "HÜSEYİN", "İBRAHİM", "İSMAİL", "OSMAN", "HALİL",
    "SÜLEYMAN", "YUSUF", "ÖMER", "KEMAL", "RAMAZAN", "MURAT", "HAKAN", "GÖKHAN", "VOLKAN", "SERKAN",
    "BURAK", "YASİN", "FURKAN", "EMRE", "ENES", "YUNUS", "FATİH", "SELÇUK", "KADİR", "BEKİR",
    "ADEM", "CENGİZ", "ERDAL", "BARIŞ", "TOLGA", "ERHAN", "SİNAN", "KENAN", "VEYSEL", "MİKAİL",
    "ZÜHTÜ", "İDRİS", "SEZER", "ALAATTİN", "BURHAN", "MAHMUT", "SEYİT", "SABRİ", "METİN", "EMİN",
    "CEYHUN", "SUAT", "DOĞAN", "YAVUZ", "AYHAN", "ERKAN", "ŞEREF", "İMDAT", "YAKUP", "BÜLENT",
    "OKTAY", "SALİH", "CELAL", "CEMAL", "ENGİN", "NİHAT", "UĞUR", "ZAFER", "İHSAN", "RUŞEN",
    "CÜNEYT", "TURAN", "SAMET", "SEDAT", "ALPASLAN", "MUHSİN", "BATUHAN", "OĞUZHAN", "İLKER", "AYDIN",
    "RECAİ", "ŞENOL", "TAYFUN", "YAHYA", "BİLAL", "CUMA", "DAVUT", "EYÜP", "HARUN", "İLYAS", "LOKMAN",
    "MUSA", "İSA", "NUH", "SADIK", "TAHİR", "ÜMİT", "VELİ", "YAŞAR", "ZEKİ", "AKIN", "ALİCAN", "ARİF",
    "CAN", "CEM", "COŞKUN", "DURSUN", "EKREM", "ERCAN", "FAHRETTİN", "FARUK", "GÜRKAN", "HAYRİ", "İZZET",
    "KASIM", "LEVENT", "MERT", "NECDET", "NEVZAT", "ORHAN", "ÖZCAN", "ÖZGÜR", "RECEP", "REMZİ", "SAİT",
    "SEÇİM", "SELİM", "SEZAİ", "ŞABAN", "ŞEVKİ", "TARIK", "TEVFİK", "TUNCAY", "VEDAT", "YALÇIN", "YENER"
}

SOYISIMLER = {
    "YILMAZ", "KAYA", "DEMİR", "ÇELİK", "ŞAHİN", "YILDIZ", "YILDIRIM", "ÖZTÜRK", "AYDIN", "ÖZDEMİR",
    "ARSLAN", "DOĞAN", "KILIÇ", "ASLAN", "ÇETİN", "KARA", "KOÇ", "KURT", "ÖZKAN", "ŞİMŞEK",
    "POLAT", "ÖZCAN", "KORKMAZ", "ÇAKIR", "ERDOĞAN", "YAVUZ", "CAN", "AKTAŞ", "GÜLER", "YALÇIN",
    "GÜNEŞ", "BOZKURT", "BULUT", "KESKİN", "TURAN", "ÖZER", "IŞIK", "YILDIRAN", "YILMAZER", "KAPLAN",
    "YÜCEL", "AKGÜN", "ERDEM", "KÖSE", "AKIN", "ÇETİNKAYA", "SOYLU", "AY", "ÇOŞKUN", "YAZICI", "MUTLU",
    "BAŞARAN", "TÜRK", "ÇALIŞKAN", "AKSOY", "BAŞ", "ÖZ", "GÜL", "AYHAN", "EREN", "ŞEN", "KOCAMAN",
    "KARADAĞ", "AVCI", "GÖK", "YAMAN", "ALTUN", "GÜNGÖR", "TAŞ", "ÇİÇEK", "ÖZMEN", "UYSAL", "AKGÜL",
    "TURGUT", "GÜVEN", "KARACA", "TEKİN", "BİÇER", "ER", "SARI", "YİĞİT", "GÜRBÜZ", "KARTAL",
    "TUNCER", "KÖKSAL", "AKSU", "BAYRAM", "ÖZDEN", "ŞAHİNER", "ALBAYRAK", "GÖÇER", "EROL", "TAŞDEMİR",
    "DURMAZ", "KARAKAYA", "ÇELİKKAYA", "KILIÇARSLAN", "KARATAŞ", "GÜNDÜZ", "DUMAN", "AK", "ERİŞ", "BOZ"
}

ISIM_DUZELTME_HARITASI = {
    "ALATTİN": "ALAATTİN",
    "ALAETİN": "ALAATTİN",
    "MEHMET": "MEHMET",
    "MEMET": "MEHMET",
    "MUAMER": "MUAMER",
    "MAMUT": "MAHMUT"
}
