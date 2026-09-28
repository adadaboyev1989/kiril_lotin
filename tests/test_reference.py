# -*- coding: utf-8 -*-
import unittest

from reference import cyr_to_lat, lat_to_cyr, word_convert

CYR_LAT = [
    ("Салом дунё", "Salom dunyo"),
    ("Ўзбекистон", "Oʻzbekiston"),
    ("ЎЗБЕКИСТОН", "OʻZBEKISTON"),
    ("Тошкент шаҳри", "Toshkent shahri"),
    ("ШАҲАР", "SHAHAR"),
    ("Шаҳар", "Shahar"),
    ("ер", "yer"),
    ("Европа", "Yevropa"),
    ("поезд", "poyezd"),
    ("цирк", "sirk"),
    ("милиция", "militsiya"),
    ("станция", "stansiya"),
    ("маъно", "maʼno"),
    ("объект", "obyekt"),
    ("тоғ", "togʻ"),
    ("қишлоқ", "qishloq"),
    ("Ёшлар", "Yoshlar"),
    ("мактаб, 2024 йил.", "maktab, 2024 yil."),
    ("Ю", "Yu"),
    ("ЮНЕСКО", "YUNESKO"),
    ("тайёр", "tayyor"),
]

LAT_CYR = [
    ("Salom dunyo", "Салом дунё"),
    ("O'zbekiston", "Ўзбекистон"),
    ("Oʻzbekiston", "Ўзбекистон"),
    ("O‘ZBEKISTON", "ЎЗБЕКИСТОН"),
    ("Toshkent shahri", "Тошкент шаҳри"),
    ("SHAHAR", "ШАҲАР"),
    ("yo'l", "йўл"),
    ("tog'", "тоғ"),
    ("ma'no", "маъно"),
    ("MA'NO", "МАЪНО"),
    ("Is'hoq", "Исҳоқ"),
    ("mashhur", "машҳур"),
    ("ekin", "экин"),
    ("poeziya", "поэзия"),
    ("yetti", "етти"),
    ("kel", "кел"),
    ("choy", "чой"),
    ("tayyor", "тайёр"),
    ("'Salom'", "'Салом'"),
    ("https://gov.uz sayti", "https://gov.uz сайти"),
    ("info@mail.uz manzil", "info@mail.uz манзил"),
    ("Qo'qon", "Қўқон"),
    ("G'ulom", "Ғулом"),
]


class TestTranslit(unittest.TestCase):
    def test_cyr_to_lat(self):
        for src, expected in CYR_LAT:
            with self.subTest(src=src):
                self.assertEqual(cyr_to_lat(src), expected)

    def test_ascii_apostrophe(self):
        self.assertEqual(cyr_to_lat("Ўзбекистон маъно", ascii_apos=True), "O'zbekiston ma'no")

    def test_lat_to_cyr(self):
        for src, expected in LAT_CYR:
            with self.subTest(src=src):
                self.assertEqual(lat_to_cyr(src), expected)

    def test_round_trip(self):
        text = "Ўзбекистон Республикаси пойтахти Тошкент шаҳри"
        self.assertEqual(lat_to_cyr(cyr_to_lat(text)), text)


# Foydalanuvchi xabar qilgan xatolar (Word'da harflar alohida o'girilgan edi)
# va rasmiy hujjatlarda ko'p uchraydigan so'zlar.
WORD_LAT_CYR = [
    ("viloyati", "вилояти"),
    ("boshqarma", "бошқарма"),
    ("boshlig‘i", "бошлиғи"),
    ("shahar", "шаҳар"),
    ("Toshkent shahar hokimligi", "Тошкент шаҳар ҳокимлиги"),
    ("Samarqand viloyati sog‘liqni saqlash boshqarmasi boshlig‘i",
     "Самарқанд вилояти соғлиқни сақлаш бошқармаси бошлиғи"),
    ("O‘zbekiston Respublikasi Prezidentining qarori",
     "Ўзбекистон Республикаси Президентининг қарори"),
    ("ta’lim va ma’lumot", "таълим ва маълумот"),
    ("Qo‘shimcha ma'lumot uchun: www.gov.uz", "Қўшимча маълумот учун: www.gov.uz"),
    ("yoshlar ishlari agentligi", "ёшлар ишлари агентлиги"),
    ("yangi yil, sentyabr, oktyabr", "янги йил, сентябр, октябр"),
    ("choy, chiroq, ishchi", "чой, чироқ, ишчи"),
    ("yo‘l, yo'nalish, tog‘", "йўл, йўналиш, тоғ"),
    ("VILOYATI BOSHQARMASI", "ВИЛОЯТИ БОШҚАРМАСИ"),
    ("SHAHAR", "ШАҲАР"),
]

WORD_CYR_LAT = [
    ("келгусида", "kelgusida"),
    ("вилояти", "viloyati"),
    ("бошқарма бошлиғи", "boshqarma boshligʻi"),
    ("шаҳар", "shahar"),
    ("Келгусида вилоят ҳокимлиги", "Kelgusida viloyat hokimligi"),
    ("Ўзбекистон Республикаси Вазирлар Маҳкамаси", "Oʻzbekiston Respublikasi Vazirlar Mahkamasi"),
    ("ер, ерда, Европа, поезд, лойиҳа", "yer, yerda, Yevropa, poyezd, loyiha"),
    ("мактаб директори, педагог, келажак", "maktab direktori, pedagog, kelajak"),
    ("СЕНТЯБРЬ, ОКТЯБРЬ", "SENTYABR, OKTYABR"),
    ("ташкилот, ташаббус, маъсул", "tashkilot, tashabbus, maʼsul"),
    ("милиция, цирк, станция", "militsiya, sirk, stansiya"),
    ("ШАҲАР ҲОКИМЛИГИ", "SHAHAR HOKIMLIGI"),
]


class TestWordMode(unittest.TestCase):
    """Word'dagi so'zma-so'z o'girish matnni butunligicha o'girish bilan bir xil bo'lishi kerak."""

    def test_lat_to_cyr(self):
        for src, expected in WORD_LAT_CYR:
            with self.subTest(src=src):
                self.assertEqual(word_convert(src, to_latin=False), expected)
                self.assertEqual(lat_to_cyr(src), expected)

    def test_cyr_to_lat(self):
        for src, expected in WORD_CYR_LAT:
            with self.subTest(src=src):
                self.assertEqual(word_convert(src, to_latin=True), expected)
                self.assertEqual(cyr_to_lat(src), expected)

    def test_word_mode_equals_string_mode(self):
        texts = [s for s, _ in WORD_LAT_CYR] + [s for s, _ in LAT_CYR]
        for t in texts:
            with self.subTest(t=t):
                self.assertEqual(word_convert(t, to_latin=False), lat_to_cyr(t))
        texts = [s for s, _ in WORD_CYR_LAT] + [s for s, _ in CYR_LAT]
        for t in texts:
            with self.subTest(t=t):
                self.assertEqual(word_convert(t, to_latin=True), cyr_to_lat(t))


if __name__ == "__main__":
    unittest.main()
