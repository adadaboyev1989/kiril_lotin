# -*- coding: utf-8 -*-
import unittest

from reference import cyr_to_lat, lat_to_cyr

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


if __name__ == "__main__":
    unittest.main()
