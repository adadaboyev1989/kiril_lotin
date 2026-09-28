# Kirill-Lotin

Microsoft **Word** va **Excel** uchun o'zbek tilidagi matnni **kirill alifbosidan lotin alifbosiga** va **lotindan kirillga** o'giradigan qo'shimcha (add-in).

O'rnatilgach, Word va Excel har safar ochilganda qo'shimcha avtomatik yuklanadi va lentada **Kirill-Lotin** yorlig'i paydo bo'ladi.

## Imkoniyatlar

| | Word | Excel |
|---|---|---|
| Belgilangan matnni o'girish | ✅ | ✅ (kataklar, matnli shakllar) |
| Hech narsa belgilanmasa — butun hujjat / varaq | ✅ (kolontitul, izoh, matn maydonlari ham) | ✅ |
| Formatlash saqlanadi (shrift, rang, jadval) | ✅ | ✅ (formulalar va sonlarga tegilmaydi) |
| Bitta Ctrl+Z bilan bekor qilish | ✅ | ✅ |
| Tezkor tugmalar | Alt+Shift+L / Alt+Shift+K | Alt+Shift+L / Alt+Shift+K |
| Formulalar | — | `=LOTINGA(A1)`, `=KIRILGA(A1)` |

* **Alt+Shift+L** — Kirill → Lotin (Word da tugmalar qo'shimcha yuklanganda `AutoExec` makrosi orqali o'rnatiladi)
* **Alt+Shift+K** — Lotin → Kirill
* Havolalar (`https://...`, `www...`) va e-pochta manzillari kirillga o'girilmaydi.
* **Sozlamalar → «Oddiy apostrof (')»**: belgilansa `o'`, `g'` oddiy `'` bilan yoziladi; aks holda rasmiy imlodagi `oʻ`, `gʻ` (U+02BB) va tutuq belgisi `ʼ` (U+02BC) ishlatiladi. Lotindan kirillga o'girishda barcha turdagi apostroflar (`'` `` ` `` `‘` `’` `ʻ` `ʼ`) tushuniladi.

## O'rnatish

Talablar:

* Windows 7, 8, 8.1, 10, 11 (Windows XP/Vista da — Windows PowerShell 2.0 o'rnatilgan bo'lsa);
* Microsoft Office **2003, 2007, 2010, 2013, 2016, 2019, 2021, 365** — 32 va 64 bitli, jumladan rus va boshqa tildagi Office.

| Office | Qo'shimcha fayllari | Tugmalar |
|---|---|---|
| 2007 va yangiroq | `KirillLotin.dotm`, `KirillLotin.xlam` | «Kirill-Lotin» lenta yorlig'i |
| 2003 | `KirillLotin.dot`, `KirillLotin.xla` (o'rnatish paytida Office yordamida yig'iladi) | «Kirill-Lotin» asboblar paneli |

Qo'shimchalar Word/Excel ning haqiqiy STARTUP papkasiga joylanadi (u Office'ning o'zidan so'raladi, shuning uchun lokallashtirilgan Office'da ham to'g'ri joy tanlanadi).

Muammo bo'lsa, o'rnatish jurnali: `%TEMP%\KirillLotin-install.log` (nusxasi `%APPDATA%\KirillLotin\install.log`).

### 1-usul: EXE o'rnatuvchi (tavsiya etiladi)

1. [`dist/KirillLotin-Setup.exe`](dist/KirillLotin-Setup.exe) faylini yuklab oling.
2. Word va Excel oynalarini yoping.
3. `KirillLotin-Setup.exe` ni ishga tushiring va **Keyingi** tugmalarini bosing.
4. Tayyor! Word yoki Excel ni oching — lentada **Kirill-Lotin** yorlig'i paydo bo'ladi.

Agar Word yoki Excel ochiq bo'lsa, o'rnatuvchi ochiq Office dasturlari ro'yxatini ko'rsatadi:

* **Qayta tekshirish** — dasturlarni o'zingiz (ishingizni saqlab) yopganingizdan keyin bosing;
* **Majburan yopish** — ro'yxatdagi barcha ochiq Office dasturlarini (Word, Excel, PowerPoint, Outlook, Access, Publisher, OneNote, Visio, Project) darhol yopadi. **Saqlanmagan o'zgarishlar yo'qoladi.**

Administrator huquqi kerak emas. Dastur **Sozlamalar → Ilovalar** (yoki **Boshqaruv paneli → Dasturlar**) ro'yxatida «Kirill-Lotin (Word va Excel uchun)» nomi bilan ko'rinadi va o'sha yerdan o'chiriladi.

> **Windows SmartScreen** «Noma'lum nashriyotchi» deb ogohlantirishi mumkin, chunki fayl raqamli imzo bilan imzolanmagan. **Batafsil → Baribir ishga tushirish** ni bosing.

Jimjit o'rnatish (bir nechta kompyuterga tarqatish uchun): `KirillLotin-Setup.exe /S`.

### 2-usul: skript orqali

1. Loyihani yuklab oling: **Code → Download ZIP** va arxivni istalgan papkaga oching.
2. Word va Excel oynalarini yoping.
3. **`install.bat`** faylini ikki marta bosing.

O'rnatuvchi nima qiladi:

* `addins/` papkasidagi tayyor `KirillLotin.dotm` (Word) va `KirillLotin.xlam` (Excel) fayllarini Office avtomatik yuklaydigan papkalarga nusxalaydi:
  * `%APPDATA%\Microsoft\Word\STARTUP\KirillLotin.dotm`
  * `%APPDATA%\Microsoft\Excel\XLSTART\KirillLotin.xlam`

  Bu papkalar Office uchun ishonchli joy hisoblanadi, shuning uchun makros haqida ogohlantirish chiqmaydi;
* nusxalashdan oldin har bir qo'shimchani Word/Excel da bir marta yashirin ochib tekshiradi (odatda bir necha soniya; Office javob bermasa 90 soniyadan keyin tekshiruv to'xtatiladi);
* agar tayyor fayl ishlamasa, zaxira usulda qo'shimchani Word/Excel ning o'zi yordamida `src/` dan qayta yig'adi.

Tekshiruvsiz eng tez o'rnatish: `powershell -ExecutionPolicy Bypass -File install.ps1 -SkipCheck`.

Administrator huquqi kerak emas — hammasi joriy foydalanuvchi uchun o'rnatiladi.

### Lentada «Kirill-Lotin» chiqmasa

O'rnatuvchi o'rnatishdan keyin Word'ni ishga tushirib, qo'shimcha haqiqatan yuklanishini tekshiradi:

* tayyor fayl ishlamasa — qo'shimchani Word/Excel ning o'zi yordamida avtomatik qayta yig'adi;
* Office «O'chirilgan elementlar» (Disabled Items) ro'yxatiga kiritgan bo'lsa — u yerdan olib tashlaydi;
* STARTUP/XLSTART papkalarini ishonchli joy sifatida qo'shadi;
* Office sozlamalari qo'shimchani to'sayotgan bo'lsa (masalan, *«Barcha ilova qo'shimchalarini o'chirish»*, *«Ilova qo'shimchalari imzolangan bo'lishi shart»*, *«Barcha ishonchli joylarni o'chirish»*, guruh siyosati bilan VBA o'chirilgan) — nima qilish kerakligini yozib, faylda ko'rsatadi.

Keyinchalik tekshirish uchun: **Boshlash → Kirill-Lotin → Kirill-Lotin tashxis**. Hisobot (`%TEMP%\KirillLotin-tashxis.txt`) Notepad'da ochiladi — muammo hal bo'lmasa, shu faylni yuboring.

Qo'lda tekshirish: **Word → Fayl → Parametrlar → Qo'shimchalar → Boshqarish: O'chirilgan elementlar → O'tish** (ro'yxatda KirillLotin bo'lsa, yoqing), so'ng **Ishonch markazi → Ishonch markazi parametrlari → Qo'shimchalar** bo'limida belgilar olib tashlanganini tekshiring.

### O'chirish

EXE bilan o'rnatilgan bo'lsa — **Sozlamalar → Ilovalar** dan «Kirill-Lotin» ni o'chiring. Skript bilan o'rnatilgan bo'lsa — Word va Excel ni yopib, **`uninstall.bat`** ni ikki marta bosing.

### Qo'lda o'rnatish (o'rnatuvchi ishlamasa)

1. Word yoki Excel da **Alt+F11** (VBA muharriri) ni bosing.
2. **File → Import File...** orqali `src/KLCore.bas` va `src/KLWord.bas` (Word uchun) yoki `src/KLExcel.bas` (Excel uchun) fayllarini `Normal` / `PERSONAL.XLSB` loyihasiga import qiling.
3. Makroslarni **Alt+F8** orqali ishga tushiring: `KL_WordToLatin`, `KL_WordToCyrillic` (Word) yoki `KL_ExcelToLatin`, `KL_ExcelToCyrillic` (Excel).

Bu usulda lenta yorlig'i bo'lmaydi, lekin makroslarni tezkor panelga yoki tugmalarga o'zingiz biriktirishingiz mumkin.

## Transliteratsiya qoidalari

Rasmiy o'zbek lotin alifbosi (1995) qoidalari asosida:

| Kirill | Lotin | Izoh |
|---|---|---|
| Ў ў | Oʻ oʻ | |
| Ғ ғ | Gʻ gʻ | |
| Қ қ | Q q | |
| Ҳ ҳ | H h | |
| Х х | X x | |
| Ш ш, Щ щ | Sh sh | |
| Ч ч | Ch ch | |
| Ж ж | J j | |
| Й й | Y y | |
| Ё ё | Yo yo | |
| Ю ю | Yu yu | |
| Я я | Ya ya | |
| Е е | Ye ye / E e | so'z boshida, unli, `ъ`, `ь` dan keyin — **ye** (ер → yer, поезд → poyezd), qolgan holatda — **e** |
| Э э | E e | |
| Ц ц | S s / Ts ts | unlidan keyin — **ts** (милиция → militsiya), qolgan holatda — **s** (цирк → sirk) |
| Ъ ъ | ʼ | `е, ё, ю, я` oldidan tushib qoladi (объект → obyekt) |
| Ь ь | — | tushib qoladi |
| Ы ы | I i | |

Bosh harflar: `Шаҳар → Shahar`, `ШАҲАР → SHAHAR`.

Lotindan kirillga: `sh → ш`, `ch → ч`, `oʻ/o' → ў`, `gʻ/g' → ғ`, `yo/yu/ya/ye → ё/ю/я/е` (`yo'l → йўл`), `e` so'z boshida va unlidan keyin `э` (ekin → экин, poeziya → поэзия), harflar orasidagi apostrof `ъ` (ma'no → маъно), `s'h → сҳ` (Is'hoq → Исҳоқ).

### Cheklovlar

Lotindan kirillga o'girish to'liq bir qiymatli emas, ayrim so'zlarni qo'lda tuzatish kerak bo'lishi mumkin:

* `ts` o'zgartirilmaydi (`tsirk → тсирк`, lekin `otsiz → отсиз` to'g'ri chiqadi);
* ruscha o'zlashma so'zlardagi `ь`, `ъ` tiklanmaydi (`obyekt → обект`);
* ingliz tilidagi so'zlar ham o'giriladi — bunday matnni belgilamang.

Excel da katak ichidagi qisman formatlash (masalan, bitta so'z qalin) o'girishdan keyin yo'qoladi.

## Loyiha tuzilishi

```
dist/KirillLotin-Setup.exe      tayyor EXE o'rnatuvchi
addins/KirillLotin.dotm         tayyor Word qo'shimchasi (build_addins.py yasaydi)
addins/KirillLotin.xlam         tayyor Excel qo'shimchasi (build_addins.py yasaydi)
installer/build_addins.py       .dotm/.xlam ni src/ dan Office'siz yasovchi skript
installer/KirillLotin.nsi       EXE o'rnatuvchi skripti (NSIS)
installer/KirillLotin.ico       dastur belgisi (make_icon.py bilan yasalgan)
install.bat / install.ps1       o'rnatuvchi
uninstall.bat / uninstall.ps1   o'chiruvchi
src/KLCore.bas                  transliteratsiya yadrosi (Word va Excel uchun umumiy)
src/KLWord.bas                  Word buyruqlari
src/KLExcel.bas                 Excel buyruqlari va formulalari
src/ThisWorkbook.vba            Excel qo'shimchasi ochilganda tezkor tugmalarni o'rnatadi
src/customUI.xml                lenta (Ribbon) yorlig'i
tests/reference.py              qoidalarning Python nusxasi
tests/test_reference.py         qoidalar testlari
```

`.bas`, `.vba`, `.ps1`, `.bat` fayllari Windows qator oxiri (CRLF) bilan saqlanadi va faqat ASCII belgilardan iborat — VBA muharriri UTF-8 ni tushunmaydi, shuning uchun kirill harflar `ChrW()` orqali yoziladi.

### Testlar

Qoidalar `tests/reference.py` da Python'da ham yozilgan (VBA kodi bilan bir xil), ularni Office'siz tekshirish mumkin:

```
cd tests
python -m unittest -v
```

Qoidani o'zgartirsangiz, `src/KLCore.bas` va `tests/reference.py` ni birga yangilang.

### Qo'shimchalar va EXE ni qayta yig'ish

`src/` dagi kod o'zgarganda (Python 3 va [NSIS 3](https://nsis.sourceforge.io) kerak, Linux: `apt install nsis`):

```
python3 installer/build_addins.py     # addins/KirillLotin.dotm, addins/KirillLotin.xlam
makensis installer/KirillLotin.nsi    # dist/KirillLotin-Setup.exe
```

`build_addins.py` VBA loyihasini ([MS-OVBA]) va OLE konteynerni ([MS-CFB]) to'g'ridan-to'g'ri yozadi, Office kerak emas. Loyihada kompilyatsiya keshi yo'q (`_VBA_PROJECT` versiyasi `0xFFFF`), shuning uchun Office birinchi ochishda kodni manbadan o'zi kompilyatsiya qiladi — 32 va 64 bitli Office uchun bitta fayl.

GitHub'da har bir push da hammasi avtomatik yig'iladi (Actions → artifact), `v1.1.0` kabi teg qo'yilganda esa Releases bo'limiga qo'shiladi.
