## Kirill-Lotin — Word va Excel uchun

O'zbek tilidagi matnni **kirill → lotin** va **lotin → kirill** alifbosiga o'giradigan qo'shimcha.

### O'rnatish
1. **`KirillLotin-Setup.exe`** ni yuklab oling.
2. Word va Excel oynalarini yoping.
3. Faylni ishga tushiring (Windows SmartScreen ogohlantirsa: **Batafsil → Baribir ishga tushirish**).
4. Word yoki Excel ni oching — lentada **Kirill-Lotin** yorlig'i paydo bo'ladi.

Tezkor tugmalar: **Alt+Shift+L** (Kirill → Lotin), **Alt+Shift+K** (Lotin → Kirill).
Excel formulalari: `=LOTINGA(A1)`, `=KIRILGA(A1)`.

### 1.3.0 dagi o'zgarishlar — Office'ning barcha versiyalari
* **Excel 2003/2007** da qo'shimcha umuman ishlamasligi tuzatildi (Excel 2010 dan paydo bo'lgan `Range.CountLarge` ishlatilgan edi).
* **Word 2003/2007** da qo'shimcha ishlamasligi tuzatildi (Word 2010 dan paydo bo'lgan `UndoRecord` ishlatilgan edi).
* **Windows 7 (PowerShell 2.0)** da o'rnatuvchining zaxira usuli ishlamasligi tuzatildi.
* **Rus va boshqa tildagi Office**: Word STARTUP papkasi endi Word'ning o'zidan so'raladi.
* **Office 2003** qo'llab-quvvatlanadi: `.dot` / `.xla`, tugmalar asboblar panelida.
* Yuklab olingan fayllardagi «Internetdan olingan» belgisi olib tashlanadi — yangi Office makroslarni bloklamaydi.
* O'rnatish jurnali: `%TEMP%\KirillLotin-install.log`.

### 1.2.0 dagi o'zgarishlar
* Word yoki Excel ochiq bo'lsa, o'rnatuvchi (va o'chiruvchi) ochiq Office dasturlari ro'yxatini ko'rsatadi va **«Majburan yopish»** tugmasini taklif qiladi: u barcha ochiq Office dasturlarini (Word, Excel, PowerPoint, Outlook va boshqalar) darhol yopadi. Saqlanmagan o'zgarishlar yo'qoladi, shuning uchun avval ishingizni saqlang.
* **«Qayta tekshirish»** tugmasi — dasturlarni o'zingiz yopganingizdan keyin davom etish uchun.

### 1.1.0 dagi o'zgarishlar
* O'rnatish ancha tezlashdi: qo'shimchalar oldindan tayyorlangan, o'rnatuvchi ularni faqat nusxalaydi va bir marta tekshiradi.
* «Tezkor tugmalarni o'rnatib bo'lmadi» xatosi tuzatildi: Word tugmalari endi Word ichida avtomatik o'rnatiladi.
* Office javob bermay qolsa, o'rnatuvchi osilib qolmaydi (vaqt chegarasi).

`KirillLotin.dotm` va `KirillLotin.xlam` — qo'lda o'rnatish uchun (Word STARTUP va Excel XLSTART papkalariga nusxalang).
