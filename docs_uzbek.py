from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import os

DOCS_DIR = "docs"
os.makedirs(DOCS_DIR, exist_ok=True)

def set_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '1B3A6B')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def blue_cell(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1B3A6B')
    tcPr.append(shd)

def light_cell(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'E8EEF7')
    tcPr.append(shd)

def add_header(doc, title, qonun=""):
    # Ko'k sarlavha
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.cell(0, 0)
    blue_cell(cell)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("O'ZBEKISTON RESPUBLIKASI")
    r.font.color.rgb = RGBColor(255, 255, 255)
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.name = 'Times New Roman'

    doc.add_paragraph()

    # Asosiy sarlavha
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(title)
    tr.font.size = Pt(16)
    tr.font.bold = True
    tr.font.name = 'Times New Roman'
    tr.font.color.rgb = RGBColor(27, 58, 107)

    if qonun:
        qp = doc.add_paragraph()
        qp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        qr = qp.add_run(f"({qonun})")
        qr.font.size = Pt(9)
        qr.font.italic = True
        qr.font.name = 'Times New Roman'
        qr.font.color.rgb = RGBColor(100, 100, 100)

    sep = doc.add_paragraph()
    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sep.add_run("─" * 60)
    sr.font.color.rgb = RGBColor(27, 58, 107)
    doc.add_paragraph()

def add_table(doc, rows):
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_border(table)
    for i, (label, value) in enumerate(rows):
        lc = table.cell(i, 0)
        light_cell(lc)
        lp = lc.paragraphs[0]
        lr = lp.add_run(f"  {label}")
        lr.font.bold = True
        lr.font.size = Pt(11)
        lr.font.name = 'Times New Roman'
        lr.font.color.rgb = RGBColor(27, 58, 107)
        vc = table.cell(i, 1)
        vp = vc.paragraphs[0]
        vr = vp.add_run(f"  {value}")
        vr.font.size = Pt(11)
        vr.font.name = 'Times New Roman'

def add_text(doc, text, bold=False, size=11, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = 'Times New Roman'
    if color:
        r.font.color.rgb = color
    return p

def add_signature(doc, left, right, left_name, right_name):
    doc.add_paragraph()
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run("TOMONLARNING IMZOLARI")
    sr.font.bold = True
    sr.font.size = Pt(12)
    sr.font.color.rgb = RGBColor(27, 58, 107)
    sr.font.name = 'Times New Roman'
    doc.add_paragraph()

    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sana = datetime.now().strftime('%d.%m.%Y')

    headers = [left, right]
    names = [left_name, right_name]

    for col, h in enumerate(headers):
        cell = table.cell(0, col)
        blue_cell(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(11)
        r.font.name = 'Times New Roman'

    for col, name in enumerate(names):
        cell = table.cell(1, col)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(name)
        r.font.size = Pt(11)
        r.font.name = 'Times New Roman'

    labels = ["Manzil:", "Imzo:", "Sana:"]
    values = ["___________________", "___________________", sana]
    for i, (label, value) in enumerate(zip(labels, values)):
        for col in range(2):
            cell = table.cell(i+2, col)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"{label} {value}")
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'

    set_border(table)

def add_footer(doc):
    doc.add_paragraph()
    sep = doc.add_paragraph()
    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sep.add_run("─" * 60)
    sr.font.color.rgb = RGBColor(27, 58, 107)
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("© Yuridik Yordamchi Bot | @YuridikYordamchi")
    fr.font.size = Pt(8)
    fr.font.italic = True
    fr.font.color.rgb = RGBColor(150, 150, 150)
    fr.font.name = 'Times New Roman'

def make_doc(doc_type, data, order_id):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    sana = datetime.now().strftime('%d.%m.%Y')

    # ── IJARA ─────────────────────────────────────────────────────────────────
    if doc_type == "ijara":
        add_header(doc, "KO'CHMAS MULKNI IJARAGA BERISH SHARTNOMASI",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 535-572-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"YY-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Ijara beruvchi:", data.get('beruvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('beruvchi_pasport', '___')),
            ("Ijara oluvchi:", data.get('oluvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('oluvchi_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. SHARTNOMA PREDMETI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Mulk manzili:", data.get('manzil', '___')),
            ("Ijara muddati:", f"{data.get('muddat', '___')} oy"),
            ("Boshlanish sanasi:", sana),
        ])
        doc.add_paragraph()

        add_text(doc, "3. IJARA HAQI VA TO'LOV TARTIBI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Oylik ijara haqi:", f"{data.get('narx', '___')} so'm"),
            ("To'lov muddati:", "Har oyning 5-kunigacha"),
            ("To'lov usuli:", "Naqd yoki bank o'tkazmasi"),
        ])
        doc.add_paragraph()

        add_text(doc, "4. TOMONLARNING HUQUQ VA MAJBURIYATLARI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "4.1. Ijara beruvchi mulkni yaxshi holatda topshirishga majbur.")
        add_text(doc, "4.2. Ijara oluvchi mulkni maqsadli foydalanishga va belgilangan muddat ichida ijaraga qaytarishga majbur.")
        add_text(doc, "4.3. Kommunal to'lovlar ijara oluvchi tomonidan to'lanadi.")
        doc.add_paragraph()

        add_text(doc, "5. NIZOLARNI HAL ETISH TARTIBI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, f"5.1. Nizolar O'zbekiston Respublikasi qonunchiligi asosida {data.get('shahar', '___')} shahri sudi orqali hal etiladi.")
        doc.add_paragraph()

        add_text(doc, "6. SHARTNOMANING AMAL QILISH MUDDATI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, f"6.1. Shartnoma imzolangan kundan boshlab {data.get('muddat', '___')} oy davomida amal qiladi.")

        add_signature(doc, "IJARA BERUVCHI", "IJARA OLUVCHI",
                     data.get('beruvchi_fio', '___'), data.get('oluvchi_fio', '___'))

    # ── OLDI-SOTDI ────────────────────────────────────────────────────────────
    elif doc_type == "oldi_sotdi":
        add_header(doc, "OLDI-SOTDI SHARTNOMASI",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 386-425-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"OS-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Sotuvchi:", data.get('sotuvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('sotuvchi_pasport', '___')),
            ("Xaridor:", data.get('xaridor_fio', '___')),
            ("Pasport seriyasi:", data.get('xaridor_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. SHARTNOMA PREDMETI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Mulk tavsifi:", data.get('tovar', '___')),
            ("Mulk narxi:", f"{data.get('narx', '___')} so'm"),
            ("To'lov muddati:", "Shartnoma imzolanish kuni"),
        ])
        doc.add_paragraph()

        add_text(doc, "3. TOMONLARNING MAJBURIYATLARI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Sotuvchi mulkni belgilangan narxda va yaxshi holatda topshirishga majbur.")
        add_text(doc, "3.2. Xaridor mulk narxini to'liq to'lashga majbur.")
        add_text(doc, "3.3. Mulk huquqi xaridorga to'lov amalga oshirilgandan so'ng o'tadi.")
        doc.add_paragraph()

        add_text(doc, "4. KAFOLAT", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "4.1. Sotuvchi mulk uchinchi shaxslarning da'vosidan xoli ekanligini kafolatlaydi.")
        doc.add_paragraph()

        add_text(doc, "5. NIZOLARNI HAL ETISH", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, f"5.1. Nizolar {data.get('shahar', '___')} shahri sudi orqali hal etiladi.")

        add_signature(doc, "SOTUVCHI", "XARIDOR",
                     data.get('sotuvchi_fio', '___'), data.get('xaridor_fio', '___'))

    # ── TILXAT ────────────────────────────────────────────────────────────────
    elif doc_type == "tilxat":
        add_header(doc, "TILXAT",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 732-747-moddalari asosida")

        add_table(doc, [
            ("Yozilgan sana:", sana),
            ("Yozilgan joy:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "Men,", bold=False)
        add_table(doc, [
            ("To'liq ismi:", data.get('fio', '___')),
            ("Pasport seriyasi:", data.get('pasport', '___')),
            ("Yashash manzili:", data.get('manzil', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "quyidagini tasdiqlayman:", bold=False)
        doc.add_paragraph()
        tp = doc.add_paragraph()
        tr = tp.add_run(data.get('mazmun', '___'))
        tr.font.size = Pt(12)
        tr.font.name = 'Times New Roman'
        tr.font.bold = True
        doc.add_paragraph()

        add_table(doc, [
            ("Miqdori:", f"{data.get('miqdor', '___')} so'm"),
            ("Sana:", sana),
        ])
        doc.add_paragraph()

        add_text(doc, "Ushbu tilxat O'zbekiston Respublikasi Fuqarolik kodeksi talablariga muvofiq tuzildi va huquqiy kuchga ega.", size=10)
        doc.add_paragraph()

        ip = doc.add_paragraph()
        ip.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ir = ip.add_run(f"Imzo: ___________________\nIsmi: {data.get('fio', '___')}\nSana: {sana}")
        ir.font.size = Pt(11)
        ir.font.name = 'Times New Roman'

    # ── ISHONCHNOMA ───────────────────────────────────────────────────────────
    elif doc_type == "ishonchnoma":
        add_header(doc, "ISHONCHNOMA",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 121-129-moddalari asosida")

        add_table(doc, [
            ("Raqam:", f"ISH-{order_id}/{datetime.now().year}"),
            ("Sana:", sana),
            ("Joy:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "Men,", bold=False)
        add_table(doc, [
            ("Ishonchnoma beruvchi:", data.get('beruvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('beruvchi_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "quyidagi shaxsga ishonchnoma beraman:", bold=False)
        add_table(doc, [
            ("Ishonchnoma oluvchi:", data.get('oluvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('oluvchi_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "VAKOLATLAR:", bold=True, color=RGBColor(27, 58, 107))
        vp = doc.add_paragraph()
        vr = vp.add_run(data.get('vakolat', '___'))
        vr.font.size = Pt(11)
        vr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_table(doc, [
            ("Amal qilish muddati:", data.get('muddat', '___')),
            ("Boshlanish sanasi:", sana),
        ])
        doc.add_paragraph()

        add_text(doc, "Ishonchnoma qayta topshirish huquqisiz berilgan.", size=10)

        add_signature(doc, "ISHONCHNOMA BERUVCHI", "ISHONCHNOMA OLUVCHI",
                     data.get('beruvchi_fio', '___'), data.get('oluvchi_fio', '___'))

    # ── NIKOH SHARTNOMASI ─────────────────────────────────────────────────────
    elif doc_type == "nikoh":
        add_header(doc, "NIKOH SHARTNOMASI",
                   "O'zbekiston Respublikasi Oila kodeksining 27-31-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"NK-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Er:", data.get('er_fio', '___')),
            ("Pasport seriyasi:", data.get('er_pasport', '___')),
            ("Xotin:", data.get('xotin_fio', '___')),
            ("Pasport seriyasi:", data.get('xotin_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. MULKIY SHARTLAR", bold=True, color=RGBColor(27, 58, 107))
        mp = doc.add_paragraph()
        mr = mp.add_run(data.get('shartlar', '___'))
        mr.font.size = Pt(11)
        mr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_text(doc, "3. UMUMIY QOIDALAR", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Nikoh davomida orttirilgan mulk qonun bo'yicha umumiy mulk hisoblanadi.")
        add_text(doc, "3.2. Ushbu shartnomada ko'rsatilgan mulkiy shartlar O'zbekiston Respublikasi Oila kodeksiga zid bo'lmasligi shart.")
        add_text(doc, "3.3. Shartnoma notarius tomonidan tasdiqlanganidan keyin kuchga kiradi.")

        add_signature(doc, "ER", "XOTIN",
                     data.get('er_fio', '___'), data.get('xotin_fio', '___'))

    # ── QARZ SHARTNOMASI ──────────────────────────────────────────────────────
    elif doc_type == "qarz":
        add_header(doc, "QARZ SHARTNOMASI",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 732-747-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"QR-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Qarz beruvchi:", data.get('beruvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('beruvchi_pasport', '___')),
            ("Qarz oluvchi:", data.get('oluvchi_fio', '___')),
            ("Pasport seriyasi:", data.get('oluvchi_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. QARZ SHARTLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Qarz miqdori:", f"{data.get('miqdor', '___')} so'm"),
            ("Foiz stavkasi:", f"{data.get('foiz', '0')}% yillik"),
            ("Qaytarish muddati:", data.get('muddat', '___')),
            ("Qaytarish sanasi:", "___________________"),
        ])
        doc.add_paragraph()

        add_text(doc, "3. TOMONLARNING MAJBURIYATLARI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Qarz beruvchi belgilangan miqdorni shartnoma imzolanishi bilan beradi.")
        add_text(doc, "3.2. Qarz oluvchi qarzni belgilangan muddat ichida to'liq qaytarishga majbur.")
        add_text(doc, "3.3. Kechikish holatida qarz oluvchi kunlik 0.1% jarima to'laydi.")
        doc.add_paragraph()

        add_text(doc, "4. NIZOLARNI HAL ETISH", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, f"4.1. Nizolar {data.get('shahar', '___')} shahri sudi orqali hal etiladi.")

        add_signature(doc, "QARZ BERUVCHI", "QARZ OLUVCHI",
                     data.get('beruvchi_fio', '___'), data.get('oluvchi_fio', '___'))

    # ── MEHNAT SHARTNOMASI ────────────────────────────────────────────────────
    elif doc_type == "mehnat":
        add_header(doc, "MEHNAT SHARTNOMASI",
                   "O'zbekiston Respublikasi Mehnat kodeksining 73-100-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"MH-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Ish beruvchi:", data.get('ish_beruvchi', '___')),
            ("Xodim:", data.get('xodim_fio', '___')),
            ("Pasport seriyasi:", data.get('xodim_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. ISH SHAROITLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Lavozim:", data.get('lavozim', '___')),
            ("Oylik maosh:", f"{data.get('oylik', '___')} so'm"),
            ("Shartnoma muddati:", data.get('muddat', '___')),
            ("Ish vaqti:", "09:00 - 18:00 (Dushanba - Juma)"),
            ("Tanaffus:", "13:00 - 14:00"),
            ("Dam olish kunlari:", "Shanba, Yakshanba"),
            ("Yillik ta'til:", "15 ish kuni (MK 134-modda)"),
        ])
        doc.add_paragraph()

        add_text(doc, "3. TOMONLARNING MAJBURIYATLARI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Ish beruvchi maoshni har oyning 5-kunigacha to'lashga majbur.")
        add_text(doc, "3.2. Xodim ish joyida intizomga rioya qilishga majbur.")
        add_text(doc, "3.3. Shartnomani bekor qilishda 2 hafta oldindan ogohlantirish shart.")

        add_signature(doc, "ISH BERUVCHI", "XODIM",
                     data.get('ish_beruvchi', '___'), data.get('xodim_fio', '___'))

    # ── PUDRAT SHARTNOMASI ────────────────────────────────────────────────────
    elif doc_type == "pudrat":
        add_header(doc, "PUDRAT SHARTNOMASI",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 631-665-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"PD-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Buyurtmachi:", data.get('buyurtmachi', '___')),
            ("Pudratchi:", data.get('pudratchi', '___')),
            ("Pasport/STIR:", data.get('pudratchi_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. ISH PREDMETI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Bajariladigan ish:", data.get('ish_tavsifi', '___')),
            ("Ish narxi:", f"{data.get('narx', '___')} so'm"),
            ("Bajarish muddati:", data.get('muddat', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "3. TO'LOV TARTIBI", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Avans to'lov: Shartnoma imzolanishi bilan 30%.")
        add_text(doc, "3.2. Qolgan 70% ish topshirilgandan so'ng to'lanadi.")
        doc.add_paragraph()

        add_text(doc, "4. KAFOLAT", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "4.1. Pudratchi bajarilgan ish sifatiga 12 oy kafolat beradi.")

        add_signature(doc, "BUYURTMACHI", "PUDRATCHI",
                     data.get('buyurtmachi', '___'), data.get('pudratchi', '___'))

    # ── ALIMENT ───────────────────────────────────────────────────────────────
    elif doc_type == "aliment":
        add_header(doc, "ALIMENT TO'LASH TO'G'RISIDA SHARTNOMA",
                   "O'zbekiston Respublikasi Oila kodeksining 99-113-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"AL-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Aliment to'lovchi:", data.get('tolovchi', '___')),
            ("Pasport seriyasi:", data.get('tolovchi_pasport', '___')),
            ("Aliment oluvchi:", data.get('oluvchi', '___')),
            ("Farzand(lar):", data.get('bola_ismi', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. ALIMENT MIQDORI VA TO'LOV TARTIBI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Oylik aliment miqdori:", f"{data.get('miqdor', '___')} so'm"),
            ("To'lov muddati:", "Har oyning 1-kunigacha"),
            ("To'lov usuli:", "Naqd yoki bank o'tkazmasi"),
            ("Amal qilish muddati:", data.get('muddat', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "3. UMUMIY QOIDALAR", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "3.1. Shartnoma notarius tomonidan tasdiqlanishi kerak (OK 101-modda).")
        add_text(doc, "3.2. Kechikish holatida kunlik 0.1% ustama to'lanadi.")

        add_signature(doc, "ALIMENT TO'LOVCHI", "ALIMENT OLUVCHI",
                     data.get('tolovchi', '___'), data.get('oluvchi', '___'))

    # ── HAMKORLIK ─────────────────────────────────────────────────────────────
    elif doc_type == "hamkorlik":
        add_header(doc, "HAMKORLIK (SHERIKLIK) SHARTNOMASI",
                   "O'zbekiston Respublikasi Fuqarolik kodeksining 781-796-moddalari asosida")

        add_table(doc, [
            ("Shartnoma raqami:", f"HK-{order_id}/{datetime.now().year}"),
            ("Tuzilgan sanasi:", sana),
            ("Tuzilgan joyi:", data.get('shahar', '___') + " shahri"),
        ])
        doc.add_paragraph()

        add_text(doc, "1. SHARTNOMA TOMONLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("1-tomon:", data.get('tomon1', '___')),
            ("Pasport/STIR:", data.get('tomon1_pasport', '___')),
            ("2-tomon:", data.get('tomon2', '___')),
            ("Pasport/STIR:", data.get('tomon2_pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. HAMKORLIK MAQSADI", bold=True, color=RGBColor(27, 58, 107))
        lp = doc.add_paragraph()
        lr = lp.add_run(data.get('loyiha', '___'))
        lr.font.size = Pt(11)
        lr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_text(doc, "3. FOYDA TAQSIMOTI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Foyda taqsimoti:", data.get('foyda_taqsim', '___')),
            ("Hamkorlik muddati:", data.get('muddat', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "4. MAXFIYLIK", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "4.1. Tomonlar hamkorlik doirasidagi ma'lumotlarni uchinchi shaxslarga oshkor etmaydi.")

        add_signature(doc, "1-TOMON", "2-TOMON",
                     data.get('tomon1', '___'), data.get('tomon2', '___'))

    # ── DA'VO ARIZASI ─────────────────────────────────────────────────────────
    elif doc_type == "davo":
        add_header(doc, "DA'VO ARIZASI",
                   "O'zbekiston Respublikasi Fuqarolik protsessual kodeksi asosida")

        jp = doc.add_paragraph()
        jp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        jr = jp.add_run(f"{data.get('shahar', '___')} shahri sudi raisiga\nAriza beruvchi: {data.get('ariza_beruvchi', '___')}\nSana: {sana}")
        jr.font.size = Pt(11)
        jr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_text(doc, "DA'VO ARIZASI", bold=True, size=14, color=RGBColor(27, 58, 107))
        doc.add_paragraph()

        add_text(doc, "1. ARIZA BERUVCHI MA'LUMOTLARI", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("To'liq ismi:", data.get('ariza_beruvchi', '___')),
            ("Pasport seriyasi:", data.get('pasport', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "2. JAVOBGAR", bold=True, color=RGBColor(27, 58, 107))
        add_table(doc, [
            ("Javobgar:", data.get('javobgar', '___')),
        ])
        doc.add_paragraph()

        add_text(doc, "3. HOLAT BAYONI", bold=True, color=RGBColor(27, 58, 107))
        mp = doc.add_paragraph()
        mr = mp.add_run(data.get('muammo', '___'))
        mr.font.size = Pt(11)
        mr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_text(doc, "4. TALABLAR", bold=True, color=RGBColor(27, 58, 107))
        tp = doc.add_paragraph()
        tr = tp.add_run(data.get('talab', '___'))
        tr.font.size = Pt(11)
        tr.font.name = 'Times New Roman'
        doc.add_paragraph()

        add_text(doc, "5. ILOVA", bold=True, color=RGBColor(27, 58, 107))
        add_text(doc, "— Ushbu ariza nusxasi\n— Pasport nusxasi\n— Boshqa hujjatlar")
        doc.add_paragraph()

        ip = doc.add_paragraph()
        ip.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ir = ip.add_run(f"Ariza beruvchi: ___________________\nIsmi: {data.get('ariza_beruvchi', '___')}\nSana: {sana}")
        ir.font.size = Pt(11)
        ir.font.name = 'Times New Roman'

    add_footer(doc)
    path = f"{DOCS_DIR}/{doc_type}_{order_id}.docx"
    doc.save(path)
    return path

print("Tayyor!")
