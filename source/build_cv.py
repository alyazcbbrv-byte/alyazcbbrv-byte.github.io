"""
Build the CV in Azerbaijani and English as HTML (for PDF printing) and Word (.docx) from one data set.

    python source/build_cv.py        # writes source/cv.az.html, source/cv.en.html, source/CV_..._AZ.docx, ..._EN.docx
PDF: print the HTML with headless Edge/Chrome (see README note in the chat).
"""
from __future__ import annotations

import html
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

OUT = Path(__file__).resolve().parent
NAVY, TEAL = RGBColor(0x14, 0x21, 0x3D), RGBColor(0x2A, 0x9D, 0x8F)

FULL_EXTRA = ["055-360-46-08", "Bakı / Baku, Yeni Yasamal"]
CONTACT = ["Baku, Azerbaijan", "alyazcbbrv@gmail.com", "linkedin.com/in/alyaz-cabbarova-450a6140b", "github.com/alyazcbbrv-byte", "alyazcbbrv-byte.github.io", "youtube.com/@memm_mova"]

DATA = {
    "az": {
        "lang": "az",
        "file": "CV_Alyaz_Cabbarova_AZ",
        "name": "Alyaz Cabbarova",
        "role": "Data analitik",
        "contact": ["Bakı, Azərbaycan"] + CONTACT[1:],
        "h": {"summary": "Xülasə", "skills": "Bacarıqlar", "edu": "Təhsil və kurslar", "exp": "Təcrübə", "proj": "Layihələr"},
        "summary": "Riyaziyyat və informatika təhsilindən data analitikaya keçirəm. Analitik və məntiqi düşüncəni təhsildən, texniki bacarıqları (Power BI, DAX, SQL, Python, Excel) isə kursda və GitHub-dakı 8 layihədə qazanmışam. Xam biznes datasını təmizləyir, modelləşdirir, ölçür və dashboard-larla nəticəyə çevirirəm. Power BI və DAX, SQL, Python (pandas) və Excel ilə işləyirəm. Datanı istifadə etməzdən əvvəl yoxlayıram: məsələn, mağaza layihəmdə planın əslində satışın 29 gün gecikmiş nüsxəsi olduğunu aşkar etdim. Hər layihəmi tapıntı, tövsiyə və məhdudiyyətlərlə sənədləşdirirəm.",
        "skills": [
            ("BI və dashboard", "Power BI, DAX (zaman intellekti, kohort və churn ölçüləri), Power Query (M), ulduz sxemi, PBIP/TMDL"),
            ("SQL", "Join, subquery, CTE (rekursiv daxil), window funksiyaları, sxem dizaynı, bütövlük yoxlamaları"),
            ("Python", "pandas, NumPy, matplotlib, seaborn, SciPy, statsmodels, Jupyter"),
            ("Excel", "SUMIFS, INDEX/MATCH, data validation, nativ qrafiklər, düsturla işləyən dashboard"),
            ("Statistika", "Hipotez testləri (chi-square, Kruskal-Wallis, Mann-Whitney, Spearman), bootstrap, RFM, Pareto, proqnoz və backtest"),
            ("Proqramlar", "Word, PowerPoint, Excel, Teams, Zoom, Canva, GeoGebra, Desmos"),
            ("Dillər", "Azərbaycan dili, İngilis dili"),
            ("Digər", "Git/GitHub, data keyfiyyəti yoxlamaları, ETL, iki dildə sənədləşdirmə (Azərbaycan, ingilis)"),
        ],
        "edu": [("Azərbaycan Dövlət Pedaqoji Universiteti (ADPU), Şəki filialı", "2020-2025", ["Riyaziyyat-İnformatika müəllimliyi"]), ("Data analitika kursu", "4 ay, davam edir · məşğulluq xidmətinin dəstəyi ilə",
                 ["Mövzular: Excel, Power BI və DAX, SQL, Python, statistika (Pareto, ABC/XYZ, RFM, kohort analizi, proqnoz).",
                  "Kurs tapşırıqlarını GitHub-da tam layihələrə çevirmişəm (aşağıda)."])],
        "exp": [
            ("Təhsil İdarəsi", "2024 mart-may · Təhsil Tələbə Kredit Fondunun təcrübə proqramı", [
                "Sənədlər və hesabatlarla işlədim; monitorinqlərin təşkilində iştirak etdim.",
                "Müəllimlərin sertifikasiya imtahanlarının təşkili prosesində iştirak etdim.",
                "Məlumatların təhlili və doldurulması: Excel və digər platformalarda."]),
            ("Şəki şəhər Fizika, Riyaziyyat və Humanitar Fənlər Təmayüllü Lisey", "2025 may-iyun · təcrübə", [
                "5-ci sinifdə dərs demişəm; 5-9-cu siniflərə onlayn dərslər keçmişəm."]),
            ("SualCavab.com platforması", "ortaq müəllim · 5 ay", [
                "Riyaziyyat və informatika üzrə buraxılış, qəbul və sinif sualları hazırlayıram; bu günə qədər təxminən 2,000-ə yaxın sual hazırlamışam."]),
            ("YouTube kanalı «Math1era»", "müstəqil layihə · youtube.com/@memm_mova", [
                "5-ci sinif riyaziyyatı üzrə 30 video dərs hazırlayıb yerləşdirmişəm: natural ədədlər, adi və onluq kəsrlər, faiz.",
                "48 abunəçi; ən çox izlənən video 549 baxışdır."])],
        "proj": [
            ("Mağaza şəbəkəsi: satış, plan və churn analizi", "Power BI, DAX, Python", [
                "274 mağaza, 50 region, 4 illik gündəlik data (113,829 sətir); ulduz sxemi, 25 DAX ölçüsü, 3 səhifəli dashboard.",
                "Planın hər mağaza və gün üçün satışın 29 gün gecikmiş nüsxəsi olduğunu aşkar etdim: \"plan icrası\" real hədəfə çatmanı yox, aylıq dəyişməni göstərirdi.",
                "Artımın mənbəyi mağaza başına satışdır (+144%); mağazaların 66%-i bağlanıb, median ömür 500 gündür."]),
            ("Toyota satış dashboard-u", "Power BI, DAX", [
                "672,765 əməliyyat, 157 ölkə: satış xəritəsi, stok və Avropa analizi (3 səhifə).",
                "Keyfiyyət yoxlamasında yanlış 54% YoY rəqəmini (düzgünü 34.8%), ay sıralamasını və təkrarlanan filtrləri düzəltdim."]),
            ("Amazon satış analizi", "Python, pandas, Power BI", [
                "672 min əməliyyat: ETL, 22% illik artım (CAGR); artım əməliyyat sayından gəlir, üç kateqoriya satışın 69%-ni verir."]),
            ("SQL ilə HR analitikası", "SQL (SQLite), Python", [
                "7 cədvəlli HR bazası, 32 biznes sorğusu: join, CTE, rekursiv CTE, window funksiyaları, 12 bütövlük yoxlaması.",
                "Satış şöbəsi maaş fondunun 44.5%-dir; compa-ratio ilə aşağı ödənilən vəzifələr və maaş sıxılması aşkarlandı."]),
            ("Superstore mənfəət analizi", "Python, pandas, SciPy", [
                "9,994 sətir: 20%-dən çox endirim 12% marjanı zərərə çevirir; sətirlərin 14%-i 135 min $ itirir, zərərli ştatlar çox endirim edilən ştatlardır."]),
            ("RFM seqmentasiyası və satış proqnozu", "Python, statsmodels, SciPy", [
                "850 müştəri: Champions və Loyal müştərilərin 46%-i, gəlirin 67%-ni verir; Pareto, kanal müqayisəsi (bootstrap), 3 üsullu proqnoz backtest-i (MAPE 12.5%) və kənar qiymətin təsiri."]),
            ("Tələbə nəticələri: kəşfiyyat data analizi", "Python, pandas, seaborn, SciPy", [
                "1,000 nəticə, 244 tələbə: F yalnız 60% və aşağı davamiyyətdə çıxır; ixtisaslar arasında statistik fərq yoxdur; sadə qayda tələbələrin 12%-ni risk qrupu kimi tapır."]),
            ("Restoran satışı: interaktiv Excel dashboard", "Excel", [
                "Makrosuz, yalnız SUMIFS düsturları ilə iki açılan filtr, KPI kartları və 5 qrafik; düsturlar ayrıca hesablama ilə yoxlanılıb."]),
        ],
        "note": "Layihələrdə istifadə olunan datalar kurs nümunə datalarıdır. Bütün layihələr və kodlar: github.com/alyazcbbrv-byte",
    },
    "en": {
        "lang": "en",
        "file": "CV_Alyaz_Cabbarova_EN",
        "name": "Alyaz Cabbarova",
        "role": "Data Analyst",
        "contact": CONTACT,
        "h": {"summary": "Summary", "skills": "Skills", "edu": "Education and courses", "exp": "Experience", "proj": "Projects"},
        "summary": "Moving from mathematics and informatics education into data analytics. I built my analytical thinking through my degree and my technical skills (Power BI, DAX, SQL, Python, Excel) through a course and 8 projects on GitHub. I clean, model and measure raw business data and turn it into decisions with dashboards, using Power BI and DAX, SQL, Python (pandas) and Excel. I check data before trusting it: in my retail project I found that the plan was an exact copy of sales lagged by 29 days. Every project is documented with findings, recommendations and limitations.",
        "skills": [
            ("BI and dashboards", "Power BI, DAX (time intelligence, cohort and churn measures), Power Query (M), star schema, PBIP/TMDL"),
            ("SQL", "Joins, subqueries, CTEs (incl. recursive), window functions, schema design, integrity checks"),
            ("Python", "pandas, NumPy, matplotlib, seaborn, SciPy, statsmodels, Jupyter"),
            ("Excel", "SUMIFS, INDEX/MATCH, data validation, native charts, formula-driven dashboards"),
            ("Statistics", "Hypothesis tests (chi-square, Kruskal-Wallis, Mann-Whitney, Spearman), bootstrap, RFM, Pareto, forecasting and backtesting"),
            ("Tools", "Word, PowerPoint, Excel, Teams, Zoom, Canva, GeoGebra, Desmos"),
            ("Languages", "Azerbaijani, English"),
            ("Other", "Git/GitHub, data-quality checks, ETL, bilingual documentation (Azerbaijani, English)"),
        ],
        "edu": [("Azerbaijan State Pedagogical University (ADPU), Sheki branch", "2020-2025", ["Teacher of Mathematics and Informatics"]), ("Data analytics course", "4 months, ongoing · supported by an employment service programme",
                 ["Topics: Excel, Power BI and DAX, SQL, Python, statistics (Pareto, ABC/XYZ, RFM, cohort analysis, forecasting).",
                  "I turned the course assignments into complete projects on GitHub (below)."])],
        "exp": [
            ("Education Department", "March-May 2024 · internship, Student Loan Fund programme", [
                "Worked with documents and reports; took part in organising monitoring activities.",
                "Took part in organising the teacher certification exams.",
                "Analysed data and entered it in Excel and other platforms."]),
            ("Sheki City Lyceum of Physics, Mathematics and Humanities", "May-June 2025 · internship", [
                "Taught a 5th-grade class; gave online lessons to grades 5-9."]),
            ("SualCavab.com platform", "partner teacher · 5 months", [
                "I prepare graduation, admission and class questions in mathematics and informatics; I have written about 2,000 questions so far."]),
            ("YouTube channel \"Math1era\"", "independent project · youtube.com/@memm_mova", [
                "Created and published 30 video lessons for 5th-grade mathematics: natural numbers, common and decimal fractions, percentages.",
                "48 subscribers; the most-watched video has 549 views."])],
        "proj": [
            ("Retail chain: sales, plan and churn analysis", "Power BI, DAX, Python", [
                "274 stores, 50 regions, 4 years of daily data (113,829 rows); star schema, 25 DAX measures, 3-page dashboard.",
                "Found that the plan is an exact copy of each store's sales lagged by 29 days, so \"plan achievement\" was really month-over-month change, not progress to a target.",
                "Growth comes from sales per store (+144%); 66% of stores have closed, median lifetime 500 days."]),
            ("Toyota sales dashboard", "Power BI, DAX", [
                "672,765 transactions, 157 countries: sales map, stock and Europe analysis (3 pages).",
                "In a quality pass I fixed a wrong 54% YoY figure (correct: 34.8%), month sorting and duplicated filters."]),
            ("Amazon sales analysis", "Python, pandas, Power BI", [
                "672k transactions: ETL, 22% CAGR; growth is driven by transaction count, three categories make up 69% of sales."]),
            ("HR analytics with SQL", "SQL (SQLite), Python", [
                "7-table HR database, 32 business queries: joins, CTEs, recursive CTE, window functions, 12 integrity checks.",
                "Sales is 44.5% of payroll; compa-ratio exposed under-paid jobs and salary compression."]),
            ("Superstore profitability analysis", "Python, pandas, SciPy", [
                "9,994 order lines: discounts above 20% turn a 12% margin into losses; 14% of lines lose $135k and loss-making states are the heavy-discount states."]),
            ("RFM segmentation and sales forecast", "Python, statsmodels, SciPy", [
                "850 customers: Champions and Loyal are 46% of customers and 67% of revenue; Pareto, channel comparison (bootstrap), 3-method forecast backtest (MAPE 12.5%) and the effect of an outlier."]),
            ("Student performance: exploratory data analysis", "Python, pandas, seaborn, SciPy", [
                "1,000 results, 244 students: an F appears only at 60% attendance or below; no significant difference between majors; a simple rule flags 12% of students at risk."]),
            ("Restaurant sales: interactive Excel dashboard", "Excel", [
                "No macros, formulas only (SUMIFS): two dropdown filters, KPI cards and 5 charts; formulas independently recalculated and checked."]),
        ],
        "note": "Data used in the projects is course sample data. All projects and code: github.com/alyazcbbrv-byte",
    },
}

CSS = """
@page { size: A4; margin: 11mm 14mm; }
* { box-sizing: border-box; }
body { font-family: "Segoe UI", Calibri, Arial, sans-serif; color: #1f2933; font-size: 9.5pt; line-height: 1.36; margin: 0; }
header { border-bottom: 3px solid #2a9d8f; padding-bottom: 7px; margin-bottom: 8px; }
h1 { margin: 0; font-size: 25pt; color: #14213d; }
.role { font-size: 12.5pt; color: #2a9d8f; font-weight: 600; margin-top: 1px; }
.contact { margin-top: 5px; font-size: 9pt; color: #3d4a5c; }
.contact span { margin-right: 12px; white-space: nowrap; }
h2 { font-size: 10.5pt; color: #14213d; text-transform: uppercase; letter-spacing: 1.2px; border-bottom: 1px solid #d7dde5; padding-bottom: 2px; margin: 12px 0 6px; }
.skills { display: grid; grid-template-columns: 125px 1fr; gap: 2px 10px; }
.skills b { color: #14213d; }
.item { margin-bottom: 5px; page-break-inside: avoid; }
.item .t { font-weight: 700; color: #14213d; }
.item .s { color: #6b7785; font-size: 9pt; font-weight: 400; }
ul { margin: 2px 0 0 15px; padding: 0; }
.note { color: #6b7785; font-size: 8.5pt; margin-top: 8px; }
"""


def contact_of(d: dict, full: bool) -> list:
    c = [x for x in d["contact"]]
    if full:
        c = [c[1], FULL_EXTRA[0], FULL_EXTRA[1]] + c[2:]
    return c


def make_html(d: dict, full: bool = False) -> str:
    e = html.escape
    out = [f'<!doctype html><html lang="{d["lang"]}"><head><meta charset="utf-8"><title>{e(d["name"])}: CV</title><style>{CSS}</style></head><body>']
    out.append(f'<header><h1>{e(d["name"])}</h1><div class="role">{e(d["role"])}</div><div class="contact">' + "".join(f"<span>{e(c)}</span>" for c in contact_of(d, full)) + "</div></header>")
    out.append(f'<h2>{d["h"]["summary"]}</h2><p>{e(d["summary"])}</p>')
    out.append(f'<h2>{d["h"]["skills"]}</h2><div class="skills">' + "".join(f"<b>{e(a)}</b><span>{e(b)}</span>" for a, b in d["skills"]) + "</div>")
    out.append(f'<h2>{d["h"]["edu"]}</h2>')
    for t, s, bl in d["edu"]:
        out.append(f'<div class="item"><div class="t">{e(t)} <span class="s">· {e(s)}</span></div><ul>' + "".join(f"<li>{e(b)}</li>" for b in bl) + "</ul></div>")
    out.append(f'<h2>{d["h"]["exp"]}</h2>')
    for t, s, bl in d["exp"]:
        out.append(f'<div class="item"><div class="t">{e(t)} <span class="s">· {e(s)}</span></div>' + ("<ul>" + "".join(f"<li>{e(b)}</li>" for b in bl) + "</ul>" if bl else "") + "</div>")
    out.append(f'<h2>{d["h"]["proj"]}</h2>')
    for t, s, bl in d["proj"]:
        out.append(f'<div class="item"><div class="t">{e(t)} <span class="s">· {e(s)}</span></div><ul>' + "".join(f"<li>{e(b)}</li>" for b in bl) + "</ul></div>")
    out.append(f'<p class="note">{e(d["note"])}</p></body></html>')
    return "".join(out)


def rule(par) -> None:
    pPr = par._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "2A9D8F")):
        bt.set(qn(k), v)
    b.append(bt); pPr.append(b)


def make_docx(d: dict, path: Path, full: bool = True) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(16); sec.top_margin = sec.bottom_margin = Mm(14)
    st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.paragraph_format.space_after = Pt(2)

    def para(text="", bold=False, size=None, color=None, after=2, before=0, align=None):
        p = doc.add_paragraph(); r = p.add_run(text); r.bold = bold
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color
        p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
        if align: p.alignment = align
        return p

    para(d["name"], True, 24, NAVY, 0)
    para(d["role"], True, 12.5, TEAL, 2)
    para("  ·  ".join(contact_of(d, full)), size=9, after=4)
    rule(doc.paragraphs[-1])

    def head(t):
        p = para(t.upper(), True, 10.5, NAVY, 3, 10); rule(p)

    head(d["h"]["summary"]); para(d["summary"])
    head(d["h"]["skills"])
    for a, b in d["skills"]:
        p = doc.add_paragraph(); r = p.add_run(a + ": "); r.bold = True; r.font.color.rgb = NAVY; p.add_run(b); p.paragraph_format.space_after = Pt(1)

    def item(t, s, bl):
        p = doc.add_paragraph(); r = p.add_run(t); r.bold = True; r.font.color.rgb = NAVY
        r2 = p.add_run("  ·  " + s); r2.font.size = Pt(9); r2.font.color.rgb = RGBColor(0x6B, 0x77, 0x85)
        p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1); p.paragraph_format.keep_with_next = True
        for b in bl:
            q = doc.add_paragraph(b, style="List Bullet"); q.paragraph_format.space_after = Pt(0)

    head(d["h"]["edu"])
    for t, s, bl in d["edu"]: item(t, s, bl)
    head(d["h"]["exp"])
    for t, s, bl in d["exp"]: item(t, s, bl)
    head(d["h"]["proj"])
    for t, s, bl in d["proj"]: item(t, s, bl)
    para(d["note"], size=8.5, color=RGBColor(0x6B, 0x77, 0x85), before=8)
    doc.save(path)


if __name__ == "__main__":
    import sys
    full_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else None   # private versions (phone, address, YouTube) go here, never into the repo
    for k, d in DATA.items():
        (OUT / f"cv.{k}.html").write_text(make_html(d, False), encoding="utf-8")       # public version for the website
        if full_dir:
            full_dir.mkdir(parents=True, exist_ok=True)
            (full_dir / f"cv.{k}.full.html").write_text(make_html(d, True), encoding="utf-8")
            make_docx(d, full_dir / f"{d['file']}.docx", True)
        print("built", k)
