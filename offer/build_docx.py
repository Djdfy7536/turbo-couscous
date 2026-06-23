#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка РЕДАКТИРУЕМОГО Word-шаблона коммерческого предложения «РБ Факторинг».

Документ открывается в Microsoft Word / LibreOffice, поля заполняются прямо
в тексте, после чего сохраняется в PDF (Файл → Экспорт/Сохранить как → PDF).

Запуск:
  python3 build_docx.py            # RU и EN -> output/*.docx
  python3 build_docx.py --lang ru  # только RU
"""
import argparse, pathlib
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = pathlib.Path(__file__).parent
ASSETS = HERE / "assets"
OUT = HERE / "output"

# ---- фирменные цвета ----
INK       = RGBColor(0x16, 0x16, 0x1A)
INK_SOFT  = RGBColor(0x3C, 0x3C, 0x44)
GOLD      = RGBColor(0xB0, 0x89, 0x5A)
GOLD_DEEP = RGBColor(0x8C, 0x6F, 0x45)
GOLD_LT   = RGBColor(0xCB, 0xA8, 0x6B)
MUTED     = RGBColor(0x8C, 0x88, 0x7F)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
TXT       = RGBColor(0x2A, 0x2A, 0x2E)
CREAM_HEX = "FBF9F5"
GOLDBG_HEX= "F6EEE0"
INK_HEX   = "16161A"
YELLOW_HEX= "FFDD2D"
LINE_HEX  = "E7E2D8"

HEAD = "Georgia"     # заголовки — выдержанный серифный шрифт (есть в Windows/Mac)
BODY = "Calibri"     # текст — универсальный шрифт MS Office

# =================================================================== данные
DEFAULTS = dict(
    factor="ООО «РБ Факторинг»",
    financing_pct="100", deferral_days="90", limit_bn="70",
    key_rate_markup="3.9", factoring_commission="0.15", limit_term_days="10",
    validity_days="30", currency_ru="Российский рубль (RUB)", currency_en="Russian Rouble (RUB)",
    manager_name="Илья Кириллин", manager_phone="+7 926 090 23 15",
    manager_email="i.a.kirillin@tbank.ru", manager_site="www.tbank.ru",
)
EXAMPLE = dict(DEFAULTS, client_name_ru="ООО «КОНКОРД»", client_name_en="KONKORD LLC",
               debtors_ru="ООО «ВСЕИНСТРУМЕНТЫ.РУ»", debtors_en="VSEINSTRUMENTY.RU LLC",
               date_ru="09 июня 2026", date_en="09 June 2026",
               doc_ru="КП-2026/06-09", doc_en="CO-2026/06-09",
               ftype_ru="Факторинг с регрессом", ftype_en="Factoring with recourse",
               ftype_lc_ru="с правом регресса", ftype_lc_en="with recourse")

TIERS = [  # индикативная сетка: диапазон ru/en, финансирование, надбавка, комиссия
    ("до 100 млн ₽",       "up to RUB 100 m",   "90",  "6.0", "0.40"),
    ("100 млн – 1 млрд ₽", "RUB 100 m – 1 bn",  "95",  "5.0", "0.30"),
    ("1 – 10 млрд ₽",      "RUB 1 – 10 bn",     "100", "4.4", "0.20"),
    ("от 10 млрд ₽",       "from RUB 10 bn",    "100", "3.9", "0.15"),
]
CURRENT_TIER = 3  # для лимита 70 млрд

# =================================================================== тексты
T = {
"ru": {
 "title": "Коммерческое предложение",
 "brand": "РБ Факторинг · Группа Т-Банк",
 "confidential": "КОНФИДЕНЦИАЛЬНО",
 "prepared": "Персональное предложение, подготовленное для",
 "cover_eyebrow": "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ",
 "cover_title": "Факторинговое обслуживание",
 "cover_sub": "Финансирование под уступку дебиторской задолженности, профессиональное управление расчётами и защита от кассовых разрывов — на условиях, разработанных индивидуально под вашу компанию.",
 "m_prepared": "ПОДГОТОВЛЕНО ДЛЯ", "m_date": "ДАТА", "m_doc": "ДОКУМЕНТ",
 "hl": [("до 100%","финансирование от суммы счёта-фактуры"),
        ("до 90 дн.","отсрочка платежа для покупателей"),
        ("до 10 дн.","срок установления лимита")],
 "sec_letter": "01 · ОБРАЩЕНИЕ",
 "letter_eyebrow": "СОПРОВОДИТЕЛЬНОЕ ПИСЬМО",
 "salutation": "Уважаемые Господа!",
 "letter": [
   "Мы рады представить вашему вниманию комплекс факторинговых услуг, разработанный индивидуально для вашей компании. Надеемся, что он в полной мере ответит вашим ожиданиям и задачам бизнеса.",
   "**РБ Факторинг** оказывает полный спектр факторинговых услуг. В рамках настоящего предложения мы предлагаем факторинговое обслуживание **{ftype_lc}**.",
   "Это поможет пополнить оборотные средства компании и предоставить отсрочку платежа покупателям, избежав при этом риска неоплаты с их стороны, с возможностью списания дебиторской задолженности в момент уступки денежных требований фактору.",
   "Мы также возьмём на себя **административное управление дебиторской задолженностью**: контроль за своевременным погашением задолженности вашими покупателями и процедуру напоминаний при задержке платежей. Вы будете получать оперативную информацию в форме удобных отчётов о состоянии дебиторской задолженности и о графике платежей покупателей.",
   "Надеемся на взаимовыгодное и долгосрочное сотрудничество.",
 ],
 "quote": "Согласно условиям договора факторинга, мы предоставим вам финансирование под уступку дебиторской задолженности до {pct}% от суммы выставленного счёта-фактуры.",
 "sig_role": "Региональный менеджер · РБ Факторинг",
 "sec_value": "02 · ЦЕННОСТЬ ДЛЯ БИЗНЕСА",
 "value_eyebrow": "ПОЧЕМУ ФАКТОРИНГ С РБ ФАКТОРИНГ",
 "value_title": "Оборотный капитал, который работает на ваш рост",
 "cards": [
   ("Финансирование до {pct}%","Деньги по поставке сразу — без ожидания оплаты от покупателя и без отвлечения собственных средств."),
   ("Отсрочка до {def_d} дней","Предоставляйте покупателям комфортные условия оплаты и выигрывайте тендеры, не теряя в оборачиваемости."),
   ("Защита от риска неоплаты","Профессиональная оценка дебиторов и управление рисками снижают вероятность кассовых разрывов."),
   ("Управление дебиторкой «под ключ»","Контроль сроков, напоминания покупателям и прозрачная отчётность по платежам — на нашей стороне."),
   ("Решение за {lim_d} дней","Установление лимита — до {lim_d} рабочих дней с момента предоставления полного пакета документов."),
   ("Цифровой сервис в ЛК Фактора","Электронный документооборот, верификация от дебитора и онлайн-мониторинг финансирования."),
 ],
 "sec_terms": "03 · УСЛОВИЯ СОТРУДНИЧЕСТВА",
 "terms_eyebrow": "ПАРАМЕТРЫ СДЕЛКИ",
 "terms_title": "Условия сотрудничества",
 "terms": [
   ("Поставщик", "{client}"),
   ("Фактор", "{factor}"),
   ("Дебиторы", "{debtors}"),
   ("Тип факторинга", "{ftype}"),
   ("Финансирование", "{pct}% от суммы счёта-фактуры (определяется по результатам анализа)"),
   ("Отсрочка платежа", "до {def_d} к.д. (в соответствии с контрактной отсрочкой платежа)"),
   ("Валюта", "{currency}"),
   ("Схема финансирования", "Через ЛК Фактора + верификация от дебитора · для лимита от {limit}"),
   ("Комиссия за денежные ресурсы *", "Ключевая ставка ЦБ РФ ** + {markup}% годовых"),
   ("Факторинговая комиссия *", "{comm}% от суммы УПД"),
   ("Правовые нормы / юрисдикция", "Арбитражный суд г. Москвы · законодательство РФ"),
   ("Сроки установления лимитов", "до {lim_d} рабочих дней с момента предоставления полного пакета документов"),
 ],
 "fn1": "*  Комиссии указаны без НДС.",
 "fn2": "**  Ключевая ставка Банка России фиксируется в момент заключения Договора и действует до момента её изменения в порядке, установленном законодательством РФ. При изменении ключевой ставки ставки Комиссии за финансирование изменяются автоматически со дня, следующего за днём объявления изменения.",
 "sec_tariffs": "04 · ТАРИФНАЯ СЕТКА",
 "tariffs_eyebrow": "УСЛОВИЯ ПО ОБЪЁМУ ЛИМИТА",
 "tariffs_title": "Тарифы под разные лимиты",
 "tariffs_lead": "Ставки снижаются по мере роста объёма уступаемой задолженности. Ниже — индикативная сетка; финальные условия фиксируются по результатам анализа.",
 "th": ["Объём лимита","Финансирование","Комиссия за ресурсы","Факторинговая комиссия"],
 "tier_fmt": ("до {fin}%","ключевая + {markup}%","{comm}%"),
 "your_tier": "  ← ваш уровень",
 "tariff_note": "Ставки указаны без НДС и носят индикативный характер. «Комиссия за ресурсы» = ключевая ставка Банка России + указанная надбавка, годовых.",
 "sec_start": "05 · СТАРТ СОТРУДНИЧЕСТВА",
 "start_eyebrow": "КАК МЫ НАЧИНАЕМ РАБОТУ",
 "start_title": "Пять шагов до первого финансирования",
 "steps": [
   ("Заявка и пакет документов","Вы направляете комплект документов по компании и дебиторам — мы помогаем сформировать перечень."),
   ("Финансовый анализ и оценка лимита","Анализируем вашу компанию и дебиторскую задолженность, рассчитываем индивидуальные параметры."),
   ("Утверждение Кредитным комитетом","Кредитный комитет утверждает окончательные условия — в срок до {lim_d} рабочих дней."),
   ("Договор и подключение к ЛК","Подписываем договор факторинга и настраиваем электронный документооборот."),
   ("Финансирование поставок","Уступаете денежные требования — и получаете финансирование до {pct}% от суммы счёта-фактуры."),
 ],
 "cta_title": "Готовы обсудить детали?",
 "cta_text": "Свяжитесь с вашим персональным менеджером — и мы подготовим финальное предложение под параметры вашего бизнеса.",
 "disclaimer": "Важно. Данное предложение является предварительным и подлежит рассмотрению Кредитным комитетом. Окончательные условия будут утверждены после передачи пакета документов с вашей стороны и финансового анализа компании. Документ не является публичной офертой в значении ст. 437 ГК РФ.",
 "field": "client_name_ru", "debtors_k":"debtors_ru","date_k":"date_ru","doc_k":"doc_ru",
 "ftype_k":"ftype_ru","ftypelc_k":"ftype_lc_ru","cur_k":"currency_ru",
 "limit_str":"70 млрд ₽", "valid_fmt":"действительно {n} дней",
},
"en": {
 "title": "Commercial Proposal",
 "brand": "RB Factoring · T-Bank Group",
 "confidential": "CONFIDENTIAL",
 "prepared": "A personalised proposal prepared for",
 "cover_eyebrow": "COMMERCIAL PROPOSAL",
 "cover_title": "Factoring Services",
 "cover_sub": "Financing against assignment of receivables, professional management of settlements and protection from cash-flow gaps — on terms tailored individually to your company.",
 "m_prepared": "PREPARED FOR", "m_date": "DATE", "m_doc": "DOCUMENT",
 "hl": [("up to 100%","financing of the invoice amount"),
        ("up to 90 days","payment deferral for buyers"),
        ("up to 10 days","time to approve the limit")],
 "sec_letter": "01 · COVER LETTER",
 "letter_eyebrow": "COVER LETTER",
 "salutation": "Dear Sirs,",
 "letter": [
   "We are pleased to present a suite of factoring services designed specifically for your company. We trust it will fully meet your expectations and business objectives.",
   "**RB Factoring** provides the full range of factoring services. Under this proposal we offer factoring **{ftype_lc}**.",
   "This helps replenish your working capital and grant payment deferrals to your buyers, while protecting you from the risk of non-payment, with the option to derecognise the receivable at the moment it is assigned to the factor.",
   "We will also take over **administrative management of your receivables**: monitoring timely repayment by your buyers and handling reminders in case of delays. You will receive prompt reports on the status of your receivables and your buyers' payment schedule.",
   "We look forward to a mutually beneficial and long-term partnership.",
 ],
 "quote": "Under the factoring agreement, we will provide financing against assignment of receivables of up to {pct}% of the issued invoice amount.",
 "sig_role": "Regional Manager · RB Factoring",
 "sec_value": "02 · VALUE FOR YOUR BUSINESS",
 "value_eyebrow": "WHY FACTORING WITH RB FACTORING",
 "value_title": "Working capital that fuels your growth",
 "cards": [
   ("Financing up to {pct}%","Cash for each delivery immediately — without waiting for buyer payment or tying up your own funds."),
   ("Deferral up to {def_d} days","Offer your buyers comfortable payment terms and win tenders without sacrificing turnover."),
   ("Protection from non-payment","Professional debtor assessment and risk management reduce the likelihood of cash-flow gaps."),
   ("End-to-end receivables management","Deadline control, buyer reminders and transparent payment reporting — handled on our side."),
   ("Decision within {lim_d} days","A limit is set within {lim_d} business days from submission of the complete document package."),
   ("Digital service in Factor portal","Electronic document flow, debtor verification and online monitoring of financing."),
 ],
 "sec_terms": "03 · TERMS OF COOPERATION",
 "terms_eyebrow": "DEAL PARAMETERS",
 "terms_title": "Terms of Cooperation",
 "terms": [
   ("Supplier", "{client}"),
   ("Factor", "{factor}"),
   ("Debtors", "{debtors}"),
   ("Factoring type", "{ftype}"),
   ("Financing", "{pct}% of the invoice amount (determined by the analysis)"),
   ("Payment deferral", "up to {def_d} calendar days (in line with the contractual deferral)"),
   ("Currency", "{currency}"),
   ("Financing scheme", "Via the Factor portal + debtor verification · for a limit from {limit}"),
   ("Funding commission *", "CBR key rate ** + {markup}% p.a."),
   ("Factoring fee *", "{comm}% of the invoice (UPD) amount"),
   ("Governing law / jurisdiction", "Moscow Arbitration Court · laws of the Russian Federation"),
   ("Limit approval timeline", "up to {lim_d} business days from submission of the complete document package"),
 ],
 "fn1": "*  Fees are stated exclusive of VAT.",
 "fn2": "**  The Bank of Russia key rate is fixed when the Agreement is concluded and applies until it changes in the manner established by law. Should the key rate change, the funding commission rates change automatically from the day following the announcement.",
 "sec_tariffs": "04 · TARIFF GRID",
 "tariffs_eyebrow": "TERMS BY LIMIT VOLUME",
 "tariffs_title": "Tariffs for different limits",
 "tariffs_lead": "Rates decrease as the volume of assigned receivables grows. Below is an indicative grid; final terms are fixed based on the results of the analysis.",
 "th": ["Limit volume","Financing","Funding rate","Factoring fee"],
 "tier_fmt": ("up to {fin}%","key rate + {markup}%","{comm}%"),
 "your_tier": "  ← your tier",
 "tariff_note": "Rates are stated exclusive of VAT and are indicative. The funding rate = Bank of Russia key rate + the stated margin, per annum.",
 "sec_start": "05 · GETTING STARTED",
 "start_eyebrow": "HOW WE GET STARTED",
 "start_title": "Five steps to your first funding",
 "steps": [
   ("Application & document package","You send a set of documents on your company and debtors — we help you compile the list."),
   ("Financial analysis & limit assessment","We analyse your company and receivables and calculate the individual parameters."),
   ("Credit Committee approval","The Credit Committee approves the final terms — within {lim_d} business days."),
   ("Agreement & portal onboarding","We sign the factoring agreement and set up electronic document flow."),
   ("Financing of deliveries","You assign the monetary claims — and receive financing of up to {pct}% of the invoice amount."),
 ],
 "cta_title": "Ready to discuss?",
 "cta_text": "Get in touch with your dedicated manager — and we will prepare a final proposal matched to your business parameters.",
 "disclaimer": "Important. This proposal is preliminary and subject to review by the Credit Committee. Final terms will be approved after you provide the document package and the financial analysis is completed. This document does not constitute a public offer within the meaning of Art. 437 of the Civil Code of the Russian Federation.",
 "field": "client_name_en", "debtors_k":"debtors_en","date_k":"date_en","doc_k":"doc_en",
 "ftype_k":"ftype_en","ftypelc_k":"ftype_lc_en","cur_k":"currency_en",
 "limit_str":"RUB 70 bn", "valid_fmt":"valid for {n} days",
},
}

# =================================================================== low-level helpers
def _set_font(run, name):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    for a in ('w:ascii','w:hAnsi','w:cs'):
        rfonts.set(qn(a), name)

def _tracking(run, twips=30):
    rpr = run._element.get_or_add_rPr()
    sp = OxmlElement('w:spacing'); sp.set(qn('w:val'), str(twips)); rpr.append(sp)

def run(p, text, *, font=BODY, size=10.5, color=TXT, bold=False, italic=False, track=None, caps=False):
    r = p.add_run(text.upper() if caps else text)
    _set_font(r, font); r.font.size = Pt(size); r.font.color.rgb = color
    r.font.bold = bold; r.font.italic = italic
    if track is not None: _tracking(r, track)
    return r

def para(container, *, before=0, after=0, line=1.15, align=None):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    if align is not None: p.alignment = align
    return p

def rich(p, text, *, font=BODY, size=10.5, color=TXT):
    """**жирные** фрагменты -> отдельные run-ы."""
    for i, seg in enumerate(text.split("**")):
        if not seg:
            continue
        is_b = (i % 2 == 1)
        run(p, seg, font=font, size=size, color=(INK if is_b else color), bold=is_b)

def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for tag, val in (('top',top),('bottom',bottom),('start',left),('end',right),('left',left),('right',right)):
        e = OxmlElement('w:'+tag); e.set(qn('w:w'), str(val)); e.set(qn('w:type'),'dxa'); m.append(e)
    tcPr.append(m)

def no_borders(table):
    tbl = table._tbl; tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement('w:'+edge); e.set(qn('w:val'),'none'); e.set(qn('w:sz'),'0'); borders.append(e)
    tblPr.append(borders)

def hline(container, color="E7E2D8", space_after=0, sz=6):
    p = para(container, after=space_after)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr'); bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'),'single'); bottom.set(qn('w:sz'),str(sz))
    bottom.set(qn('w:space'),'1'); bottom.set(qn('w:color'),color)
    pbdr.append(bottom); pPr.append(pbdr)
    return p

def gold_rule(container):
    """короткая золотая линия-акцент."""
    p = para(container, before=4, after=6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr'); bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'),'single'); bottom.set(qn('w:sz'),'24')
    bottom.set(qn('w:space'),'1'); bottom.set(qn('w:color'),'B0895A')
    pbdr.append(bottom); pPr.append(pbdr)
    # ограничим ширину «линии» табуляцией
    r = run(p, "")
    ind = OxmlElement('w:ind'); ind.set(qn('w:right'),"8500"); pPr.append(ind)
    return p

def set_col_widths(table, widths_mm):
    table.autofit = False
    for row in table.rows:
        for i, w in enumerate(widths_mm):
            row.cells[i].width = Mm(w)

# =================================================================== document
def build(lang):
    t = T[lang]; d = EXAMPLE
    pct, def_d, lim_d = d["financing_pct"], d["deferral_days"], d["limit_term_days"]
    markup, comm = d["key_rate_markup"], d["factoring_commission"]
    ctx = dict(pct=pct, def_d=def_d, lim_d=lim_d, markup=markup, comm=comm,
               client=d[t["field"]], factor=d["factor"], debtors=d[t["debtors_k"]],
               ftype=d[t["ftype_k"]], ftype_lc=d[t["ftypelc_k"]], currency=d[t["cur_k"]],
               limit=t["limit_str"])

    doc = Document()
    # базовый стиль
    st = doc.styles['Normal']; st.font.name = BODY; st.font.size = Pt(10.5)
    st.font.color.rgb = TXT
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    for a in ('w:ascii','w:hAnsi','w:cs'):
        rfonts.set(qn(a), BODY)

    sec = doc.sections[0]
    sec.page_height, sec.page_width = Mm(297), Mm(210)
    sec.top_margin = sec.bottom_margin = Mm(16)
    sec.left_margin = sec.right_margin = Mm(18)

    # ---------- колонтитул (футер) ----------
    footer = sec.footer
    fp = footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(fp, f"{t['brand']}   ·   {t['title']}   ·   {t['confidential'].title()}",
        font=BODY, size=7.5, color=MUTED, track=20)

    # =============================== СТР.1 ОБЛОЖКА ===============================
    # логотип
    try:
        lp = para(doc, after=0); rr = lp.add_run(); rr.add_picture(str(ASSETS/"logo.png"), height=Mm(11))
    except Exception:
        run(para(doc), "T × P", font=HEAD, size=18, color=INK, bold=True)
    p = para(doc, before=2, after=10); run(p, t["confidential"], font=BODY, size=8, color=GOLD_DEEP, bold=True, track=40)

    run(para(doc, after=2), t["brand"], font=BODY, size=9, color=GOLD, bold=True, track=40)
    p = para(doc, after=6); run(p, t["prepared"] + "   ", size=10.5, color=MUTED)
    run(p, d[t["field"]], size=10.5, color=INK, bold=True)

    # полоса ключевых цифр
    hl = doc.add_table(rows=1, cols=3); hl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(hl); set_col_widths(hl, [58,58,58])
    for i,(big,lab) in enumerate(t["hl"]):
        c = hl.rows[0].cells[i]; c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        cell_margins(c, 120, 120, 40, 120)
        pp = c.paragraphs[0]; pp.paragraph_format.space_after = Pt(3)
        run(pp, big, font=HEAD, size=18, color=INK, bold=True)
        lp = para(c, line=1.1); run(lp, lab, size=8.5, color=MUTED)
    hline(doc, space_after=2)

    para(doc, after=8)
    # тёмная панель
    panel = doc.add_table(rows=1, cols=1); panel.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(panel); set_col_widths(panel, [174])
    pc = panel.rows[0].cells[0]; shade(pc, INK_HEX); cell_margins(pc, 360, 360, 320, 320)
    pc.paragraphs[0].paragraph_format.space_after = Pt(2)
    run(pc.paragraphs[0], t["cover_eyebrow"], font=BODY, size=8.5, color=GOLD_LT, bold=True, track=40)
    tp = para(pc, before=2, after=4); run(tp, t["cover_title"], font=HEAD, size=30, color=WHITE, bold=True)
    sp = para(pc, after=12, line=1.35); run(sp, t["cover_sub"], size=10.5, color=RGBColor(0xC9,0xC6,0xCE))
    # мета-строка
    meta = pc.add_table(rows=1, cols=3); no_borders(meta)
    metas = [(t["m_prepared"], d[t["field"]], ""),
             (t["m_date"], d[t["date_k"]], t["valid_fmt"].format(n=d["validity_days"])),
             (t["m_doc"], d[t["doc_k"]], "")]
    for i,(k,v,note) in enumerate(metas):
        c = meta.rows[0].cells[i]; cell_margins(c, 60, 60, 0, 120)
        kp = c.paragraphs[0]; kp.paragraph_format.space_after = Pt(3)
        run(kp, k, font=BODY, size=7, color=GOLD_LT, bold=True, track=30)
        vp = para(c, line=1.1); run(vp, v, size=10, color=WHITE, bold=True)
        if note:
            np = para(c, line=1.0); run(np, note, size=7.5, color=RGBColor(0xA7,0xA3,0xAC))
    for cc in meta.rows[0].cells: shade(cc, INK_HEX)

    # =============================== СТР.2 ПИСЬМО ===============================
    doc.add_page_break()
    section_head(doc, t["sec_letter"])
    run(para(doc, after=2), t["letter_eyebrow"], font=BODY, size=8, color=GOLD, bold=True, track=40)
    run(para(doc, after=8), t["salutation"], font=HEAD, size=15, color=INK, bold=True)
    rich(para(doc, after=8, line=1.45), t["letter"][0])
    rich(para(doc, after=10, line=1.45), t["letter"][1].format(**ctx))
    # цитата-плашка
    q = doc.add_table(rows=1, cols=1); no_borders(q); set_col_widths(q,[174])
    qc = q.rows[0].cells[0]; shade(qc, CREAM_HEX); cell_margins(qc, 160, 160, 220, 200)
    left_accent(qc)
    run(qc.paragraphs[0], t["quote"].format(**ctx), font=HEAD, size=11.5, color=INK_SOFT, bold=True)
    para(doc, after=2)
    for ptxt in t["letter"][2:]:
        rich(para(doc, after=8, line=1.45), ptxt.format(**ctx))
    para(doc, after=4); hline(doc, space_after=6)
    # подпись
    sig = doc.add_table(rows=1, cols=2); no_borders(sig); set_col_widths(sig, [18,156])
    mark = sig.rows[0].cells[0]; shade(mark, INK_HEX); cell_margins(mark,140,140,40,40)
    mark.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    mp = mark.paragraphs[0]; mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(mp, initials(d["manager_name"]), font=HEAD, size=15, color=GOLD_LT, bold=True)
    info = sig.rows[0].cells[1]; cell_margins(info,40,40,160,40); info.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run(info.paragraphs[0], d["manager_name"], font=HEAD, size=12, color=INK, bold=True)
    run(para(info, after=4), t["sig_role"], size=9.5, color=MUTED)
    cp = para(info, line=1.3)
    run(cp, d["manager_phone"] + "     ", size=9, color=INK_SOFT)
    run(cp, d["manager_email"] + "     ", size=9, color=INK_SOFT)
    run(cp, d["manager_site"], size=9, color=INK_SOFT)

    # =============================== СТР.3 ПРЕИМУЩЕСТВА ===============================
    doc.add_page_break()
    section_head(doc, t["sec_value"])
    run(para(doc, after=2), t["value_eyebrow"], font=BODY, size=8, color=GOLD, bold=True, track=40)
    run(para(doc, after=10), t["value_title"], font=HEAD, size=19, color=INK, bold=True)
    grid = doc.add_table(rows=3, cols=2); no_borders(grid); set_col_widths(grid,[86,86])
    cards = t["cards"]
    for idx,(ti,bo) in enumerate(cards):
        c = grid.rows[idx//2].cells[idx%2]; cell_margins(c,150,150,170,170); shade(c, CREAM_HEX)
        hp = c.paragraphs[0]; hp.paragraph_format.space_after = Pt(3)
        run(hp, f"{idx+1:02d}   ", font=HEAD, size=10, color=GOLD_LT, bold=True)
        run(hp, ti.format(**ctx), font=HEAD, size=11.5, color=INK, bold=True)
        run(para(c, line=1.3), bo.format(**ctx), size=9, color=MUTED)
    # тёмная лента метрик
    para(doc, after=6)
    ms = doc.add_table(rows=1, cols=3); no_borders(ms); set_col_widths(ms,[58,58,58])
    mvals = [("100 %","финансирование счёта-фактуры" if lang=="ru" else "of the invoice amount"),
             (f"{def_d} {'дн.' if lang=='ru' else 'days'}","макс. отсрочка платежа" if lang=="ru" else "max payment deferral"),
             (f"{d['limit_bn']} {'млрд ₽' if lang=='ru' else 'bn RUB'}","доступный лимит" if lang=="ru" else "available limit")]
    for i,(big,lab) in enumerate(mvals):
        c = ms.rows[0].cells[i]; shade(c, INK_HEX); cell_margins(c,150,150,150,150)
        bp = c.paragraphs[0]; bp.paragraph_format.space_after = Pt(3)
        run(bp, big, font=HEAD, size=17, color=RGBColor(0xFF,0xDD,0x2D), bold=True)
        run(para(c, line=1.1), lab, size=8, color=RGBColor(0xBB,0xB7,0xBF))

    # =============================== СТР.4 УСЛОВИЯ ===============================
    doc.add_page_break()
    section_head(doc, t["sec_terms"])
    run(para(doc, after=2), t["terms_eyebrow"], font=BODY, size=8, color=GOLD, bold=True, track=40)
    run(para(doc, after=8), t["terms_title"], font=HEAD, size=19, color=INK, bold=True)
    terms = doc.add_table(rows=len(t["terms"]), cols=2); no_borders(terms); set_col_widths(terms,[58,116])
    hl_rows = {4, 8}
    for i,(k,v) in enumerate(t["terms"]):
        kc, vc = terms.rows[i].cells
        bg = GOLDBG_HEX if i in hl_rows else (CREAM_HEX if i%2 else "FFFFFF")
        shade(kc, bg); shade(vc, bg); cell_margins(kc,90,90,150,120); cell_margins(vc,90,90,140,150)
        kc.vertical_alignment = vc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        run(kc.paragraphs[0], k, font=BODY, size=9, color=(GOLD_DEEP if i in hl_rows else INK_SOFT), bold=True)
        run(vc.paragraphs[0], v.format(**ctx), size=9.5, color=(INK if i in hl_rows else TXT),
            bold=(i in hl_rows))
    para(doc, after=4)
    run(para(doc, after=3, line=1.3), t["fn1"], size=7.6, color=MUTED)
    run(para(doc, line=1.3), t["fn2"], size=7.6, color=MUTED)

    # =============================== СТР.5 ТАРИФЫ ===============================
    doc.add_page_break()
    section_head(doc, t["sec_tariffs"])
    run(para(doc, after=2), t["tariffs_eyebrow"], font=BODY, size=8, color=GOLD, bold=True, track=40)
    run(para(doc, after=4), t["tariffs_title"], font=HEAD, size=19, color=INK, bold=True)
    run(para(doc, after=8, line=1.35), t["tariffs_lead"], size=10, color=MUTED)
    tar = doc.add_table(rows=len(TIERS)+1, cols=4); no_borders(tar); set_col_widths(tar,[56,34,46,38])
    for j,htext in enumerate(t["th"]):
        c = tar.rows[0].cells[j]; shade(c, INK_HEX); cell_margins(c,100,100,120,120)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        run(c.paragraphs[0], htext, font=BODY, size=7.6, color=GOLD_LT, bold=True, track=20)
    fin_f, rate_f, comm_f = t["tier_fmt"]
    for i,(rng_ru,rng_en,fin,mk,cm) in enumerate(TIERS):
        r = tar.rows[i+1]; cur = (i==CURRENT_TIER)
        bg = GOLDBG_HEX if cur else (CREAM_HEX if i%2 else "FFFFFF")
        rng = rng_ru if lang=="ru" else rng_en
        for c in r.cells: shade(c,bg); cell_margins(c,120,120,120,110); c.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        tp = r.cells[0].paragraphs[0]
        run(tp, rng, font=HEAD, size=9, color=INK, bold=True)
        if cur: run(tp, t["your_tier"], font=BODY, size=7.5, color=GOLD_DEEP, bold=True)
        run(r.cells[1].paragraphs[0], fin_f.format(fin=fin), font=HEAD, size=11, color=GOLD_DEEP, bold=True)
        run(r.cells[2].paragraphs[0], rate_f.format(markup=fmt(mk,lang)), size=9, color=TXT)
        run(r.cells[3].paragraphs[0], comm_f.format(comm=fmt(cm,lang)), size=9.5, color=INK, bold=True)
    para(doc, after=2)
    run(para(doc, line=1.35), t["tariff_note"], size=7.8, color=MUTED)

    # =============================== СТР.6 ПРОЦЕСС ===============================
    doc.add_page_break()
    section_head(doc, t["sec_start"])
    run(para(doc, after=2), t["start_eyebrow"], font=BODY, size=8, color=GOLD, bold=True, track=40)
    run(para(doc, after=10), t["start_title"], font=HEAD, size=19, color=INK, bold=True)
    steps = doc.add_table(rows=len(t["steps"]), cols=2); no_borders(steps); set_col_widths(steps,[14,160])
    for i,(stt,stb) in enumerate(t["steps"]):
        nc, cc = steps.rows[i].cells
        cell_margins(nc,80,80,20,40); nc.vertical_alignment=WD_ALIGN_VERTICAL.TOP
        npg=nc.paragraphs[0]; npg.alignment=WD_ALIGN_PARAGRAPH.CENTER
        run(npg, str(i+1), font=HEAD, size=13, color=GOLD_DEEP, bold=True)
        cell_margins(cc,60,140,120,40)
        run(cc.paragraphs[0], stt, font=HEAD, size=11.5, color=INK, bold=True)
        run(para(cc, line=1.3), stb.format(**ctx), size=9.3, color=MUTED)
    # CTA-плашка
    para(doc, after=4)
    cta = doc.add_table(rows=1, cols=1); no_borders(cta); set_col_widths(cta,[174])
    cc = cta.rows[0].cells[0]; shade(cc, INK_HEX); cell_margins(cc,300,300,300,300)
    run(cc.paragraphs[0], t["cta_title"], font=BODY, size=8.5, color=GOLD_LT, bold=True, track=40)
    run(para(cc, before=4, after=6), d["manager_name"] + " · " + d["manager_phone"], font=HEAD, size=15, color=WHITE, bold=True)
    run(para(cc, after=8, line=1.35), t["cta_text"], size=9.6, color=RGBColor(0xC9,0xC6,0xCE))
    lp = para(cc)
    run(lp, d["manager_email"] + "     ", size=9.5, color=RGBColor(0xFF,0xDD,0x2D), bold=True)
    run(lp, d["manager_site"], size=9.5, color=RGBColor(0xD7,0xD4,0xDC))
    # дисклеймер
    para(doc, after=4)
    dis = doc.add_table(rows=1, cols=1); no_borders(dis); set_col_widths(dis,[174])
    dc = dis.rows[0].cells[0]; shade(dc, CREAM_HEX); cell_margins(dc,140,140,160,160)
    dc.paragraphs[0].paragraph_format.line_spacing = 1.35
    run(dc.paragraphs[0], t["disclaimer"], size=7.8, color=MUTED)

    OUT.mkdir(exist_ok=True)
    out = OUT / f"Коммерческое-предложение-{'RU' if lang=='ru' else 'EN'}.docx"
    doc.save(str(out))
    print(f"DOCX  -> {out.relative_to(HERE)}")
    return out

# ---- маленькие помощники, зависящие от run/para ----
def section_head(doc, sec_label):
    tb = doc.add_table(rows=1, cols=2); no_borders(tb); set_col_widths(tb,[120,54])
    l = tb.rows[0].cells[0]; r = tb.rows[0].cells[1]
    cell_margins(l,0,0,0,0); cell_margins(r,0,0,0,0)
    try:
        rr = l.paragraphs[0].add_run(); rr.add_picture(str(ASSETS/"logo.png"), height=Mm(7))
    except Exception:
        run(l.paragraphs[0], "T × P", font=HEAD, size=11, color=INK, bold=True)
    rp = r.paragraphs[0]; rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(rp, sec_label, font=BODY, size=8, color=MUTED, bold=True, track=30)
    hline(doc, color="E7E2D8", space_after=6)

def left_accent(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders'); left = OxmlElement('w:left')
    left.set(qn('w:val'),'single'); left.set(qn('w:sz'),'24'); left.set(qn('w:space'),'0'); left.set(qn('w:color'),'B0895A')
    borders.append(left); tcPr.append(borders)

def initials(name):
    parts = [p for p in name.replace("«","").replace("»","").split() if p]
    return "".join(p[0].upper() for p in parts[:2]) or "·"

def fmt(s, lang):
    return s.replace(".", ",") if lang == "ru" else s

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--lang", default="ru,en")
    args = ap.parse_args()
    for lang in [x.strip() for x in args.lang.split(",") if x.strip()]:
        build(lang)

if __name__ == "__main__":
    main()
