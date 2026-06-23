#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор персональных коммерческих предложений «РБ Факторинг».

Возможности:
  • автоподстановка данных конкретного клиента из CSV (offer/clients.csv);
  • две языковые версии — RU и EN (для международных контрагентов);
  • тарифная сетка под разные объёмы лимита с подсветкой уровня клиента;
  • самодостаточный HTML (встроенные шрифты + логотип) и, по флагу, PDF.

Запуск:
  python3 generate.py                 # из clients.csv -> output/*.html
  python3 generate.py --pdf           # то же + рендер PDF (нужен Node+Playwright)
  python3 generate.py --csv my.csv    # другой источник данных
  python3 generate.py --self-test     # один демо-клиент RU+EN без CSV
"""
import argparse, base64, csv, html, pathlib, re, subprocess, sys, os

HERE = pathlib.Path(__file__).parent
ASSETS = HERE / "assets"
OUT = HERE / "output"

# ---------------------------------------------------------------- ассеты
def load_assets():
    fonts = (ASSETS / "fonts_embedded.css").read_text(encoding="utf-8")
    styles = (ASSETS / "styles.css").read_text(encoding="utf-8")
    logo = base64.b64encode((ASSETS / "logo.png").read_bytes()).decode()
    return fonts, styles, logo

# ---------------------------------------------------------------- значения по умолчанию
DEFAULTS = {
    "factor": "ООО «РБ Факторинг»",
    "factoring_type": "recourse",   # recourse | nonrecourse
    "financing_pct": "100",
    "deferral_days": "90",
    "currency": "RUB",
    "limit_bn": "70",               # доступный лимит, млрд
    "key_rate_markup": "3.9",       # надбавка к ключевой ставке, % годовых
    "factoring_commission": "0.15", # факторинговая комиссия, % от УПД
    "limit_term_days": "10",
    "validity_days": "30",
    "doc_number": "",
    "date": "2026-06-09",
    "manager_name": "Илья Кириллин",
    "manager_phone": "+7 926 090 23 15",
    "manager_email": "i.a.kirillin@tbank.ru",
    "manager_site": "www.tbank.ru",
}

# Тарифная сетка по объёму лимита (индикативная). min_bn — порог для подсветки уровня клиента.
TARIFF_TIERS = [
    {"min_bn": 0.0,  "fin": "90",  "markup": "6.0", "comm": "0.40"},
    {"min_bn": 0.1,  "fin": "95",  "markup": "5.0", "comm": "0.30"},
    {"min_bn": 1.0,  "fin": "100", "markup": "4.4", "comm": "0.20"},
    {"min_bn": 10.0, "fin": "100", "markup": "3.9", "comm": "0.15"},
]

# ---------------------------------------------------------------- иконки (inline SVG)
IC = {
    "coin": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="12" cy="12" r="9"/><path d="M12 7v10M9.5 9.2c0-1.2 1.1-1.9 2.5-1.9s2.4.7 2.4 1.8c0 2.6-4.9 1.6-4.9 4.2 0 1.2 1.1 1.9 2.5 1.9s2.5-.7 2.5-1.9"/></svg>',
    "clock": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    "shield": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 3l8 4v5c0 5-3.4 7.7-8 9-4.6-1.3-8-4-8-9V7l8-4Z"/><path d="m9 12 2 2 4-4"/></svg>',
    "chart": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M3 3v18h18"/><path d="m7 14 3-3 3 3 5-6"/></svg>',
    "bolt": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M13 2 4.5 13H11l-1 9 8.5-11H12l1-9Z"/></svg>',
    "screen": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M8 4v5M8 14h8"/></svg>',
    "lock": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></svg>',
    "scale": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 3v18M5 21h14M7 7h10M7 7 4 14h6L7 7Zm10 0-3 7h6l-3-7Z"/></svg>',
    "gear": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="12" cy="12" r="3.2"/><path d="M19 12a7 7 0 0 0-.1-1.2l2-1.5-2-3.4-2.3 1a7 7 0 0 0-2-1.2L16.2 3h-4l-.4 2.5a7 7 0 0 0-2 1.2l-2.3-1-2 3.4 2 1.5A7 7 0 0 0 5 12a7 7 0 0 0 .1 1.2l-2 1.5 2 3.4 2.3-1a7 7 0 0 0 2 1.2L11.8 21h4l.4-2.5a7 7 0 0 0 2-1.2l2.3 1 2-3.4-2-1.5A7 7 0 0 0 19 12Z"/></svg>',
    "ph": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92Z"/></svg>',
    "mail": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/></svg>',
    "globe": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10Z"/></svg>',
}

# ---------------------------------------------------------------- словарь переводов
MONTHS_RU = ["", "января", "февраля", "марта", "апреля", "мая", "июня", "июля",
             "августа", "сентября", "октября", "ноября", "декабря"]
MONTHS_EN = ["", "January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"]

I18N = {
"ru": {
  "html_lang": "ru",
  "doc_title": "Коммерческое предложение · Факторинговое обслуживание",
  "brand_full": "РБ Факторинг · Группа Т-Банк",
  "confidential": "Конфиденциально",
  "foot_tag": "Коммерческое предложение · Конфиденциально",
  "up_to": "до",
  # обложка
  "cover_prepared": "Персональное предложение, подготовленное для",
  "hl1": "финансирование от&nbsp;суммы счёта-фактуры",
  "hl2": "отсрочка платежа для&nbsp;ваших покупателей",
  "hl3": "срок установления лимита финансирования",
  "hl_days": "дн.",
  "cover_eyebrow": "Коммерческое предложение",
  "cover_title1": "Факторинговое",
  "cover_title2": "обслуживание",
  "cover_sub": "Финансирование под уступку дебиторской задолженности, профессиональное управление расчётами и защита от кассовых разрывов — на условиях, разработанных индивидуально под вашу компанию.",
  "meta_prepared": "Подготовлено для",
  "meta_date": "Дата",
  "meta_doc": "Документ",
  "meta_valid": "действительно {n} дней",
  "meta_doc_note": "предварительные условия",
  # письмо
  "sec_letter": "Обращение",
  "eyebrow_letter": "Сопроводительное письмо",
  "salutation": "Уважаемые Господа!",
  "letter_p1": "Мы рады представить вашему вниманию комплекс факторинговых услуг, разработанный индивидуально для вашей компании. Надеемся, что он в полной мере ответит вашим ожиданиям и задачам бизнеса.",
  "letter_p2": "<strong>РБ Факторинг</strong> оказывает полный спектр факторинговых услуг. В рамках настоящего предложения мы предлагаем факторинговое обслуживание <strong>{ftype_lc}</strong>.",
  "letter_quote": "Согласно условиям договора факторинга, мы предоставим вам финансирование под уступку дебиторской задолженности <span class=\"big\">до {pct}% от суммы выставленного счёта-фактуры.</span>",
  "letter_p4": "Это поможет пополнить оборотные средства компании и предоставить отсрочку платежа покупателям, избежав при этом риска неоплаты с их стороны, с возможностью списания дебиторской задолженности в момент уступки денежных требований фактору.",
  "letter_p5": "Мы также возьмём на себя <strong>административное управление дебиторской задолженностью</strong>: контроль за своевременным погашением задолженности вашими покупателями и процедуру напоминаний при задержке платежей. Вы будете получать оперативную информацию в форме удобных отчётов о состоянии дебиторской задолженности и о графике платежей покупателей.",
  "letter_p6": "Надеемся на взаимовыгодное и долгосрочное сотрудничество.",
  "sig_role": "Региональный менеджер · РБ Факторинг",
  # преимущества
  "sec_value": "Ценность для бизнеса",
  "eyebrow_value": "Почему факторинг с РБ Факторинг",
  "value_title": "Оборотный капитал, который<br>работает на ваш рост",
  "value_lead": "Шесть причин, по которым крупнейшие поставщики выбирают факторинговое обслуживание в составе финансовой экосистемы Т-Банка.",
  "cards": [
    ("coin",   "Финансирование <span class=\"hl\">до {pct}%</span> от суммы счёта-фактуры", "Получайте деньги по поставке сразу — без ожидания оплаты от покупателя и без отвлечения собственных средств."),
    ("clock",  "Отсрочка покупателям <span class=\"hl\">до {def_d} дней</span>", "Предоставляйте клиентам комфортные условия оплаты и выигрывайте тендеры, не теряя в оборачиваемости."),
    ("shield", "Защита от <span class=\"hl\">риска неоплаты</span>", "Профессиональная оценка дебиторов и управление рисками снижают вероятность кассовых разрывов."),
    ("chart",  "Управление <span class=\"hl\">дебиторкой</span> «под ключ»", "Контроль сроков, напоминания покупателям и прозрачная отчётность по платежам — на нашей стороне."),
    ("bolt",   "Решение <span class=\"hl\">за {lim_d} дней</span>", "Установление лимита — до {lim_d} рабочих дней с момента предоставления полного пакета документов."),
    ("screen", "Цифровой сервис в <span class=\"hl\">ЛК Фактора</span>", "Полностью электронный документооборот, верификация от дебитора и онлайн-мониторинг финансирования."),
  ],
  "ml1": "финансирование от суммы счёта-фактуры",
  "ml2": "максимальная отсрочка платежа",
  "ml3": "доступный размер лимита",
  "unit_bn": "млрд ₽",
  # условия
  "sec_terms": "Условия сотрудничества",
  "eyebrow_terms": "Параметры сделки",
  "terms_title": "Условия сотрудничества",
  "t_supplier": "Поставщик",
  "t_factor": "Фактор",
  "t_debtors": "Дебиторы",
  "t_ftype": "Тип факторинга",
  "t_financing": "Финансирование",
  "t_deferral": "Отсрочка платежа",
  "t_currency": "Валюта",
  "t_scheme": "Схема финансирования",
  "t_rate": "Комиссия за денежные ресурсы",
  "t_comm": "Факторинговая комиссия",
  "t_law": "Правовые нормы / юрисдикция",
  "t_limterm": "Сроки установления лимитов",
  "v_financing": "{pct}% от суммы счёта-фактуры",
  "v_financing_note": "(определяется по результатам анализа)",
  "v_deferral": "до {def_d} к.д. (в соответствии с контрактной отсрочкой платежа)",
  "v_scheme": "Через ЛК Фактора + верификация от дебитора · для лимита от {limit}",
  "v_rate": "Ключевая ставка ЦБ&nbsp;РФ<sup>**</sup> <b>+ {markup}% годовых</b>",
  "v_comm": "<b>{comm}%</b> от суммы УПД",
  "v_law": "Арбитражный суд г. Москвы · законодательство РФ",
  "v_limterm": "до {lim_d} рабочих дней с момента предоставления полного пакета документов",
  "fn1": "Комиссии указаны без НДС.",
  "fn2": "Для целей начисления Комиссии за финансирование ключевая ставка Банка России фиксируется в момент заключения Договора и действует до момента её изменения в порядке, установленном действующим законодательством РФ. В случае изменения ключевой ставки Банка России ставки Комиссии за финансирование автоматически изменяются со дня, следующего за днём объявления изменения. Изменённые ставки распространяются на суммы выплаченного Клиенту Финансирования, не погашенного полностью или в части к моменту изменения ключевой ставки, а также на Финансирование, выплачиваемое после введения в действие нового размера ключевой ставки.",
  # тарифы
  "sec_tariffs": "Тарифная сетка",
  "eyebrow_tariffs": "Условия по объёму лимита",
  "tariffs_title": "Тарифы под разные лимиты",
  "tariffs_lead": "Ставки снижаются по мере роста объёма уступаемой задолженности. Ниже — индикативная сетка; финальные условия фиксируются по результатам анализа.",
  "th_tier": "Объём лимита",
  "th_fin": "Финансирование",
  "th_rate": "Комиссия за ресурсы",
  "th_comm": "Факторинговая комиссия",
  "tier_ranges": ["до 100 млн ₽", "100 млн – 1 млрд ₽", "1 – 10 млрд ₽", "от 10 млрд ₽"],
  "fin_cell": "до {fin}%",
  "rate_cell": "ключевая + {markup}%",
  "comm_cell": "{comm}%",
  "your_tier": "ваш уровень",
  "tariff_note": "Ставки указаны без НДС и носят индикативный характер. «Комиссия за ресурсы» рассчитывается как ключевая ставка Банка России плюс указанная надбавка, годовых.",
  "info": [
    ("lock",  "Прозрачность", "Никаких скрытых платежей: все комиссии зафиксированы в договоре."),
    ("gear",  "Индивидуальный расчёт", "Финальные ставки определяются по итогам анализа компании и дебиторов."),
    ("scale", "Гибкая структура", "Регресс или без регресса, открытый или закрытый — под вашу модель продаж."),
  ],
  # процесс
  "sec_start": "Старт сотрудничества",
  "eyebrow_start": "Как мы начинаем работу",
  "start_title": "Пять шагов до первого<br>финансирования",
  "steps": [
    ("Заявка и пакет документов", "Вы направляете нам комплект документов по компании и дебиторам — мы помогаем сформировать перечень."),
    ("Финансовый анализ и оценка лимита", "Проводим анализ вашей компании и дебиторской задолженности, рассчитываем индивидуальные параметры."),
    ("Утверждение Кредитным комитетом", "Кредитный комитет утверждает окончательные условия сотрудничества — в срок до {lim_d} рабочих дней."),
    ("Заключение договора и подключение к ЛК", "Подписываем договор факторинга и настраиваем электронный документооборот в Личном кабинете Фактора."),
    ("Финансирование поставок", "Уступаете денежные требования — и получаете финансирование до {pct}% от суммы счёта-фактуры."),
  ],
  "cta_eyebrow": "Готовы обсудить детали?",
  "cta_h3": "Рассчитаем индивидуальные условия для&nbsp;{client}",
  "cta_p": "Свяжитесь с вашим персональным менеджером — и мы подготовим финальное предложение под параметры вашего бизнеса.",
  "disclaimer": "<b>Важно.</b> Данное предложение является предварительным и подлежит рассмотрению Кредитным комитетом. Окончательные условия сотрудничества будут утверждены после передачи пакета документов с вашей стороны и проведения финансового анализа компании. Настоящий документ не является публичной офертой в значении ст. 437 ГК РФ. Мы надеемся, что наше предложение соответствует вашим пожеланиям.",
  "ftype": {"recourse": "Факторинг с регрессом", "nonrecourse": "Факторинг без регресса"},
  "ftype_lc": {"recourse": "с правом регресса", "nonrecourse": "без права регресса"},
  "currency_name": {"RUB": "Российский рубль (RUB)", "USD": "Доллар США (USD)", "EUR": "Евро (EUR)", "CNY": "Китайский юань (CNY)"},
},
"en": {
  "html_lang": "en",
  "doc_title": "Commercial Proposal · Factoring Services",
  "brand_full": "RB Factoring · T-Bank Group",
  "confidential": "Confidential",
  "foot_tag": "Commercial Proposal · Confidential",
  "up_to": "up to",
  "cover_prepared": "A personalised proposal prepared for",
  "hl1": "financing of the invoice amount",
  "hl2": "payment deferral for your buyers",
  "hl3": "to approve the financing limit",
  "hl_days": "days",
  "cover_eyebrow": "Commercial Proposal",
  "cover_title1": "Factoring",
  "cover_title2": "services",
  "cover_sub": "Financing against assignment of receivables, professional management of settlements and protection from cash-flow gaps — on terms tailored individually to your company.",
  "meta_prepared": "Prepared for",
  "meta_date": "Date",
  "meta_doc": "Document",
  "meta_valid": "valid for {n} days",
  "meta_doc_note": "indicative terms",
  "sec_letter": "Cover Letter",
  "eyebrow_letter": "Cover letter",
  "salutation": "Dear Sirs,",
  "letter_p1": "We are pleased to present a suite of factoring services designed specifically for your company. We trust it will fully meet your expectations and business objectives.",
  "letter_p2": "<strong>RB Factoring</strong> provides the full range of factoring services. Under this proposal we offer factoring <strong>{ftype_lc}</strong>.",
  "letter_quote": "Under the factoring agreement, we will provide financing against assignment of receivables <span class=\"big\">of up to {pct}% of the issued invoice amount.</span>",
  "letter_p4": "This helps replenish your working capital and grant payment deferrals to your buyers, while protecting you from the risk of non-payment, with the option to derecognise the receivable at the moment it is assigned to the factor.",
  "letter_p5": "We will also take over <strong>administrative management of your receivables</strong>: monitoring timely repayment by your buyers and handling reminders in case of delays. You will receive prompt information via convenient reports on the status of your receivables and your buyers' payment schedule.",
  "letter_p6": "We look forward to a mutually beneficial and long-term partnership.",
  "sig_role": "Regional Manager · RB Factoring",
  "sec_value": "Value for Your Business",
  "eyebrow_value": "Why factoring with RB Factoring",
  "value_title": "Working capital that<br>fuels your growth",
  "value_lead": "Six reasons why leading suppliers choose factoring services within the T-Bank financial ecosystem.",
  "cards": [
    ("coin",   "Financing of <span class=\"hl\">up to {pct}%</span> of the invoice amount", "Receive cash for each delivery immediately — without waiting for buyer payment or tying up your own funds."),
    ("clock",  "Buyer deferral of <span class=\"hl\">up to {def_d} days</span>", "Offer your clients comfortable payment terms and win tenders without sacrificing turnover."),
    ("shield", "Protection from <span class=\"hl\">non-payment risk</span>", "Professional debtor assessment and risk management reduce the likelihood of cash-flow gaps."),
    ("chart",  "End-to-end <span class=\"hl\">receivables</span> management", "Deadline control, buyer reminders and transparent payment reporting — handled on our side."),
    ("bolt",   "Decision <span class=\"hl\">within {lim_d} days</span>", "A limit is set within {lim_d} business days from submission of the complete document package."),
    ("screen", "Digital service in the <span class=\"hl\">Factor portal</span>", "Fully electronic document flow, debtor verification and online monitoring of financing."),
  ],
  "ml1": "financing of the invoice amount",
  "ml2": "maximum payment deferral",
  "ml3": "available limit size",
  "unit_bn": "bn",
  "sec_terms": "Terms of Cooperation",
  "eyebrow_terms": "Deal parameters",
  "terms_title": "Terms of Cooperation",
  "t_supplier": "Supplier",
  "t_factor": "Factor",
  "t_debtors": "Debtors",
  "t_ftype": "Factoring type",
  "t_financing": "Financing",
  "t_deferral": "Payment deferral",
  "t_currency": "Currency",
  "t_scheme": "Financing scheme",
  "t_rate": "Funding commission",
  "t_comm": "Factoring fee",
  "t_law": "Governing law / jurisdiction",
  "t_limterm": "Limit approval timeline",
  "v_financing": "{pct}% of the invoice amount",
  "v_financing_note": "(determined by the results of the analysis)",
  "v_deferral": "up to {def_d} calendar days (in line with the contractual deferral)",
  "v_scheme": "Via the Factor portal + debtor verification · for a limit from {limit}",
  "v_rate": "CBR key rate<sup>**</sup> <b>+ {markup}% p.a.</b>",
  "v_comm": "<b>{comm}%</b> of the invoice (UPD) amount",
  "v_law": "Moscow Arbitration Court · laws of the Russian Federation",
  "v_limterm": "up to {lim_d} business days from submission of the complete document package",
  "fn1": "Fees are stated exclusive of VAT.",
  "fn2": "For the purpose of calculating the funding commission, the Bank of Russia key rate is fixed at the moment the Agreement is concluded and applies until it changes in the manner established by the applicable laws of the Russian Federation. Should the key rate change, the funding commission rates change automatically from the day following the announcement. The revised rates apply to amounts of Financing disbursed to the Client and not fully or partly repaid by the time of the change, as well as to Financing disbursed after the new key rate takes effect.",
  "sec_tariffs": "Tariff Grid",
  "eyebrow_tariffs": "Terms by limit volume",
  "tariffs_title": "Tariffs for different limits",
  "tariffs_lead": "Rates decrease as the volume of assigned receivables grows. Below is an indicative grid; final terms are fixed based on the results of the analysis.",
  "th_tier": "Limit volume",
  "th_fin": "Financing",
  "th_rate": "Funding rate",
  "th_comm": "Factoring fee",
  "tier_ranges": ["up to RUB 100 m", "RUB 100 m – 1 bn", "RUB 1 – 10 bn", "from RUB 10 bn"],
  "fin_cell": "up to {fin}%",
  "rate_cell": "key rate + {markup}%",
  "comm_cell": "{comm}%",
  "your_tier": "your tier",
  "tariff_note": "Rates are stated exclusive of VAT and are indicative. The funding rate is calculated as the Bank of Russia key rate plus the stated margin, per annum.",
  "info": [
    ("lock",  "Transparency", "No hidden charges: all fees are fixed in the agreement."),
    ("gear",  "Tailored pricing", "Final rates are set based on the analysis of your company and debtors."),
    ("scale", "Flexible structure", "Recourse or non-recourse, open or closed — to fit your sales model."),
  ],
  "sec_start": "Getting Started",
  "eyebrow_start": "How we get started",
  "start_title": "Five steps to your<br>first funding",
  "steps": [
    ("Application & document package", "You send us a set of documents on your company and debtors — we help you compile the list."),
    ("Financial analysis & limit assessment", "We analyse your company and receivables and calculate the individual parameters."),
    ("Credit Committee approval", "The Credit Committee approves the final terms of cooperation — within {lim_d} business days."),
    ("Agreement & portal onboarding", "We sign the factoring agreement and set up electronic document flow in the Factor portal."),
    ("Financing of deliveries", "You assign the monetary claims — and receive financing of up to {pct}% of the invoice amount."),
  ],
  "cta_eyebrow": "Ready to discuss?",
  "cta_h3": "Let us tailor the terms for&nbsp;{client}",
  "cta_p": "Get in touch with your dedicated manager — and we will prepare a final proposal matched to your business parameters.",
  "disclaimer": "<b>Important.</b> This proposal is preliminary and subject to review by the Credit Committee. The final terms of cooperation will be approved after the document package is provided by you and the financial analysis of your company is completed. This document does not constitute a public offer within the meaning of Art. 437 of the Civil Code of the Russian Federation. We hope our proposal meets your expectations.",
  "ftype": {"recourse": "Factoring with recourse", "nonrecourse": "Non-recourse factoring"},
  "ftype_lc": {"recourse": "with recourse", "nonrecourse": "without recourse"},
  "currency_name": {"RUB": "Russian Rouble (RUB)", "USD": "US Dollar (USD)", "EUR": "Euro (EUR)", "CNY": "Chinese Yuan (CNY)"},
},
}

# ---------------------------------------------------------------- хелперы форматирования
def fmt_num(s, lang):
    """Десятичный разделитель: запятая для RU, точка для EN."""
    s = str(s).strip()
    return s.replace(".", ",") if lang == "ru" else s.replace(",", ".")

def fmt_date(iso, lang):
    try:
        y, m, d = (int(x) for x in iso.split("-"))
    except Exception:
        return iso
    months = MONTHS_RU if lang == "ru" else MONTHS_EN
    return f"{d:02d} {months[m]} {y}"

def fmt_limit(limit_bn, lang):
    n = fmt_num(limit_bn, lang)
    return f"{n} {I18N[lang]['unit_bn']}" if lang == "ru" else f"RUB {n} {I18N[lang]['unit_bn']}"

def initials(name):
    parts = [p for p in re.split(r"\s+", name.strip()) if p]
    return "".join(p[0].upper() for p in parts[:2]) if parts else "·"

def doc_number(rec, lang):
    if rec.get("doc_number"):
        return rec["doc_number"]
    iso = rec.get("date", "2026-01-01")
    try:
        y, m, d = iso.split("-")
    except ValueError:
        y, m, d = "2026", "01", "01"
    prefix = "КП" if lang == "ru" else "CO"
    return f"{prefix}-{y}/{m}-{d}"

def slugify(s):
    s = re.sub(r"[«»\"'’]", "", s)
    s = re.sub(r"[^0-9A-Za-zА-Яа-я]+", "-", s).strip("-").lower()
    translit = {"а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"e","ж":"zh","з":"z","и":"i","й":"i","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r","с":"s","т":"t","у":"u","ф":"f","х":"h","ц":"c","ч":"ch","ш":"sh","щ":"sch","ъ":"","ы":"y","ь":"","э":"e","ю":"yu","я":"ya"}
    return "".join(translit.get(ch, ch) for ch in s) or "client"

# ---------------------------------------------------------------- сборка страниц
def build_html(rec, lang, fonts, styles, logo):
    t = I18N[lang]
    g = lambda k: t[k]
    img = f'data:image/png;base64,{logo}'

    pct   = rec["financing_pct"]
    def_d = rec["deferral_days"]
    lim_d = rec["limit_term_days"]
    markup = fmt_num(rec["key_rate_markup"], lang)
    comm   = fmt_num(rec["factoring_commission"], lang)
    limit_bn = rec["limit_bn"]
    limit_str = fmt_limit(limit_bn, lang)
    client = html.escape(rec["client_name"])
    debtors = html.escape(rec["debtors"])
    ftype_key = rec.get("factoring_type", "recourse")
    mgr = rec["manager_name"]; mgr_e = html.escape(mgr)
    role = g("sig_role")
    date_str = fmt_date(rec["date"], lang)
    docno = html.escape(doc_number(rec, lang))
    cur = rec.get("currency", "RUB")

    foot = lambda pg: f"""  <div class="sheet__foot">
    <span class="brand">{g('brand_full')}</span>
    <span>{g('foot_tag')}</span>
    <span class="pg"><b>{pg:02d}</b> / 06</span>
  </div>"""

    head = lambda num, sec: f"""  <div class="sheet__head">
    <div class="logo-lockup sm"><img src="{img}" alt=""></div>
    <div class="section-id">{num} · <b>{sec}</b></div>
  </div>
  <div class="sheet__band"></div>"""

    # --- стр.1 обложка ---
    p1 = f"""<section class="page cover">
  <div class="cover__topbar"></div>
  <div class="cover__head">
    <div class="logo-lockup"><img src="{img}" alt="T-Bank × Partner"></div>
    <div class="conf-badge"><span class="dot"></span>{g('confidential')}</div>
  </div>
  <div class="cover__intro">
    <div class="eyebrow">{g('brand_full')}</div>
    <div class="cover__client">{g('cover_prepared')}&nbsp; <b>{client}</b></div>
    <div class="cover__highlights">
      <div class="h"><div class="hv">{g('up_to')}&nbsp;<span>{pct}%</span></div><div class="hl">{g('hl1')}</div></div>
      <div class="h"><div class="hv">{g('up_to')}&nbsp;<span>{def_d}</span> {g('hl_days')}</div><div class="hl">{g('hl2')}</div></div>
      <div class="h"><div class="hv">{g('up_to')}&nbsp;<span>{lim_d}</span> {g('hl_days')}</div><div class="hl">{g('hl3')}</div></div>
    </div>
  </div>
  <div class="cover__panel">
    <div class="cover__panel-inner">
      <div class="eyebrow">{g('cover_eyebrow')}</div>
      <div class="cover__title">{g('cover_title1')}<br><span class="accent">{g('cover_title2')}</span></div>
      <div class="cover__divider"></div>
      <div class="cover__sub">{g('cover_sub')}</div>
      <div class="cover__meta">
        <div class="cell"><div class="k">{g('meta_prepared')}</div><div class="v">{client}</div></div>
        <div class="cell"><div class="k">{g('meta_date')}</div><div class="v">{date_str}<br><small>{g('meta_valid').format(n=rec['validity_days'])}</small></div></div>
        <div class="cell"><div class="k">{g('meta_doc')}</div><div class="v">{docno}<br><small>{g('meta_doc_note')}</small></div></div>
      </div>
    </div>
  </div>
</section>"""

    # --- стр.2 письмо ---
    contacts = f"""<span class="ci">{IC['ph']}{html.escape(rec['manager_phone'])}</span>
          <span class="ci">{IC['mail']}<a href="mailto:{html.escape(rec['manager_email'])}">{html.escape(rec['manager_email'])}</a></span>
          <span class="ci">{IC['globe']}<a href="https://{html.escape(rec['manager_site'])}">{html.escape(rec['manager_site'])}</a></span>"""
    p2 = f"""<section class="page sheet letter">
{head('01', g('sec_letter'))}
  <div class="sheet__body">
    <div class="eyebrow">{g('eyebrow_letter')}</div>
    <p class="salut" style="margin-top:8px;">{g('salutation')}</p>
    <p class="drop">{g('letter_p1')}</p>
    <p>{g('letter_p2').format(ftype_lc=g('ftype_lc')[ftype_key])}</p>
    <div class="pullquote"><p>{g('letter_quote').format(pct=pct)}</p></div>
    <p>{g('letter_p4')}</p>
    <p>{g('letter_p5')}</p>
    <p>{g('letter_p6')}</p>
    <div class="signature">
      <div class="sig-mark">{initials(mgr)}</div>
      <div>
        <div class="who">{mgr_e}</div>
        <div class="role">{role}</div>
        <div class="contacts">
          {contacts}
        </div>
      </div>
    </div>
  </div>
{foot(2)}
</section>"""

    # --- стр.3 преимущества ---
    cards_html = "\n".join(
        f"""      <div class="card">
        <span class="num">{i+1:02d}</span>
        <div class="ic">{IC[ic]}</div>
        <h3>{title.format(pct=pct, def_d=def_d, lim_d=lim_d)}</h3>
        <p>{body.format(pct=pct, def_d=def_d, lim_d=lim_d)}</p>
      </div>"""
        for i, (ic, title, body) in enumerate(g("cards")))
    p3 = f"""<section class="page sheet">
{head('02', g('sec_value'))}
  <div class="sheet__body">
    <div class="eyebrow">{g('eyebrow_value')}</div>
    <h2 class="section-title" style="margin-top:8px;">{g('value_title')}</h2>
    <p class="section-lead">{g('value_lead')}</p>
    <div class="cards">
{cards_html}
    </div>
    <div class="metric-strip">
      <div class="metric"><div class="mv">{pct}<small>%</small></div><div class="ml">{g('ml1')}</div></div>
      <div class="metric"><div class="mv">{def_d}<small> {g('hl_days')}</small></div><div class="ml">{g('ml2')}</div></div>
      <div class="metric"><div class="mv">{fmt_num(limit_bn, lang)}<small> {g('unit_bn')}</small></div><div class="ml">{g('ml3')}</div></div>
    </div>
  </div>
{foot(3)}
</section>"""

    # --- стр.4 условия ---
    def trow(k, v, hl=False):
        cls = " highlight" if hl else ""
        return f"""      <div class="terms__row{cls}">
        <div class="terms__k"><span class="kdot"></span>{k}</div>
        <div class="terms__v">{v}</div>
      </div>"""
    rows = [
        trow(g('t_supplier'), f"<b>{client}</b>"),
        trow(g('t_factor'), html.escape(rec['factor'])),
        trow(g('t_debtors'), debtors),
        trow(g('t_ftype'), f'<span class="pill">{g("ftype")[ftype_key]}</span>'),
        trow(g('t_financing'), f'<b>{g("v_financing").format(pct=pct)}</b>&nbsp;<span style="color:var(--muted)">{g("v_financing_note")}</span>', hl=True),
        trow(g('t_deferral'), g('v_deferral').format(def_d=def_d)),
        trow(g('t_currency'), g('currency_name').get(cur, cur)),
        trow(g('t_scheme'), g('v_scheme').format(limit=limit_str)),
        trow(f"{g('t_rate')}<sup>*</sup>", g('v_rate').format(markup=markup), hl=True),
        trow(f"{g('t_comm')}<sup>*</sup>", g('v_comm').format(comm=comm)),
        trow(g('t_law'), g('v_law')),
        trow(g('t_limterm'), g('v_limterm').format(lim_d=lim_d)),
    ]
    p4 = f"""<section class="page sheet">
{head('03', g('sec_terms'))}
  <div class="sheet__body">
    <div class="eyebrow">{g('eyebrow_terms')}</div>
    <h2 class="section-title" style="margin-top:8px;">{g('terms_title')}</h2>
    <div class="terms">
{chr(10).join(rows)}
    </div>
    <div class="footnotes">
      <p><span class="mk">*</span>{g('fn1')}</p>
      <p><span class="mk">**</span>{g('fn2')}</p>
    </div>
  </div>
{foot(4)}
</section>"""

    # --- стр.5 тарифы ---
    try:
        lim_val = float(str(limit_bn).replace(",", "."))
    except ValueError:
        lim_val = 0.0
    cur_idx = max((i for i, tier in enumerate(TARIFF_TIERS) if lim_val >= tier["min_bn"]), default=0)
    tcells = []
    for i, tier in enumerate(TARIFF_TIERS):
        even = " r-even" if i % 2 == 1 else ""
        cur = " cur" if i == cur_idx else ""
        badge = f'<span class="yourbadge">{g("your_tier")}</span>' if i == cur_idx else ""
        tcells.append(f"""      <div class="td tier{cur}{even}"><span class="kdot"></span>{g('tier_ranges')[i]}{badge}</div>
      <div class="td{cur}{even}"><span class="big">{g('fin_cell').format(fin=tier['fin'])}</span></div>
      <div class="td{cur}{even}">{g('rate_cell').format(markup=fmt_num(tier['markup'], lang))}</div>
      <div class="td{cur}{even}"><b>{g('comm_cell').format(comm=fmt_num(tier['comm'], lang))}</b></div>""")
    info_html = "\n".join(
        f"""      <div class="info"><div class="ic">{IC[ic]}</div><h4>{ti}</h4><p>{bo}</p></div>"""
        for ic, ti, bo in g("info"))
    p5 = f"""<section class="page sheet">
{head('04', g('sec_tariffs'))}
  <div class="sheet__body">
    <div class="eyebrow">{g('eyebrow_tariffs')}</div>
    <h2 class="section-title" style="margin-top:8px;">{g('tariffs_title')}</h2>
    <p class="section-lead">{g('tariffs_lead')}</p>
    <div class="tariffs"><div class="tariffs__grid">
      <div class="th">{g('th_tier')}</div><div class="th">{g('th_fin')}</div><div class="th">{g('th_rate')}</div><div class="th">{g('th_comm')}</div>
{chr(10).join(tcells)}
    </div></div>
    <p class="tariff-note"><span class="mk">*</span>{g('tariff_note')}</p>
    <div class="info-grid">
{info_html}
    </div>
  </div>
{foot(5)}
</section>"""

    # --- стр.6 процесс + CTA ---
    steps_html = "\n".join(
        f"""      <div class="step">
        <div class="badge">{i+1}</div>
        <div class="st-body"><h4>{st_t}</h4><p>{st_b.format(pct=pct, lim_d=lim_d)}</p></div>
      </div>"""
        for i, (st_t, st_b) in enumerate(g("steps")))
    p6 = f"""<section class="page sheet">
{head('05', g('sec_start'))}
  <div class="sheet__body">
    <div class="eyebrow">{g('eyebrow_start')}</div>
    <h2 class="section-title" style="margin-top:8px;">{g('start_title')}</h2>
    <div class="steps">
{steps_html}
    </div>
    <div class="cta">
      <div class="eyebrow">{g('cta_eyebrow')}</div>
      <h3>{g('cta_h3').format(client=client)}</h3>
      <p>{g('cta_p')}</p>
      <div class="row">
        <span class="btn">{mgr_e} · {html.escape(rec['manager_phone'])}</span>
        <span class="contact-inline"><b>{html.escape(rec['manager_email'])}</b> · {html.escape(rec['manager_site'])}</span>
      </div>
    </div>
    <div class="disclaimer">{g('disclaimer')}</div>
  </div>
{foot(6)}
</section>"""

    css = fonts + "\n" + styles
    return f"""<!DOCTYPE html>
<html lang="{g('html_lang')}">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{g('doc_title')} · {client}</title>
<style>
{css}
</style>
</head>
<body>
{p1}
{p2}
{p3}
{p4}
{p5}
{p6}
</body>
</html>"""

# ---------------------------------------------------------------- ввод-вывод
def read_clients(csv_path):
    rows = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            rec = {k: (v.strip() if isinstance(v, str) else v) for k, v in raw.items() if k}
            if not rec.get("client_name"):
                continue
            merged = dict(DEFAULTS)
            merged.update({k: v for k, v in rec.items() if v not in (None, "")})
            rows.append(merged)
    return rows

def render_pdf(html_path):
    env = dict(os.environ)
    env.setdefault("PLAYWRIGHT_BROWSERS_PATH", "/opt/pw-browsers")
    pkg = env.get("PLAYWRIGHT_PKG")
    if not pkg:
        for c in ("/opt/node22/lib/node_modules/playwright", "/usr/lib/node_modules/playwright"):
            if pathlib.Path(c).exists():
                pkg = c; break
    if pkg:
        env["PLAYWRIGHT_PKG"] = pkg
    node = "/opt/node22/bin/node" if pathlib.Path("/opt/node22/bin/node").exists() else "node"
    subprocess.run([node, str(HERE / "render-pdf.mjs"), str(html_path)], check=True, env=env)

def main():
    ap = argparse.ArgumentParser(description="Генератор коммерческих предложений «РБ Факторинг».")
    ap.add_argument("--csv", default=str(HERE / "clients.csv"))
    ap.add_argument("--pdf", action="store_true", help="также отрендерить PDF")
    ap.add_argument("--langs", default="", help="переопределить языки, напр. ru,en")
    ap.add_argument("--self-test", action="store_true", help="один демо-клиент без CSV")
    args = ap.parse_args()

    fonts, styles, logo = load_assets()
    OUT.mkdir(exist_ok=True)

    if args.self_test:
        clients = [dict(DEFAULTS, client_name="ООО «КОНКОРД»",
                        debtors="ООО «ВСЕИНСТРУМЕНТЫ.РУ»", langs="ru,en")]
    else:
        clients = read_clients(args.csv)

    made = []
    for rec in clients:
        langs = (args.langs or rec.get("langs") or "ru").replace(";", ",")
        for lang in [x.strip() for x in langs.split(",") if x.strip()]:
            if lang not in I18N:
                print(f"  ! неизвестный язык: {lang}", file=sys.stderr); continue
            doc = build_html(rec, lang, fonts, styles, logo)
            name = f"{slugify(rec['client_name'])}__{lang}"
            html_path = OUT / f"{name}.html"
            html_path.write_text(doc, encoding="utf-8")
            print(f"HTML  -> {html_path.relative_to(HERE)}  ({html_path.stat().st_size//1024} KB)")
            made.append(html_path)
            if args.pdf:
                render_pdf(html_path)
                print(f"PDF   -> output/{name}.pdf")

    print(f"\nГотово: {len(made)} документ(ов) для {len(clients)} клиент(ов).")

if __name__ == "__main__":
    main()
