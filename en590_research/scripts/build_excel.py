"""Сборка Excel-отчёта по предложениям EN590 10 ppm.

Запуск:  python3 build_excel.py <папка data> <выходной .xlsx>
"""
import json, re, sys, os
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from manual_sources import LISTINGS, OFFICIAL, BLOCKED

DATA, OUT = sys.argv[1], sys.argv[2]
CHECK_DATE = "2026-11-09"
BENCH = 1346.0   # ICE Low Sulphur Gasoil, USD/MT, 09.11.2026 07:34 GMT (oilpriceapi.com)
BENCH_SRC = "https://www.oilpriceapi.com/live/gasoil-price"

MON = dict(Jan=1, Feb=2, Mar=3, Apr=4, May=5, Jun=6, Jul=7, Aug=8, Sep=9, Oct=10, Nov=11, Dec=12)
def iso(d):
    m = re.match(r'([A-Za-z]{3})-(\d\d)-(\d\d)', d or '')
    return f"20{m.group(3)}-{MON[m.group(1)]:02d}-{m.group(2)}" if m else d

TEN = re.compile(r'10\s*-?\s*ppm|10\s*mg/kg|\bulsd\b|ultra[\s-]*low[\s-]*sul', re.I)
FREEMAIL = re.compile(r'[\w.\-]+@(gmail|yahoo|hotmail|outlook|aol|mail\.ru|yandex|proton|icloud|live)\.[a-z.]+', re.I)
CORP = re.compile(r'\b(llc|ltd|limited|inc|corp|co\b|co\.|company|gmbh|s\.?a\.?|b\.?v\.?|pte|sdn|bhd|fze|fzco|fzc|dmcc|plc|group|trading|energy|petroleum|oil|international|resources|enterprises?|holdings?|global|services|industries|llp|too|ag|srl|s\.r\.l|spa|ooo|jsc)\b', re.I)

FLAG_RULES = [
    (r'\bicpo\b', 'ICPO первым шагом'),
    (r'passport', 'запрос паспорта'),
    (r'\bpof\b|proof of fund', 'запрос POF (proof of funds)'),
    (r'\bsblc\b|mt ?760|mt ?799|mt ?199|\bbcl\b|\brwa\b', 'банковские инструменты SBLC/MT799/MT760/BCL/RWA'),
    (r'dip\s*(&|and)?\s*pay|dip test', 'схема Dip & Pay / dip test'),
    (r'\btsa\b|tank storage agreement|storage receipt|\bator\b|\bdtc\b|\bpop\b', 'TSA/ATOR/DTC/POP-документы как «доказательство»'),
    (r'injection', 'схема «injection» в танк покупателя'),
    (r'commission|gross\s*/\s*net|\bgross\b.*\bnet\b|intermediar', 'комиссии посредникам (gross/net)'),
    (r'mandate', 'заявлен «мандат»'),
    (r'jp\s*-?54|\bd2\b.*\bd6\b|\bd6\b|\blng\b|\bespo\b|urea|sulph?ur granul|rebco|bonny', 'универсальный набор «JP54/D6/LNG/Urea…»'),
    (r'rotterdam.*(houston|fujairah|jurong)|(houston|fujairah|jurong).*rotterdam', 'типовой шаблон «Rotterdam/Houston/Fujairah/Jurong»'),
    (r'russia|russian|novorossiysk|tuapse|primorsk|ust[\s-]?luga', 'происхождение РФ: санкции и запрет экспорта для производителей до 31.01.2027'),
    (r'kazakh|atyrau|pavlodar|shymkent|aktau', 'происхождение РК: экспорт дизтоплива запрещён'),
    (r'no upfront|no advance', 'акцент «без предоплаты» (часто прикрытие сборов позже)'),
    (r'whatsapp|telegram', 'основной канал — WhatsApp/Telegram'),
]

def price_min_usd_mt(p):
    if not p: return None
    nums = [float(v.replace(',', '')) for v in re.findall(r'\$\s*([\d,]+(?:\.\d+)?)', p)]
    if not nums: return None
    v = min(nums)
    pl = p.lower()
    if 'barrel' in pl or 'bbl' in pl: v *= 7.45
    elif 'gallon' in pl: v *= 314
    elif 'liter' in pl or 'litre' in pl: v *= 1190
    elif 'kilogram' in pl or '/ kg' in pl: v *= 1000
    return v

def shingles(t, n=6):
    w = re.findall(r'[a-z0-9]+', t.lower())
    return {' '.join(w[i:i+n]) for i in range(max(0, len(w)-n+1))}

# ---------------------------------------------------------------- go4wb
g = json.load(open(os.path.join(DATA, 'g4wb_enriched.json')))
rows, excluded = [], []
for x in g:
    text = ' '.join([x['title'], x.get('desc', '')])
    item = dict(
        company=x['company'], country=x['location'].split(',')[-1].strip(), city=x['location'],
        contact=(x.get('contact') or '') + (f" ({x['designation']})" if x.get('designation') else ''),
        phone=x.get('phone', ''), email='; '.join(sorted({m.group(0) for m in re.finditer(r'[\w.\-]+@[\w\-]+\.[\w.\-]+', x.get('desc', ''))})),
        website=x.get('website', ''), price_raw=x.get('p_Price') or x['price'].replace('\n', ' '),
        basis=x.get('p_Incoterms', ''), volume=('MOQ: ' + (x.get('p_MOQ') or x['moq'].replace('\n', ' '))) + (f"; мощность: {x['p_Capacity']}" if x.get('p_Capacity') else ''),
        origin=x.get('p_Origin', ''), url=x['url'], date=iso(x['date']), wording=(x['title'] + ' — ' + x.get('desc', ''))[:600],
        source='go4WorldBusiness.com', member=x['member'], member_since=x.get('member_since', ''), established=x.get('established', ''),
        types=x.get('types', ''), supplies=x.get('supplies', ''), member_url=x['member_url'], desc=x.get('desc', ''),
        src_kind='страница открыта',
    )
    if not item['contact']:
        m = re.search(r'Contact Person: \| ([^|]+)', x.get('desc', ''))
        if m: item['contact'] = m.group(1).strip()
    if not item['phone']:
        m = re.search(r'Phone: \| ([^|]+)', x.get('desc', ''))
        if m and 'Not displayed' not in m.group(1): item['phone'] = m.group(1).strip()
    (rows if TEN.search(text) else excluded).append(item)

# ---------------------------------------------------------------- manual
for m in LISTINGS:
    rows.append(dict(company=m['company'], country=m['country'], city='', contact=m['contact'], phone=m['phone'], email=m['email'],
        website=m['website'], price_raw=m['price'], basis=m['basis'], volume=m['volume'], origin=m['origin'], url=m['url'],
        date=m['date'], wording=m['wording'], source=re.sub(r'^https?://(www\.)?([^/]+).*', r'\2', m['url'].split(' ;')[0]),
        member='', member_since='', established='', types=m['seller_type'], supplies='', member_url='', desc=m['wording'],
        manual_flags=m['flags'], manual_price=m['price_min'], src_kind=m['src_kind']))

# ---------------------------------------------------------------- duplicates (одинаковый текст у разных компаний)
def core(t):
    t = re.split(r'Full Specifications|Send Inquiry|Member Details|Product Name \|', t)[0]
    # убираем строки-спецификации (у всех EN590 одинаковые)
    return re.sub(r'(sulph?ur|cetane|density|flash|viscosity|cfpp|water|ash|lubricity|distillation|carbon residue|oxidation|copper|fame|polycyclic)[^|]*', ' ', t, flags=re.I)
sh = [shingles(core(r['desc'])) for r in rows]
dup = defaultdict(set)
for i in range(len(rows)):
    if len(sh[i]) < 25: continue
    for j in range(i+1, len(rows)):
        if len(sh[j]) < 25 or rows[i]['company'] == rows[j]['company']: continue
        inter = len(sh[i] & sh[j]);
        if inter and inter / min(len(sh[i]), len(sh[j])) > 0.45:
            dup[i].add(rows[j]['company']); dup[j].add(rows[i]['company'])

# ---------------------------------------------------------------- scoring
for i, r in enumerate(rows):
    t = (r['wording'] + ' ' + r['desc'] + ' ' + r['origin'] + ' ' + r['supplies']).lower()
    flags = [label for pat, label in FLAG_RULES if re.search(pat, t, re.I)]
    pm = r.get('manual_price') if 'manual_price' in r else price_min_usd_mt(r['price_raw'])
    r['price_usd_mt'] = round(pm, 0) if pm else None
    strong = 0
    if pm:
        ratio = pm / BENCH
        r['ratio'] = round(ratio, 2)
        if pm < 50:
            flags.insert(0, 'цена-заглушка ($1–$50/т)'); strong += 1
        elif ratio > 2:
            flags.insert(0, f'цена {ratio:.0%} от ICE Gasoil — заглушка/ошибка единиц'); strong += 1
        elif ratio < 0.8:
            flags.insert(0, f'цена {ratio:.0%} от ICE Gasoil (${BENCH:.0f}/т) — нереалистичная скидка'); strong += 1
    else:
        r['ratio'] = None
    if FREEMAIL.search(r['desc'] + ' ' + r['email']): flags.append('бесплатная почта (gmail/yahoo/…)')
    if r['established'] and r['established'].isdigit() and int(r['established']) >= 2023: flags.append(f"компания основана в {r['established']} г.")
    if r['source'] == 'go4WorldBusiness.com' and not CORP.search(r['company']): flags.append('название продавца похоже на ФИО, не на юрлицо')
    if len(dup.get(i, ())) >= 2: strong += 1
    if dup.get(i): flags.append('тот же текст у других компаний: ' + ', '.join(sorted(dup[i]))[:200])
    if re.search(r'\b(refiner(y|ies) direct|direct (from )?refiner|own refinery|our refinery)\b', t): flags.append('заявляет «прямо с НПЗ» без названия НПЗ')
    if r.get('manual_flags'): flags.append(r['manual_flags'])
    if r['date'][:4] < '2025' and r['date'][:1].isdigit(): flags.append('объявление старше 2025 г.')
    for k in ('происхождение РФ', 'происхождение РК', 'ICPO', 'паспорт', 'POF', 'SBLC', 'TSA', 'injection', 'Dip'):
        if any(f.startswith(k) or k in f for f in flags): strong += 0.5
    r['flags'] = '; '.join(dict.fromkeys(flags))
    r['n_flags'] = len(flags)
    # ---- надёжность
    has_site = bool(r['website'])
    old = r['established'].isdigit() and int(r['established']) <= 2018 if r['established'] else False
    if strong < 1 and has_site and old and len(flags) <= 3:
        r['rating'] = 'Средний'
    else:
        r['rating'] = 'Низкий'
    tl = r['types'].lower()
    if r['rating'] == 'Средний' and ('exporter' in tl or 'trad' in tl):
        r['group'] = 'C — трейдер (полномочия не подтверждены)'
    else:
        r['group'] = 'D — брокер/посредник или не установлено'
    r['status'] = (f"Компания на площадке: {r['member'] or '—'}; на площадке с: {r['member_since'] or '—'}; основана: {r['established'] or '—'}; "
                   f"сайт: {'есть' if has_site else 'нет'}; наличие топлива: НЕ подтверждено; {r['src_kind']}")

rows.sort(key=lambda r: (r['rating'] != 'Средний', r['n_flags'], r['date'] or ''), reverse=False)

# ---------------------------------------------------------------- workbook
wb = Workbook()
HDR = PatternFill('solid', fgColor='1F3864'); HF = Font(color='FFFFFF', bold=True)
FILL = {'Высокий': 'C6EFCE', 'Средний': 'FFEB9C', 'Низкий': 'F8CBAD'}

def sheet(name, headers, data, widths=None, rating_col=None):
    ws = wb.create_sheet(name[:31])
    ws.append(headers)
    for c in ws[1]: c.fill = HDR; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical='top')
    for d in data:
        ws.append([('' if v is None else v) for v in d])
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(h, 18)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
            if isinstance(c.value, str) and c.value.startswith('http'):
                c.hyperlink = c.value.split(' ;')[0].strip(); c.font = Font(color='0563C1', underline='single')
        if rating_col is not None:
            v = row[rating_col].value
            if v in FILL: row[rating_col].fill = PatternFill('solid', fgColor=FILL[v])
    ws.freeze_panes = 'C2'; ws.auto_filter.ref = ws.dimensions
    return ws

MAIN_H = ['№', 'Компания', 'Страна', 'Контактное лицо', 'Телефон', 'Email', 'Сайт', 'Цена за MT (как в объявлении)', 'Цена мин., USD/MT',
          '% от ICE Gasoil', 'Условия поставки', 'Объём', 'Заявл. происхождение', 'Источник объявления (прямая ссылка)', 'Дата публикации/обновления',
          'Точная формулировка', 'Тип продавца (заявлено)', 'Группа A/B/C/D', 'Рейтинг достоверности', 'Статус проверки', 'Признаки риска', 'Площадка']
W = {'Компания': 28, 'Точная формулировка': 60, 'Признаки риска': 60, 'Статус проверки': 45, 'Источник объявления (прямая ссылка)': 45,
     'Сайт': 28, 'Объём': 24, 'Условия поставки': 20, '№': 5}
def mainrow(n, r):
    return [n, r['company'], r['country'], r['contact'], r['phone'], r['email'], r['website'], r['price_raw'], r['price_usd_mt'],
            (f"{r['ratio']:.0%}" if r['ratio'] else ''), r['basis'], r['volume'], r['origin'], r['url'], r['date'], r['wording'],
            r['types'], r['group'], r['rating'], r['status'], r['flags'], r['source']]
RC = MAIN_H.index('Рейтинг достоверности')

# 0. README
ws = wb.active; ws.title = 'Сводка'
cnt_mid = sum(r['rating'] == 'Средний' for r in rows)
priced = [r for r in rows if r['price_usd_mt'] and r['price_usd_mt'] >= 50 and r['ratio'] <= 2]
below = sum(1 for r in priced if r['ratio'] < 0.8)
summary = [
    ['Глобальный поиск продавцов дизельного топлива EN 590 10 ppm', ''],
    ['Дата проверки', CHECK_DATE],
    ['Бенчмарк цены', f'ICE Low Sulphur Gasoil ≈ ${BENCH:.0f}/MT (09.11.2026), источник: {BENCH_SRC}. ФТ/ING: в сентябре 2026 крэк дизеля в Европе >$100/bbl, ICE Gasoil ≈ $1 423/т.'],
    ['Всего объявлений 10 ppm (2025–2026 + сторонние источники)', len(rows)],
    ['Из них с опубликованной ценой', len(priced)],
    ['Из них с ценой < 80% рынка', f'{below} из {len(priced)}'],
    ['Рейтинг «Средний» среди B2B-объявлений', cnt_mid],
    ['Рейтинг «Высокий» среди B2B-объявлений', '0 — ни одно объявление не подтверждено независимо'],
    ['Ключевой вывод', 'Практически все публичные объявления «EN590 10ppm» на B2B-площадках — посредники с ценами в 2–4 раза ниже рынка и типовыми процедурами (ICPO → TSA/POP → injection/Dip&Pay). '
     'Реальный физический 10 ppm продаётся НПЗ и крупными трейдерами через срочные контракты и sell-тендеры; см. лист «Прямые продавцы» и «Для офиц. запроса».'],
    ['Ограничения экспорта (Казахстан и соседи)', 'РК: запрет экспорта дизтоплива авто/ж.д. 22.11.2026–22.05.2027 и вне ЕАЭС 01.01–30.06.2027 (Interfax, Qazinform). КМГ официально предупреждает о ложных офертах от имени своих НПЗ. '
     'РФ: запрет экспорта дизтоплива для производителей продлён до конца января 2027 с частичными исключениями; для перепродавцов — запрет; плюс санкционные риски. '
     'Азербайджан (SOCAR, ~2,3 млн т Евро-5 в 2026) и Туркменистан (биржа) — реальные легальные варианты в регионе.'],
    ['Как пользоваться', '1) Лист «Для офиц. запроса» — кому писать в первую очередь. 2) «ТОП-20». 3) «Все предложения» — полный реестр с фильтрами. 4) «Подозрительные» — разбор схем.'],
    ['Методика', 'go4WorldBusiness: 10 поисковых запросов × до 3 страниц → 427 уникальных объявлений → 267 за 2025–2026 → открыты карточка товара и профиль продавца (год основания, телефон, контакт, адрес) → '
     'фильтр «10 ppm / ULSD / 10 mg/kg» → флаги риска по тексту, сравнение цены с ICE Gasoil, поиск одинаковых текстов (6-словные шинглы, порог 50%). Остальные площадки — открытые страницы или сниппеты поиска (помечено).'],
    ['Правило рейтинга', 'Высокий — компания подтверждена И есть независимое подтверждение предложения (тендер в отраслевой прессе / офиц. сайт). Средний — компания существует (сайт, основана ≤2018), нет сильных флагов, но наличие продукта не подтверждено. Низкий — всё прочее.'],
    ['Важно', 'SGS/ICPO/POP/TSA/ATV не являются доказательством наличия топлива. Никаких предоплат, «депозитов за dip test», TSA-сборов неизвестным компаниям. Оплата — только документарный аккредитив / после независимой инспекции на терминале, проверенной напрямую у оператора терминала.'],
]
for s in summary: ws.append(s)
ws.column_dimensions['A'].width = 45; ws.column_dimensions['B'].width = 140
for row in ws.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')
ws['A1'].font = Font(bold=True, size=14)

# 1. Все предложения
sheet('Все предложения', MAIN_H, [mainrow(i+1, r) for i, r in enumerate(rows)], W, RC)

# 2. Прямые продавцы (A/B)
OFF_H = ['№', 'Группа', 'Компания', 'Страна', 'Официальный сайт', 'Что подтверждает статус продавца', 'Источник подтверждения', 'Базис', 'Примечание', 'Рейтинг достоверности']
off_rows = [[i+1, ('A — производитель/НПЗ' if o[0] == 'A' else 'B — физический поставщик/крупный трейдер'), o[1], o[2], o[3], o[4], o[5], o[6], o[7],
             ('Средний' if ('Dangote' in o[1] or 'Bukhara' in o[1] or o[5].rstrip('/') in o[3]) else 'Высокий')] for i, o in enumerate(OFFICIAL)]
# Высокий — есть конкретное независимое подтверждение продаж (тендер/офиц. страница); Средний — компания бесспорна, но конкретная оферта 10 ppm не найдена
sheet('Прямые продавцы (A-B)', OFF_H, off_rows, {'Компания': 34, 'Что подтверждает статус продавца': 60, 'Источник подтверждения': 45, 'Примечание': 40, 'Официальный сайт': 32}, 9)

# 3. Для официального запроса
req = [o for o in off_rows if 'Казахстан' not in o[3] and 'Dangote' not in o[2]]
req_b2b = [r for r in rows if r['rating'] == 'Средний']
REQ_H = ['№', 'Компания', 'Страна', 'Сайт / ссылка', 'Почему', 'Что запросить', 'Рейтинг достоверности']
ask = 'Запрос на срочный контракт/спот-тендер ULSD 10 ppm: объём, базис, формула цены (Platts/Argus + премия), требования KYC, форма оплаты (LC)'
reqd = [[i+1, o[2], o[3], o[4], o[5], ask, o[9]] for i, o in enumerate(req)]
reqd += [[len(reqd)+i+1, r['company'], r['country'], r['website'] or r['url'], 'B2B-объявление без сильных флагов; компания с сайтом и историей',
          'Запросить: название НПЗ/терминала, номер танка, письмо оператора терминала напрямую, регистрационные документы, банк-референс; без предоплат', r['rating']] for i, r in enumerate(req_b2b)]
sheet('Для офиц. запроса', REQ_H, reqd, {'Компания': 34, 'Почему': 60, 'Что запросить': 60, 'Сайт / ссылка': 40}, 6)

# 4. С ценами
sheet('С ценами', MAIN_H, [mainrow(i+1, r) for i, r in enumerate(sorted(priced, key=lambda r: r['price_usd_mt']))], W, RC)

# 5. CIF
cif = [r for r in rows if re.search(r'\bc(if|fr|nf)\b', (r['basis'] + ' ' + r['wording'] + ' ' + r['desc']), re.I)]
sheet('CIF-CFR', MAIN_H, [mainrow(i+1, r) for i, r in enumerate(cif)], W, RC)

# 6. Казахстан и соседи
NEI = r'kazakh|казах|russia|росси|azerbaij|азерб|turkmen|туркм|uzbek|узбек|kyrgyz|кыргыз|tajik|тадж|georgia|грузи|aktau|atyrau|baku|kulevi|batumi|novoross|caspian'
kz = [r for r in rows if re.search(NEI, ' '.join([r['country'], r['origin'], r['wording'], r['desc'], r['city']]), re.I)]
ws = sheet('Казахстан и соседи', MAIN_H, [mainrow(i+1, r) for i, r in enumerate(kz)], W, RC)
ws.append([]); ws.append(['', 'ОФИЦИАЛЬНЫЕ ПРОИЗВОДИТЕЛИ И КАНАЛЫ РЕГИОНА'])
for o in off_rows:
    if re.search(NEI, o[3] + o[2], re.I):
        ws.append(['', o[2], o[3], '', '', '', o[4], '', '', '', o[7], '', '', o[6], '', o[5], o[1], o[1], o[9], o[8]])

# 7. Подозрительные
sus = sorted([r for r in rows if r['rating'] == 'Низкий'], key=lambda r: -r['n_flags'])
SUS_H = ['№', 'Компания', 'Страна', 'Цена мин., USD/MT', '% от ICE Gasoil', 'Ссылка', 'Дата', 'Кол-во флагов', 'Признаки риска', 'Телефон', 'Email']
sheet('Подозрительные', SUS_H, [[i+1, r['company'], r['country'], r['price_usd_mt'], (f"{r['ratio']:.0%}" if r['ratio'] else ''), r['url'], r['date'], r['n_flags'], r['flags'], r['phone'], r['email']]
                               for i, r in enumerate(sus)], {'Компания': 30, 'Признаки риска': 90, 'Ссылка': 45})
# схемы
ws = wb.create_sheet('Схемы мошенничества')
for line in [
    ['Схема', 'Как выглядит', 'Почему опасно', 'Как проверить'],
    ['Цена ниже рынка', 'EN590 10ppm за $340–600/т при ICE Gasoil ≈ $1 346/т', 'Ни один НПЗ не продаёт с дисконтом 50–75%; это приманка', 'Сравнить с Platts/Argus/ICE на дату оферты'],
    ['ICPO → SCO → TSA/ATOR → injection/Dip&Pay', 'Покупатель «должен» арендовать танк, оплатить «dip test», «TSA», «injection fee»', 'Расходы покупателя до проверки товара = advance-fee fraud', 'Связаться с оператором терминала (Vopak, Oiltanking и т.п.) напрямую по его официальному сайту'],
    ['Запрос паспорта и POF до переговоров', '«Buyer passport and POF are required for Seller verification»', 'Сбор персональных данных и выписок для других мошенничеств', 'Не отправлять до проверки продавца'],
    ['Фиктивный НПЗ / «мандат НПЗ»', '«KG Refinery», «our seller refinery», «refinery-aligned mandate» без названия', 'НПЗ не выдают мандаты брокерам на B2B-площадках', 'Запросить название НПЗ и проверить через его официальный сайт/пресс-службу'],
    ['Оферты от имени НПЗ КМГ / «Kazakhstan origin»', 'EN590 из Казахстана, FOB Aktau/Atyrau, CIF Rotterdam', 'Экспорт дизтоплива из РК запрещён; КМГ публично предупреждает о ложных офертах', 'https://www.kmg.kz/en/sales-petroleum/'],
    ['Одинаковые тексты', 'Один шаблон у десятков «компаний» / LinkedIn-профилей', 'Признак брокерских цепочек и копипаста', 'См. флаг «тот же текст у других компаний»'],
    ['«Russian origin / no sanctions»', 'Российский ДТ через третьи страны', 'Санкции ЕС/США/Великобритании; запрет экспорта РФ до 31.01.2027', 'Комплаенс-проверка'],
    ['Портфель «JP54, D6, LNG, Urea, Sugar»', 'Один продавец «имеет» всё', 'Реальные физические трейдеры не публикуют такие списки на досках объявлений', '—'],
]: ws.append(line)
for col, w in zip('ABCD', (32, 60, 60, 50)): ws.column_dimensions[col].width = w
for row in ws.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')
for c in ws[1]: c.fill = HDR; c.font = HF

# 8. ТОП-20
top = sorted([o for o in off_rows if 'Казахстан' not in o[3] and 'Dangote' not in o[2] and 'Bukhara' not in o[2]], key=lambda o: o[9] != 'Высокий')[:20 - min(5, len(req_b2b))]
TOP_H = ['№', 'Компания', 'Страна', 'Группа', 'Сайт / ссылка', 'Почему в ТОП', 'Рейтинг достоверности']
topd = [[i+1, o[2], o[3], o[1], o[4], o[5], o[9]] for i, o in enumerate(top)]
for r in sorted(req_b2b, key=lambda r: r['n_flags'])[:20 - len(topd)]:
    topd.append([len(topd)+1, r['company'], r['country'], r['group'], r['website'] or r['url'], 'Лучший B2B-лид по совокупности проверок (сайт, история, без сильных флагов). Наличие НЕ подтверждено.', r['rating']])
sheet('ТОП-20', TOP_H, topd, {'Компания': 36, 'Почему в ТОП': 70, 'Сайт / ссылка': 40, 'Группа': 30}, 6)

# 9. Исключены (EN590 без явного 10ppm)
sheet('EN590 без явного 10ppm', ['№', 'Компания', 'Страна', 'Заголовок/текст', 'Ссылка', 'Дата', 'Цена'],
      [[i+1, r['company'], r['country'], r['wording'][:300], r['url'], r['date'], r['price_raw']] for i, r in enumerate(excluded)], {'Заголовок/текст': 70, 'Ссылка': 45})

# 10. Архив до 2025
arch = [x for x in json.load(open(os.path.join(DATA, 'g4wb_list.json'))) if x['date'][-2:] not in ('25', '26')]
sheet('Архив до 2025', ['№', 'Компания', 'Локация', 'Заголовок', 'Цена', 'Ссылка', 'Дата'],
      [[i+1, x['company'], x['location'], x['title'], x['price'].replace('\n', ' '), x['url'], iso(x['date'])] for i, x in enumerate(arch)], {'Ссылка': 45})

# 11. Источники и ограничения
ws = sheet('Источники и ограничения', ['Источник', 'Статус'], [list(b) for b in BLOCKED] + [
    ['go4worldbusiness.com', 'открыт: поиск + карточки + профили продавцов'],
    ['franceajgroup.com, toocrudeoil.com, groups.google.com, e-shipsupply.com, kmg.kz', 'открыты'],
    ['Поисковые запросы (EN/RU/ZH)', 'EN590 10ppm for sale FOB Rotterdam; EN590 export offer refinery seller; ULSD 10ppm Fujairah Jebel Ali; EN590 Malaysia Singapore Jurong TTT; EN590 Turkey Mersin Ceyhan CIF; EN590 Oman Sohar Saudi; '
     'дизельное топливо EN590 10 ppm экспорт поставщик Казахстан; EN590 10ppm 柴油出售 出口报价; refinery sell tender 10ppm gasoil 2026; Telegram/LinkedIn EN590 seller; экспортные ограничения РК/РФ; SOCAR, TCOR, Bukhara, KMG, MRPL, Reliance, GS Caltex, FPCC'],
], {'Источник': 50, 'Статус': 120})

wb.save(OUT)
print('rows', len(rows), 'excluded', len(excluded), 'priced', len(priced), 'below80', below, 'mid', cnt_mid, 'cif', len(cif), 'kz', len(kz))
