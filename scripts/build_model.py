# -*- coding: utf-8 -*-
"""Builds UAL_model.xlsx with live formulas. Inputs = blue font, formulas = black."""
import json, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
SCR = sys.argv[2]
BLUE = Font(name="Arial", size=9, color="0000FF")
BLACK = Font(name="Arial", size=9, color="000000")
BOLD = Font(name="Arial", size=9, bold=True)
HDR = Font(name="Arial", size=9, bold=True, color="FFFFFF")
TITLE = Font(name="Arial", size=12, bold=True, color="1F3864")
NOTE = Font(name="Arial", size=8, italic=True, color="595959")
HFILL = PatternFill("solid", fgColor="1F3864")
SFILL = PatternFill("solid", fgColor="D9E1F2")
KFILL = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF")

wb = Workbook()
REG = {}  # key -> {col: "Sheet!A1"}

class S:
    def __init__(s, name, title, cols, first_col=3, widths=None):
        s.ws = wb.create_sheet(name) if wb.sheetnames != ["Sheet"] else wb.active
        if s.ws.title == "Sheet": s.ws.title = name
        s.name = name; s.cols = cols; s.fc = first_col; s.r = 1
        s.ws["A1"] = title; s.ws["A1"].font = TITLE; s.r = 3
        s.ws.column_dimensions["A"].width = 46
        s.ws.column_dimensions["B"].width = 9
        for i in range(len(cols)):
            s.ws.column_dimensions[L(first_col + i)].width = widths or 11
        s.srccol = first_col + len(cols)
        s.ws.column_dimensions[L(s.srccol)].width = 70
        s.ws.sheet_view.showGridLines = False
        s.ws.freeze_panes = "C4" if cols else None
    def col(s, c): return L(s.fc + s.cols.index(c))
    def header(s, label="", unit="", src="Источник / тип"):
        ws = s.ws
        ws.cell(s.r, 1, label); ws.cell(s.r, 2, unit)
        for i, c in enumerate(s.cols): ws.cell(s.r, s.fc + i, c)
        ws.cell(s.r, s.srccol, src)
        for j in range(1, s.srccol + 1):
            x = ws.cell(s.r, j); x.font = HDR; x.fill = HFILL; x.alignment = Alignment(horizontal="center" if j > 2 else "left")
        s.r += 1
    def section(s, text):
        ws = s.ws; ws.cell(s.r, 1, text).font = BOLD
        for j in range(1, s.srccol + 1): ws.cell(s.r, j).fill = SFILL
        s.r += 1
    def note(s, text):
        s.ws.cell(s.r, 1, text).font = NOTE; s.r += 1
    def row(s, key, label, unit, vals, fmt="#,##0", src="", inp=None, bold=False, hl=False):
        """vals: dict col->value or formula string (starting with '=') or callable(col)->str"""
        ws = s.ws; r = s.r
        ws.cell(r, 1, label).font = BOLD if bold else BLACK
        ws.cell(r, 2, unit).font = NOTE
        REG.setdefault(key, {})
        for c in s.cols:
            v = vals(c) if callable(vals) else (vals.get(c) if isinstance(vals, dict) else None)
            if v is None: continue
            cell = ws.cell(r, s.fc + s.cols.index(c), v)
            isf = isinstance(v, str) and v.startswith("=")
            cell.font = Font(name="Arial", size=9, color=("000000" if isf else "0000FF"), bold=bold)
            if isinstance(v, str) and not isf: cell.alignment = Alignment(horizontal="right")
            cell.number_format = fmt
            if hl: cell.fill = KFILL
            REG[key][c] = f"'{s.name}'!${s.col(c)}${r}"
        if src: ws.cell(r, s.srccol, src).font = NOTE
        s.r += 1
        return r
    def blank(s, n=1): s.r += n

def R(key, c):  # reference to registered cell
    return REG[key][c]

# ============================================================== README
rd = S("README", "United Airlines Holdings (UAL) — модель оценки", [], first_col=3)
rd.ws.column_dimensions["A"].width = 120
lines = [
 ("Дата анализа: 14.10.2026. Цена акции: $103,23 — закрытие NASDAQ 13.10.2026 (Yahoo Finance chart API; совпадает с отчётом Zacks/Yahoo от 13.10.2026).", BLACK),
 ("Последняя опубликованная отчётность: 10-Q за 2 кв. 2026 (подан в SEC 16.07.2026) и 8-K с пресс-релизом от 15.07.2026.", BLACK),
 ("Отчёт за 3 кв. 2026 на дату анализа НЕ вышел: компания назначила публикацию на 20.10.2026 после закрытия рынка (звонок 21.10.2026). Данные 3 кв. в модели не используются.", BLACK),
 ("Пиры: TTM на 30.06.2026 (последний общий период для всех). Delta уже отчиталась за 3 кв. (09.10.2026), но для сопоставимости берём её данные на 30.06.2026.", BLACK),
 ("", BLACK),
 ("Цвета: синий шрифт — вводные значения (факты из отчётности и допущения); чёрный — формулы. Жёлтая заливка — ключевые результаты.", BOLD),
 ("Типы данных в колонке источника: [Ф] факт из отчётности; [Д] допущение аналитика; [О] оценка аналитика (расчёт на основе фактов).", BOLD),
 ("Все финансовые показатели — млн долл. США, если не указано иное. Удельные показатели (TRASM, CASM) — центы на ASM.", BLACK),
 ("", BLACK),
 ("Листы:", BOLD),
 ("Hist — исходные данные 2019–2025, 1П25, 1П26, TTM 2К26 и квартальные данные (факты с указанием документа SEC).", BLACK),
 ("Metrics — производные метрики (формулы): загрузка, доходность, маржи, качество прибыли, долговая нагрузка.", BLACK),
 ("Assumptions — ВСЕ допущения прогноза в одном месте: рынок, WACC, драйверы по трём сценариям, терминальные параметры, вероятности.", BLACK),
 ("Beta — расчёт беты по месячным доходностям UAL и S&P 500 за 60 месяцев (функция SLOPE).", BLACK),
 ("WACC — расчёт ставки дисконтирования с учётом аренды в долге.", BLACK),
 ("DCF — FCFF по трём сценариям (2026–2032), терминальная стоимость двумя способами.", BLACK),
 ("ReverseDCF — какая поправка к росту RASM и какой WACC заложены в текущую цену.", BLACK),
 ("Multiples — пиры (США) и собственная история мультипликаторов UAL; международные пиры только для контекста.", BLACK),
 ("Scenarios — взвешенная по вероятностям целевая цена и сводка методов.", BLACK),
 ("Sensitivity — WACC × g; цена топлива × TRASM; ±10% топливо и ±1% RASM.", BLACK),
 ("Debt — график погашения, ликвидность, рейтинги, сравнение долговой нагрузки через EBITDAR и через операционный денежный поток, поставки самолётов.", BLACK),
]
r = 3
for t, f in lines:
    rd.ws.cell(r, 1, t).font = f; r += 1

# ============================================================== HIST
YC = ["2019", "2020", "2021", "2022", "2023", "2024", "2025", "1П25", "1П26", "TTM 2К26"]
h = S("Hist", "Исходные данные UAL (факты из отчётности SEC)", YC)
FY = ["2019", "2020", "2021", "2022", "2023", "2024", "2025"]
def vals(lst, ttm="flow"):
    d = {c: v for c, v in zip(YC[:9], lst) if v is not None}
    return d
def add(key, label, unit, lst, src, fmt="#,##0", ttm="flow"):
    d = vals(lst)
    if ttm == "flow" and all(c in d for c in ("2025", "1П25", "1П26")):
        d["TTM 2К26"] = None
    rr = h.r
    def f(c):
        if c in d and c != "TTM 2К26": return d[c]
        if c == "TTM 2К26":
            if ttm == "flow" and all(k in d for k in ("2025", "1П25", "1П26")):
                return f"={h.col('2025')}{rr}+{h.col('1П26')}{rr}-{h.col('1П25')}{rr}"
            if ttm == "stock" and "1П26" in d:
                return f"={h.col('1П26')}{rr}"
        return None
    return h.row(key, label, unit, f, fmt=fmt, src=src)

SRC_Q4 = "[Ф] Пресс-релизы 8-K по итогам 4 кв.: FY2019–2021 (8-K 19.01.2022), FY2022 (17.01.2023), FY2023 (22.01.2024), FY2024 (21.01.2025), FY2025 (20.01.2026); 1П — 8-K 15.07.2026"
h.header("Показатель", "Ед.")
h.section("Операционные показатели")
add("asm", "Пассажиро-кресло-мили (ASM)", "млн", [284999, 122804, 178684, 247858, 291333, 311185, 330284, 159503, 164977], SRC_Q4)
add("rpm", "Пассажиро-мили (RPM)", "млн", [239360, 73883, 128979, 206791, 244435, 258503, 271619, 129604, 136150], SRC_Q4)
add("pax", "Пассажиры", "тыс.", [162443, 57761, 104082, 144300, 164927, 173603, 181053, 86992, 91178], SRC_Q4)
add("casmex_rep", "CASM-ex (отчётный, non-GAAP)", "центы", [10.21, 17.13, 12.96, 11.73, 12.03, 12.58, 12.64, 12.74, 13.51], SRC_Q4 + "; без топлива, участия в прибыли, спецстатей и сторонних услуг", fmt="0.00", ttm=None)
add("fuelpx", "Средняя цена топлива", "$/галлон", [2.09, 1.57, 2.11, 3.63, 3.01, 2.65, 2.44, 2.43, 3.53], SRC_Q4, fmt="0.00", ttm=None)
add("gal", "Потреблено топлива", "млн галл.", [4292, 2004, 2729, 3608, 4205, 4444, 4663, 2254, 2312], SRC_Q4)
add("fleet", "Самолётов в парке на конец периода", "шт.", [1372, 1287, 1344, 1338, 1358, 1406, 1490, 1473, 1552], SRC_Q4, ttm="stock")
add("emp", "Численность персонала на конец периода", "тыс.", [95.9, 74.4, 84.1, 92.8, 103.3, 107.3, 113.2, 111.3, 117.5], SRC_Q4, fmt="0.0", ttm="stock")
add("stage", "Средняя длина рейса", "миль", [1460, 1307, 1315, 1437, 1479, 1490, 1488, 1482, 1462], SRC_Q4, ttm=None)
add("rtm", "Грузовые тонно-мили (RTM)", "млн", [3329, 2711, 3285, 3041, 3159, 3604, 3626, 1774, 1810], SRC_Q4)
h.section("Отчёт о прибылях и убытках")
add("rev_pax", "Пассажирская выручка", "$ млн", [39625, 11805, 20197, 40032, 49046, 51829, 53438, 25696, 29267], SRC_Q4)
add("rev_cargo", "Грузовая выручка", "$ млн", [1179, 1648, 2349, 2171, 1495, 1743, 1779, 859, 949], SRC_Q4 + "; 10-K 2025 (Cargo 1,779/1,743)")
add("rev_oth", "Прочая операционная выручка", "$ млн", [2455, 1902, 2088, 2752, 3176, 3491, 3853, 1893, 2064], SRC_Q4)
rr = h.r
h.row("rev", "Итого операционная выручка", "$ млн", lambda c: f"=SUM({h.col(c)}{rr-3}:{h.col(c)}{rr-1})", src="формула; сверка с отчётной выручкой ниже", bold=True)
add("rev_rep", "Справочно: выручка по отчётности", "$ млн", [43259, 15355, 24634, 44955, 53717, 57063, 59070, 28448, 32280], "[Ф] XBRL RevenueFromContractWithCustomer (10-K/10-Q)")
add("loy_oth", "в т.ч. выручка от партнёров MileagePlus в прочей выручке", "$ млн", [None, None, None, None, 2700, 2900, 3200, 1600, 1800], "[Ф] 10-K 2025 прим. 2 ($3.2/2.9/2.7 млрд); 10-Q 2К26 ($1.8/1.6 млрд за 6 мес.). 2019–2022: не найдено в использованных документах")
add("miles_earned", "Мили начислены (отложенная выручка, прирост)", "$ млн", [None, None, None, None, None, 3563, 3883, None, None], "[Ф] 10-K 2025, roll-forward Frequent flyer deferred revenue")
add("miles_red", "Мили погашены на перелёты (признано в пасс. выручке)", "$ млн", [None, None, None, None, None, 3150, 3414, None, None], "[Ф] 10-K 2025, roll-forward Frequent flyer deferred revenue")
add("fuel", "Расходы на топливо", "$ млн", [8953, 3153, 5755, 13113, 12651, 11756, 11396, 5476, 8150], SRC_Q4)
add("sal", "Зарплата и связанные расходы", "$ млн", [12071, 9522, 9566, 11466, 14787, 16678, 17647, 8568, 9248], SRC_Q4)
add("da", "Амортизация (D&A)", "$ млн", [2288, 2488, 2485, 2456, 2671, 2928, 2939, 1461, 1518], "[Ф] XBRL DepreciationDepletionAndAmortization (10-K/10-Q)")
add("rent", "Аренда самолётов (aircraft rent)", "$ млн", [288, 198, 228, 252, 197, 193, 252, 118, 195], "[Ф] XBRL AircraftRental (10-K/10-Q)")
add("spec", "Специальные статьи (доход) в опер. расходах", "$ млн", [246, -2616, -3367, 140, 949, 112, 259, 340, -534], SRC_Q4 + "; 2020–21 — гос. поддержка PSP (кредит)")
add("ps", "Участие сотрудников в прибыли (profit sharing)", "$ млн", [485, 0, 0, 133, 681, 713, 704, 231, 150], SRC_Q4 + "; 2019 — [О] 0,17 цента × ASM из таблицы CASM-ex")
add("opinc", "Операционная прибыль (GAAP)", "$ млн", [4301, -6359, -1022, 2337, 4211, 5096, 4713, 1932, 2093], "[Ф] XBRL OperatingIncomeLoss")
add("adj_opinc", "Скорр. операционная прибыль (non-GAAP)", "$ млн", [4547, -8975, -4389, 2477, 5160, 5208, 4972, 2272, 1559], SRC_Q4)
add("int_gross", "Процентные расходы (валовые)", "$ млн", [731, 1063, 1657, 1778, 1956, 1629, 1373, 717, 670], SRC_Q4)
add("int_net", "Чистые процентные расходы (за вычетом капитализ. % и % доходов)", "$ млн", [513, 942, 1541, 1375, 947, 676, 556, 288, 273], SRC_Q4 + " (сверка Adj. EBITDA)")
add("tax", "Налог на прибыль", "$ млн", [905, -1753, -593, 253, 769, 1019, 953, 366, 393], "[Ф] XBRL IncomeTaxExpenseBenefit")
add("ni", "Чистая прибыль (GAAP)", "$ млн", [3009, -7069, -1964, 737, 2618, 3149, 3353, 1361, 1504], "[Ф] XBRL NetIncomeLoss")
add("adj_pti", "Скорр. прибыль до налога (non-GAAP)", "$ млн", [4071, -9860, -5809, 1117, 4320, 4607, 4580, 2061, 1342], SRC_Q4)
add("adj_eps", "Скорр. разводнённая прибыль на акцию", "$", [12.05, -27.57, -13.94, 2.52, 10.05, 10.61, 10.62, 4.75, 3.18], SRC_Q4, fmt="0.00")
add("dil_sh", "Разводнённое средневзв. число акций", "млн", [259.9, 279.4, 321.9, 330.1, 331.9, 333.2, 328.5, 330.1, 326.7], SRC_Q4, fmt="0.0", ttm=None)
add("adj_ebitda", "Скорр. EBITDA (non-GAAP, определение UAL)", "$ млн", [6808, -6430, -1783, 4948, 7938, 8211, 8076, 3809, 3133], SRC_Q4)
add("fixlease", "Фиксированная часть операционной аренды (новое определение EBITDAR UAL)", "$ млн", [None, None, None, None, None, 855, 894, 433, 514], "[Ф] 8-K 20.01.2026 и 15.07.2026; ранее не раскрывалось")
h.section("Денежные потоки")
add("cfo", "Операционный денежный поток (CFO)", "$ млн", [6909, -4133, 2067, 6066, 6911, 9445, 8431, 5927, 6409], "[Ф] XBRL NetCashProvidedByOperatingActivities")
add("capex", "Капзатраты (GAAP, за вычетом возврата депозитов)", "$ млн", [4528, 1727, 2107, 4819, 7171, 5615, 5874, 2520, 3015], "[Ф] XBRL PaymentsToAcquireProductiveAssets")
add("adj_capex", "Скорр. капзатраты (вкл. приобретённое в долг/лизинг)", "$ млн", [None, 3449, 2907, 4838, 7948, 6263, 6266, 2468, 3167], SRC_Q4 + "; 2024 — пересчитано в релизе 20.01.2026; 2019 — не найдено")
add("fcf_rep", "FCF по определению UAL (non-GAAP)", "$ млн", [None, -7753, -1508, 1224, -1037, 4270, 2710, 3442, 3225], SRC_Q4 + "; методика изменена в 2025 (2024 пересчитан 3,424→4,270)")
add("buyback", "Выкуп акций", "$ млн", [1645, 353, 0, 0, 0, 162, 637, 589, 27], "[Ф] XBRL PaymentsForRepurchaseOfCommonStock; 10-Q 2К26")
h.section("Баланс (на конец периода)")
add("cash", "Денежные средства и эквиваленты", "$ млн", [2762, 11269, 18283, 7166, 6058, 8769, 5942, 9354, 10166], "[Ф] XBRL CashAndCashEquivalents; 8-K 15.07.2026", ttm="stock")
add("sti", "Краткосрочные инвестиции", "$ млн", [2182, 414, 123, 9248, 8330, 5706, 6298, 6262, 6471], "[Ф] XBRL ShortTermInvestments", ttm="stock")
add("debt", "Долг, финансовая аренда и прочие фин. обязательства", "$ млн", [14818, 27153, 33658, 32280, 31660, 28656, 24988, 27079, 26464], "[Ф] 2022–1П26: 8-K (таблица Adjusted total debt); 2019–2021: [О] XBRL LongTermDebt тек.+долгоср. + FinanceLeaseLiability", ttm="stock")
add("oplease", "Обязательства по операционной аренде", "$ млн", [5632, 5598, 5708, 5020, 5079, 4977, 6048, 5707, 7204], "[Ф] XBRL OperatingLeaseLiability; 8-K", ttm="stock")
add("pension", "Пенсионные и пострабочие обязательства (долгоср.)", "$ млн", [1446, 2460, 1920, 1418, 1605, 1233, 1058, 1199, 1074], "[Ф] 2022+: 8-K; 2019–2021: XBRL DefinedBenefitPensionPlanLiabilitiesNoncurrent (только пенсии — определение уже)", ttm="stock")
add("atl", "Авансы за билеты (air traffic liability)", "$ млн", [None, 4833, 6354, 7555, 6704, 7561, 8131, None, 10752], "[Ф] Балансы в 8-K по итогам 4 кв. и 15.07.2026; 2019 — не найдено", ttm="stock")
add("ffdr", "Отложенная выручка MileagePlus (тек.+долгоср.)", "$ млн", [None, 5975, 6282, 6675, 7143, 7441, 7777, None, 7971], "[Ф] Балансы в 8-K; 10-K 2025 roll-forward", ttm="stock")
add("assets", "Итого активы", "$ млн", [52611, 59548, 68175, 67358, 71104, 74083, 76448, None, 84569], "[Ф] XBRL Assets; 8-K 15.07.2026", ttm="stock")
add("cur_liab", "Итого краткосрочные обязательства", "$ млн", [None, None, None, None, None, None, 26133, None, 27764], "[Ф] 8-K 15.07.2026 (баланс)", ttm="stock")
add("nc_liab", "Итого долгосрочные обязательства", "$ млн", [None, None, None, None, None, None, 35033, None, 40108], "[Ф] 8-K 15.07.2026 (баланс)", ttm="stock")
add("equity", "Собственный капитал", "$ млн", [11531, 5960, 5029, 6896, 9324, 12675, 15282, None, 16697], "[Ф] XBRL StockholdersEquity", ttm="stock")
rr = h.r
h.row("bal_chk", "Проверка баланса: активы − обязательства − капитал (должно быть 0)", "$ млн",
      {c: f"={h.col(c)}{rr-4}-{h.col(c)}{rr-3}-{h.col(c)}{rr-2}-{h.col(c)}{rr-1}" for c in ("2025", "1П26")}, src="формула")
add("sh_out", "Акций в обращении на конец периода", "млн", [251.2, 311.8, 323.8, 326.9, 328.0, 327.9, 323.5, None, 324.6], "[Ф] XBRL CommonStockSharesOutstanding; 1П26 — обложка 10-Q (324 583 772 на 09.07.2026)", fmt="0.0", ttm="stock")
add("px_ye", "Цена акции на конец периода", "$", [88.09, 43.25, 43.78, 37.70, 41.26, 97.10, 111.82, 79.63, 135.99], "[Ф] Yahoo Finance, месячные цены закрытия (декабрь / июнь)", fmt="0.00", ttm=None)
h.blank()
h.note("TTM 2К26 = 2025 + 1П26 − 1П25 для потоков; для остатков — значение на 30.06.2026. Пустая ячейка = данные не найдены в использованных первичных источниках.")

# quarterly block
h.blank()
QC = ["2К25", "3К25", "4К25", "1К26", "2К26"]
qs = S.__new__(S); qs.ws = h.ws; qs.name = "Hist"; qs.cols = QC; qs.fc = 3; qs.r = h.r; qs.srccol = h.srccol
qs.header("Квартальные данные", "Ед.")
SRCQ = "[Ф] 8-K: 16.07.2025, 15.10.2025, 20.01.2026, 21.04.2026, 15.07.2026"
qs.row("q_asm", "ASM", "млн", dict(zip(QC, [84347, 87417, 83365, 77698, 87279])), src=SRCQ)
qs.row("q_rev", "Выручка", "$ млн", dict(zip(QC, [15236, 15225, 15397, 14608, 17672])), src=SRCQ)
qs.row("q_trasm", "TRASM", "центы", dict(zip(QC, [18.06, 17.42, 18.47, 18.80, 20.25])), fmt="0.00", src=SRCQ)
qs.row("q_casmex", "CASM-ex", "центы", dict(zip(QC, [12.36, 12.15, 12.94, 13.95, 13.12])), fmt="0.00", src=SRCQ)
qs.row("q_fuel", "Цена топлива", "$/галл.", dict(zip(QC, [2.34, 2.43, 2.49, 2.78, 4.19])), fmt="0.00", src=SRCQ)
qs.row("q_adjop", "Скорр. операционная маржа", "%", dict(zip(QC, [0.116, 0.087, 0.090, 0.042, 0.054])), fmt="0.0%", src=SRCQ)
qs.row("q_adjpt", "Скорр. маржа до налога", "%", dict(zip(QC, [0.110, 0.080, 0.085, 0.034, 0.048])), fmt="0.0%", src=SRCQ)
qs.row("q_eps", "Скорр. EPS", "$", dict(zip(QC, [3.87, 2.78, 3.10, 1.19, 1.99])), fmt="0.00", src=SRCQ)
qs.row("q_gcjet", "Справочно: спот Gulf Coast jet (среднее за квартал)", "$/галл.", dict(zip(QC, [2.00, 2.12, 2.13, 2.66, 3.64])), fmt="0.00", src="[Ф] FRED MJFUELUSGULF (EIA), среднее месячных значений")
h.r = qs.r

# ============================================================== METRICS
m = S("Metrics", "Производные метрики UAL (формулы на основе листа Hist)", YC)
m.header("Метрика", "Ед.")
def M(key, label, unit, fn, fmt="0.0", cols=YC, src="формула"):
    def f(c):
        if c not in cols: return None
        try: return fn(c)
        except KeyError: return None
    return m.row(key, label, unit, f, fmt=fmt, src=src)
H = lambda k, c: R(k, c)
m.section("Операционные")
M("lf", "Загрузка (load factor) = RPM / ASM", "%", lambda c: f"={H('rpm',c)}/{H('asm',c)}", "0.0%")
M("yield", "Доходность (yield) = пасс. выручка / RPM", "центы", lambda c: f"={H('rev_pax',c)}/{H('rpm',c)}*100", "0.00")
M("prasm", "PRASM = пасс. выручка / ASM", "центы", lambda c: f"={H('rev_pax',c)}/{H('asm',c)}*100", "0.00")
M("trasm", "TRASM (RASM) = выручка / ASM", "центы", lambda c: f"={H('rev',c)}/{H('asm',c)}*100", "0.00")
M("casm", "CASM = опер. расходы / ASM", "центы", lambda c: f"=({H('rev',c)}-{H('opinc',c)})/{H('asm',c)}*100", "0.00")
M("casmex", "CASM-ex (отчётный)", "центы", lambda c: f"={H('casmex_rep',c)}", "0.00", cols=YC[:9])
M("fcasm", "Топливо на ASM", "центы", lambda c: f"={H('fuel',c)}/{H('asm',c)}*100", "0.00")
M("spread", "Спред TRASM − CASM", "центы", lambda c: f"={R('trasm',c)}-{R('casm',c)}", "0.00")
M("gpa", "Топливная эффективность: галлонов на 1000 ASM", "галл.", lambda c: f"={H('gal',c)}/{H('asm',c)}*1000", "0.00")
M("asm_emp", "Производительность: ASM на сотрудника", "млн", lambda c: f"={H('asm',c)}/{H('emp',c)}/1000", "0.00", cols=FY)
M("asm_g", "Рост ASM г/г", "%", lambda c: f"={H('asm',c)}/{H('asm',str(int(c)-1))}-1", "0.0%", cols=FY[1:])
M("trasm_g", "Рост TRASM г/г", "%", lambda c: f"={R('trasm',c)}/{R('trasm',str(int(c)-1))}-1", "0.0%", cols=FY[1:])
M("casmex_g", "Рост CASM-ex г/г", "%", lambda c: f"={R('casmex',c)}/{R('casmex',str(int(c)-1))}-1", "0.0%", cols=FY[1:])
m.section("Структура выручки")
M("mix_pax", "Доля пассажирской выручки", "%", lambda c: f"={H('rev_pax',c)}/{H('rev',c)}", "0.0%")
M("mix_cargo", "Доля грузовой выручки", "%", lambda c: f"={H('rev_cargo',c)}/{H('rev',c)}", "0.0%")
M("mix_oth", "Доля прочей выручки", "%", lambda c: f"={H('rev_oth',c)}/{H('rev',c)}", "0.0%")
M("mix_loy", "Партнёрская выручка MileagePlus / выручка", "%", lambda c: f"={H('loy_oth',c)}/{H('rev',c)}", "0.0%", cols=["2023", "2024", "2025", "1П25", "1П26", "TTM 2К26"])
M("loy_total", "[О] MileagePlus всего ≈ партнёрская + погашения на перелёты", "$ млн", lambda c: f"={H('loy_oth',c)}+{H('miles_red',c)}", "#,##0", cols=["2024", "2025"], src="[О] нижняя оценка: не включает продажу миль, учтённую при начислении")
M("loy_total_pct", "[О] MileagePlus всего / выручка", "%", lambda c: f"={R('loy_total',c)}/{H('rev',c)}", "0.0%", cols=["2024", "2025"])
m.section("Прибыльность")
M("opm", "Операционная маржа (GAAP)", "%", lambda c: f"={H('opinc',c)}/{H('rev',c)}", "0.0%")
M("adj_opm", "Скорр. операционная маржа", "%", lambda c: f"={H('adj_opinc',c)}/{H('rev',c)}", "0.0%")
M("adj_ptm", "Скорр. маржа до налога", "%", lambda c: f"={H('adj_pti',c)}/{H('rev',c)}", "0.0%")
M("nim", "Чистая маржа", "%", lambda c: f"={H('ni',c)}/{H('rev',c)}", "0.0%")
M("ebitdar", "EBITDAR = скорр. EBITDA + аренда самолётов (постоянное определение)", "$ млн", lambda c: f"={H('adj_ebitda',c)}+{H('rent',c)}", "#,##0", src="формула; совпадает с отчётным EBITDAR UAL 2022–2024 (5,200 / 8,135 / 8,404)")
M("ebitdar_new", "Справочно: EBITDAR по новому определению UAL (+ фикс. часть опер. аренды)", "$ млн", lambda c: f"={H('adj_ebitda',c)}+{H('fixlease',c)}", "#,##0", cols=["2024", "2025", "1П25", "1П26", "TTM 2К26"], src="формула; TTM = 8,373 как в 8-K 15.07.2026")
M("ebitdar_m", "Маржа EBITDAR", "%", lambda c: f"={R('ebitdar',c)}/{H('rev',c)}", "0.0%")
m.section("Денежный поток и качество прибыли")
M("fcf", "FCF простой = CFO − капзатраты (GAAP)", "$ млн", lambda c: f"={H('cfo',c)}-{H('capex',c)}", "#,##0")
M("fcf_m", "Маржа FCF", "%", lambda c: f"={R('fcf',c)}/{H('rev',c)}", "0.0%")
M("cfo_ni", "CFO / чистая прибыль", "x", lambda c: f"=IF({H('ni',c)}>0,{H('cfo',c)}/{H('ni',c)},\"н/п\")", "0.0")
M("capex_rev", "Скорр. капзатраты / выручка", "%", lambda c: f"={H('adj_capex',c)}/{H('rev',c)}", "0.0%")
M("d_atl", "Изменение ATL + отложенной выручки MileagePlus (приток в CFO)", "$ млн", lambda c: f"=({H('atl',c)}+{H('ffdr',c)})-({H('atl',str(int(c)-1))}+{H('ffdr',str(int(c)-1))})", "#,##0", cols=["2021", "2022", "2023", "2024", "2025"])
M("cfo_exatl", "CFO без прироста ATL и отложенной выручки", "$ млн", lambda c: f"={H('cfo',c)}-{R('d_atl',c)}", "#,##0", cols=["2021", "2022", "2023", "2024", "2025"])
M("cfoex_ni", "CFO без ATL / чистая прибыль", "x", lambda c: f"=IF({H('ni',c)}>0,{R('cfo_exatl',c)}/{H('ni',c)},\"н/п\")", "0.0", cols=["2022", "2023", "2024", "2025"])
M("liab_rev", "(ATL + отложенная выручка) / выручка", "%", lambda c: f"=({H('atl',c)}+{H('ffdr',c)})/{H('rev',c)}", "0.0%", cols=["2020", "2021", "2022", "2023", "2024", "2025"])
m.section("Долговая нагрузка")
M("adj_debt", "Скорр. долг = долг + опер. аренда + пенсии", "$ млн", lambda c: f"={H('debt',c)}+{H('oplease',c)}+{H('pension',c)}", "#,##0")
M("cash_all", "Денежные средства + краткосрочные инвестиции", "$ млн", lambda c: f"={H('cash',c)}+{H('sti',c)}", "#,##0")
M("net_debt", "Чистый скорр. долг", "$ млн", lambda c: f"={R('adj_debt',c)}-{R('cash_all',c)}", "#,##0", src="формула; 1П26 = 18,105 как в 8-K 15.07.2026")
M("nd_ebitdar", "Чистый скорр. долг / EBITDAR", "x", lambda c: f"=IF({R('ebitdar',c)}>0,{R('net_debt',c)}/{R('ebitdar',c)},\"н/п\")", "0.0", cols=FY + ["TTM 2К26"])
M("nd_ebitdar_new", "Справочно: чистый долг / EBITDAR (новое опр. UAL)", "x", lambda c: f"={R('net_debt',c)}/{R('ebitdar_new',c)}", "0.0", cols=["2025", "TTM 2К26"])
M("nd_cfo", "Чистый скорр. долг / CFO", "x", lambda c: f"=IF({H('cfo',c)}>0,{R('net_debt',c)}/{H('cfo',c)},\"н/п\")", "0.0", cols=FY + ["TTM 2К26"])
M("nd_cfoex", "Чистый скорр. долг / CFO без прироста ATL", "x", lambda c: f"={R('net_debt',c)}/{R('cfo_exatl',c)}", "0.0", cols=["2022", "2023", "2024", "2025"])
M("cov", "Покрытие процентов = скорр. опер. прибыль / валовые %", "x", lambda c: f"={H('adj_opinc',c)}/{H('int_gross',c)}", "0.0", cols=FY + ["TTM 2К26"])
M("covr", "Покрытие фикс. платежей = EBITDAR / (валовые % + аренда самолётов)", "x", lambda c: f"={R('ebitdar',c)}/({H('int_gross',c)}+{H('rent',c)})", "0.0", cols=FY + ["TTM 2К26"])
M("roe", "Собственный капитал / активы", "%", lambda c: f"={H('equity',c)}/{H('assets',c)}", "0.0%", cols=FY + ["1П26"])
m.section("Нормализация через цикл")
rr = m.r
m.row("midcycle", "[О] Средняя скорр. опер. маржа: 2019, 2022–2025 (без пандемии 2020–21)", "%",
      {"2025": f"=AVERAGE({R('adj_opm','2019')},{R('adj_opm','2022')},{R('adj_opm','2023')},{R('adj_opm','2024')},{R('adj_opm','2025')})"}, fmt="0.0%", src="[О] mid-cycle маржа; 2022 — год восстановления, включён как слабая точка цикла", hl=True)
m.row("midcycle_ex22", "[О] То же без 2022", "%",
      {"2025": f"=AVERAGE({R('adj_opm','2019')},{R('adj_opm','2023')},{R('adj_opm','2024')},{R('adj_opm','2025')})"}, fmt="0.0%", src="[О] верхняя граница нормальной маржи")

# ============================================================== ASSUMPTIONS
YF = ["2026E", "2027E", "2028E", "2029E", "2030E", "2031E", "2032E"]
a = S("Assumptions", "Допущения модели (все вводные прогноза — здесь)", YF)
a.header("Параметр", "Ед.", "Обоснование / тип")
a.section("Рынок и оценка")
a.row("price", "Цена акции UAL", "$", {"2026E": 103.23}, fmt="0.00", src="[Ф] закрытие 13.10.2026, Yahoo Finance chart API (NASDAQ)")
a.row("shares", "Разводнённое число акций", "млн", {"2026E": 326.6}, fmt="0.0", src="[Ф] средневзвешенное разводнённое 2К26 (8-K 15.07.2026); базовое 324,6 млн на 09.07.2026")
a.row("rf", "Безрисковая ставка (UST 10 лет)", "%", {"2026E": 0.0524}, fmt="0.00%", src="[Ф] FRED DGS10, 09.10.2026 = 5,24%")
a.row("erp", "Премия за риск рынка акций", "%", {"2026E": 0.05}, fmt="0.00%", src="[Д] 5,0%: середина типичного диапазона 4,5–5,5% (Damodaran, implied ERP для США); не проверено на дату")
a.row("kd", "Стоимость долга до налога (вкл. аренду)", "%", {"2026E": 0.065}, fmt="0.00%", src="[Д] UST10 5,24% + ~1,25 п.п. спред BB+/Ba1; проверка: ноты 2031 (5,375%) размещены 02.2026 при UST10 4,29% → спред ~1,1 п.п.")
a.row("tax", "Налоговая ставка", "%", {"2026E": 0.22}, fmt="0.0%", src="[Д] эффективная ставка 2025 = 953/(3,353+953) = 22,1%; 1П26 = 20,7%")
a.row("nol", "PV налогового щита от убытков прошлых лет (NOL)", "$ млн", {"2026E": 2000}, src="[О] 10-K 2025: NOL $10,6 млрд (федеральные) → налоговый эффект $2,3 млрд вкл. штаты; использование за ~3–4 года, дисконт ~8,5% → ~$2,0 млрд")
a.row("h1fcff", "FCFF 1П26 фактический (FCF UAL + чистые % после налога)", "$ млн", {"2026E": f"={R('fcf_rep','1П26')}+{R('int_net','1П26')}*(1-{R('tax','2026E')})"}, src="[О] 3,225 + 273×(1−t); вычитается из FCFF 2026E, т.к. дисконтируем с 30.06.2026")
a.row("w_gordon", "Вес метода Гордона в итоговой DCF-оценке", "%", {"2026E": 0.5}, fmt="0%", src="[Д] 50/50 Гордон и выходной мультипликатор")
a.section("Общие операционные параметры (все сценарии)")
a.row("eff", "Изменение расхода топлива на ASM (обновление парка)", "%", {c: -0.01 for c in YF}, fmt="0.0%", src="[Д] 2019→2025: 15,06→14,12 галл./1000 ASM = −1,1% в год (Metrics)")
a.row("tp", "Расходы на сторонние услуги (не в CASM-ex)", "% выручки", {c: 0.0045 for c in YF}, fmt="0.00%", src="[Д] 1П26: 132/32,280 = 0,41%")
a.row("psr", "Участие в прибыли", "% от EBIT до него", {c: 0.12 for c in YF}, fmt="0.0%", src="[Д] калибровка: 2025 704/(4,972+704) = 12,4%; 2024 12,0%")
a.row("da_pct", "D&A", "% выручки", {c: 0.048 for c in YF}, fmt="0.0%", src="[Д] TTM 2К26: 2,996/62,902 = 4,8%")
a.row("fl_pct", "Фиксированная часть опер. аренды (для EBITDAR)", "% выручки", {c: 0.0155 for c in YF}, fmt="0.00%", src="[Д] TTM 2К26: 975/62,902 = 1,55%")
a.row("lease_r", "Ставка процента, заложенная в аренду", "%", {"2026E": 0.065}, fmt="0.0%", src="[Д] = стоимость долга; часть арендного платежа, переносимая в финансирование")
a.row("nwc", "Приток оборотного капитала (ATL, мили)", "% от прироста выручки", {c: 0.10 for c in YF}, fmt="0%", src="[Д] (ATL+отлож. выручка)/выручка = 27% (2025); с учётом оттоков по прочим статьям берём 10%")
a.row("netint", "Чистые процентные расходы (для EPS)", "$ млн", {c: 550 for c in YF}, src="[Д] TTM 2К26 = 541 (8-K 15.07.2026)")
a.row("capex26", "Скорр. капзатраты 2026E", "$ млн", {"2026E": 7500}, src="[Ф] прогноз UAL ~$7,5 млрд (Investor Update 15.07.2026)")
a.row("g", "Терминальный рост (Гордон)", "%", {"2026E": 0.025}, fmt="0.0%", src="[Д] ≈ номинальный рост ВВП США в долгосроке ниже текущей инфляции")
a.row("ronic", "Доходность новых инвестиций в терминале (RONIC)", "%", {"2026E": 0.10}, fmt="0.0%", src="[Д] чуть выше WACC: реинвестиции = g/RONIC = 25% NOPAT")
a.blank()
SCN = ["Bear", "Base", "Bull"]
SCN_RU = {"Bear": "Пессимистичный (рецессия + дорогое топливо)", "Base": "Базовый", "Bull": "Оптимистичный (премиальный спрос и маржа)"}
drivers = {
 "Bear": dict(asm=[.025, -.02, .02, .04, .04, .035, .03], core=[.025, -.05, .05, .045, .035, .027, .027], pt=[.8, .6, 1, 1, 1, 1, 1],
              casm=[.055, .035, .01, .015, .02, .027, .027], fuel=[3.97, 4.10, 3.40, 3.10, 3.10, 3.10, 3.10], capex=[None, .08, .08, .09, .10, .10, .10]),
 "Base": dict(asm=[.025, .04, .05, .05, .045, .04, .035], core=[.025, .02, .025, .025, .025, .025, .025], pt=[.8, .75, 1, 1, 1, 1, 1],
              casm=[.055, .025, .02, .023, .027, .027, .027], fuel=[3.97, 3.30, 3.00, 2.90, 2.90, 2.95, 3.00], capex=[None, .115, .12, .12, .11, .105, .10]),
 "Bull": dict(asm=[.025, .05, .06, .06, .05, .045, .04], core=[.025, .035, .03, .025, .025, .025, .025], pt=[.8, .6, 1, 1, 1, 1, 1],
              casm=[.055, .02, .02, .025, .027, .027, .027], fuel=[3.97, 2.90, 2.75, 2.70, 2.75, 2.80, 2.85], capex=[None, .115, .12, .12, .11, .105, .10]),
}
just = {
 "asm": "[Д] 2026: 1П +3,4%, 2П ~+1,6% (прогноз UAL 0…+2%) (прогноз UAL 21.04.2026, сокращение из-за топлива); далее — поставки 787/MAX/A321 (10-Q: ожидаемые поставки)",
 "core": "[Д] рост TRASM без переноса топлива; история 2019–2025: ~2,5% в год (TRASM 15,18→17,88, из них ~0,3 цента — топливо)",
 "pt": "[Д] доля изменения топливных затрат на ASM, переносимая в тариф; UAL: 1П ~50%, 3 кв. 80–90%, 4 кв. 100% (Investor Update 15.07.2026)",
 "casm": "[Д] 2026: 1П +6,0% (сокращение мощностей, новые договоры AFA/IAM); 2027+ — инфляция зарплат против эффекта масштаба",
 "fuel": "[Д] 2026: 1П $3,53 факт, 2П ~$4,40 (спот Gulf Coast сент.–окт. $4,34–4,39 + ~$0,35 логистика/налоги); 2027 — EIA STEO окт. 2026 (Brent $84)",
 "capex": "[Д] 2026 — прогноз UAL $7,5 млрд; далее % выручки по контрактным обязательствам 10-K (2027–2030: $5,6–9,0 млрд в год только по заказам)",
}
labels = {"asm": ("Рост ASM", "%", "0.0%"), "core": ("Базовый рост TRASM (без топлива)", "%", "0.0%"), "pt": ("Перенос топлива в тариф", "доля", "0%"),
          "casm": ("Рост CASM-ex", "%", "0.0%"), "fuel": ("Цена топлива (all-in)", "$/галл.", "0.00"), "capex": ("Скорр. капзатраты", "% выручки", "0.0%")}
for s in SCN:
    a.section(f"Сценарий: {SCN_RU[s]}")
    for k in ["asm", "core", "pt", "casm", "fuel", "capex"]:
        lab, u, fm = labels[k]
        vals_ = {c: v for c, v in zip(YF, drivers[s][k]) if v is not None}
        a.row(f"{s}_in_{k}", lab, u, vals_, fmt=fm, src=just[k] if s == "Base" else "[Д] см. логику базового сценария")
a.section("Вероятности сценариев и выходные мультипликаторы")
a.row("prob_Bear", "Вероятность: пессимистичный", "%", {"2026E": 0.25}, fmt="0%", src="[Д] Brent >$100, UST10 >5%, ослабление спроса — реальный риск 2027")
a.row("prob_Base", "Вероятность: базовый", "%", {"2026E": 0.50}, fmt="0%", src="[Д]")
a.row("prob_Bull", "Вероятность: оптимистичный", "%", {"2026E": 0.25}, fmt="0%", src="[Д]")
a.row("prob_chk", "Проверка: сумма вероятностей", "%", {"2026E": f"={R('prob_Bear','2026E')}+{R('prob_Base','2026E')}+{R('prob_Bull','2026E')}"}, fmt="0%", src="формула (должно быть 100%)")
a.row("rule_buy", "Порог рекомендации Buy: потенциал выше", "%", {"2026E": 0.15}, fmt="0%", src="[Д] правило рейтинга")
a.row("rule_sell", "Порог рекомендации Sell: потенциал ниже", "%", {"2026E": -0.10}, fmt="0%", src="[Д] правило рейтинга")

# ============================================================== BETA
import csv
def monthly(t):
    d = json.load(open(f"{SCR}/data/chart_{t}.json"))["chart"]["result"][0]
    c = d["indicators"].get("adjclose", [{}])[0].get("adjclose") or d["indicators"]["quote"][0]["close"]
    return {dt.datetime.fromtimestamp(ts, dt.UTC).strftime("%Y-%m"): v for ts, v in zip(d["timestamp"], c) if v}
u = monthly("UAL"); sp = monthly("^GSPC")
ks = sorted(set(u) & set(sp))[-61:]
b = wb.create_sheet("Beta"); b["A1"] = "Бета UAL: месячные доходности против S&P 500, 60 месяцев (10.2021–10.2026)"; b["A1"].font = TITLE
b.column_dimensions["A"].width = 12
for j, hname in enumerate(["Месяц", "UAL (скорр. цена)", "S&P 500", "Доходн. UAL", "Доходн. S&P"]):
    x = b.cell(3, 1 + j, hname); x.font = HDR; x.fill = HFILL
    b.column_dimensions[L(1 + j)].width = 16
for i, k in enumerate(ks):
    r = 4 + i
    b.cell(r, 1, k).font = BLUE; b.cell(r, 2, round(u[k], 4)).font = BLUE; b.cell(r, 3, round(sp[k], 2)).font = BLUE
    if i > 0:
        b.cell(r, 4, f"=B{r}/B{r-1}-1").number_format = "0.0%"; b.cell(r, 5, f"=C{r}/C{r-1}-1").number_format = "0.0%"
last = 4 + len(ks) - 1
b.cell(3, 7, "Бета (SLOPE)").font = BOLD; b.cell(3, 8, f"=SLOPE(D5:D{last},E5:E{last})").number_format = "0.00"
b.cell(4, 7, "Скорр. бета (Блюм)").font = BOLD; b.cell(4, 8, "=0.67*H3+0.33").number_format = "0.00"
b.cell(6, 7, "Источник: Yahoo Finance chart API, месячные цены закрытия (adjclose); последняя точка — 13.10.2026 (неполный месяц).").font = NOTE
b.cell(7, 7, "Для справки (тот же расчёт): DAL 1,42; AAL 1,44; LUV 1,23; ALK 1,42; JBLU 1,84 — медиана пиров 1,42.").font = NOTE
b.column_dimensions["G"].width = 22
REG["beta"] = {"v": "'Beta'!$H$3"}; REG["beta_adj"] = {"v": "'Beta'!$H$4"}

# ============================================================== WACC
w = S("WACC", "Ставка дисконтирования (WACC) с учётом аренды", ["Значение"])
w.header("Компонент", "Ед.")
c0 = "Значение"
w.row("w_rf", "Безрисковая ставка", "%", {c0: f"={R('rf','2026E')}"}, fmt="0.00%")
w.row("w_beta", "Бета (5 лет, месячная, сырая)", "x", {c0: f"={REG['beta']['v']}"}, fmt="0.00", src="лист Beta; скорр. по Блюму ~1,24 — ниже, даёт WACC ~8,2%")
w.row("w_erp", "Премия за риск", "%", {c0: f"={R('erp','2026E')}"}, fmt="0.00%")
w.row("ke", "Стоимость собственного капитала = Rf + β × ERP", "%", {c0: f"={R('w_rf',c0)}+{R('w_beta',c0)}*{R('w_erp',c0)}"}, fmt="0.00%", bold=True)
w.row("w_kd", "Стоимость долга до налога", "%", {c0: f"={R('kd','2026E')}"}, fmt="0.00%")
w.row("w_kd_at", "Стоимость долга после налога", "%", {c0: f"={R('w_kd',c0)}*(1-{R('tax','2026E')})"}, fmt="0.00%")
w.row("w_E", "Рыночная капитализация (цена × разводн. акции)", "$ млн", {c0: f"={R('price','2026E')}*{R('shares','2026E')}"}, fmt="#,##0")
w.row("w_D", "Скорр. долг 30.06.2026 (долг + опер. аренда + пенсии)", "$ млн", {c0: f"={R('adj_debt','1П26')}"}, fmt="#,##0", src="Metrics; аренда включена в долг")
w.row("w_we", "Доля собственного капитала", "%", {c0: f"={R('w_E',c0)}/({R('w_E',c0)}+{R('w_D',c0)})"}, fmt="0.0%")
w.row("w_wd", "Доля долга", "%", {c0: f"=1-{R('w_we',c0)}"}, fmt="0.0%")
w.row("wacc", "WACC", "%", {c0: f"={R('w_we',c0)}*{R('ke',c0)}+{R('w_wd',c0)}*{R('w_kd_at',c0)}"}, fmt="0.00%", bold=True, hl=True)
w.row("w_netdebt", "Чистый скорр. долг 30.06.2026 (вычитается из EV)", "$ млн", {c0: f"={R('net_debt','1П26')}"}, fmt="#,##0")

# ============================================================== DCF
DC = ["2025A"] + YF
d = S("DCF", "DCF (FCFF) — три сценария, 2026E–2032E; дата оценки 30.06.2026 (баланс), середина периода", DC)
TPER = {"2026E": 0.25, "2027E": 1.0, "2028E": 2.0, "2029E": 3.0, "2030E": 4.0, "2031E": 5.0, "2032E": 6.0}
def dcf_block(sh, s, pref, delta_key=None):
    """s = scenario key for drivers; pref = registry prefix; delta_key = optional ref added to core RASM 2027+"""
    A = lambda k, c: R(f"{s}_in_{k}", c)
    P = lambda c: DC[DC.index(c) - 1]
    X = lambda k, c: R(f"{pref}{k}", c)
    sh.header(f"{SCN_RU.get(s, s)}" + (" — reverse DCF" if delta_key else ""), "Ед.")
    sh.row(f"{pref}asm", "ASM", "млн", lambda c: f"={R('asm','2025')}" if c == "2025A" else f"={X('asm',P(c))}*(1+{A('asm',c)})")
    sh.row(f"{pref}gpa", "Галлонов на 1000 ASM", "галл.", lambda c: f"={R('gpa','2025')}" if c == "2025A" else f"={X('gpa',P(c))}*(1+{R('eff',c)})", fmt="0.00")
    sh.row(f"{pref}fpx", "Цена топлива", "$/галл.", lambda c: f"={R('fuelpx','2025')}" if c == "2025A" else f"={A('fuel',c)}", fmt="0.00")
    sh.row(f"{pref}fcasm", "Топливо на ASM", "центы", lambda c: f"={X('gpa',c)}*{X('fpx',c)}/10", fmt="0.00")
    dk = (lambda c: f"+{delta_key}" if c not in ("2026E",) else "") if delta_key else (lambda c: "")
    sh.row(f"{pref}trasm", "TRASM = TRASM(t−1)×(1+базовый рост) + перенос×Δтопливо/ASM", "центы",
           lambda c: f"={R('trasm','2025')}" if c == "2025A" else f"={X('trasm',P(c))}*(1+{A('core',c)}{dk(c)})+{A('pt',c)}*({X('fcasm',c)}-{X('fcasm',P(c))})", fmt="0.00")
    sh.row(f"{pref}rev", "Выручка", "$ млн", lambda c: f"={R('rev','2025')}" if c == "2025A" else f"={X('asm',c)}*{X('trasm',c)}/100", bold=True)
    sh.row(f"{pref}casmex", "CASM-ex", "центы", lambda c: f"={R('casmex','2025')}" if c == "2025A" else f"={X('casmex',P(c))}*(1+{A('casm',c)})", fmt="0.00")
    F = YF
    sh.row(f"{pref}nfc", "Нетопливные расходы (CASM-ex × ASM)", "$ млн", {c: f"={X('asm',c)}*{X('casmex',c)}/100" for c in F})
    sh.row(f"{pref}fc", "Топливо", "$ млн", {c: f"={X('asm',c)}*{X('fcasm',c)}/100" for c in F})
    sh.row(f"{pref}tpc", "Сторонние услуги", "$ млн", {c: f"={X('rev',c)}*{R('tp',c)}" for c in F})
    sh.row(f"{pref}pre", "EBIT до участия в прибыли", "$ млн", {c: f"={X('rev',c)}-{X('nfc',c)}-{X('fc',c)}-{X('tpc',c)}" for c in F})
    sh.row(f"{pref}psx", "Участие в прибыли", "$ млн", {c: f"=MAX(0,{X('pre',c)})*{R('psr',c)}" for c in F})
    sh.row(f"{pref}ebit", "Скорр. EBIT", "$ млн", lambda c: f"={R('adj_opinc','2025')}" if c == "2025A" else f"={X('pre',c)}-{X('psx',c)}", bold=True)
    sh.row(f"{pref}m", "Скорр. операционная маржа", "%", lambda c: f"={X('ebit',c)}/{X('rev',c)}", fmt="0.0%", hl=True)
    sh.row(f"{pref}da", "D&A", "$ млн", lambda c: f"={R('da','2025')}" if c == "2025A" else f"={X('rev',c)}*{R('da_pct',c)}")
    sh.row(f"{pref}ebitdar", "EBITDAR = EBIT + D&A + фикс. аренда", "$ млн", lambda c: f"={R('ebitdar_new','2025')}" if c == "2025A" else f"={X('ebit',c)}+{X('da',c)}+{X('rev',c)}*{R('fl_pct',c)}")
    sh.row(f"{pref}ll", "Обязательства по опер. аренде", "$ млн", lambda c: f"={R('oplease','1П26')}" if c == "2026E" else (None if c == "2025A" else f"={X('ll',P(c))}*(1+{A('asm',c)})"))
    sh.row(f"{pref}li", "Процентная часть аренды (возвращается в EBIT)", "$ млн", {c: f"={X('ll',c)}*{R('lease_r','2026E')}" for c in F})
    sh.row(f"{pref}nopat", "NOPAT (с поправкой на аренду) = (EBIT + % аренды) × (1 − t)", "$ млн", {c: f"=({X('ebit',c)}+{X('li',c)})*(1-{R('tax','2026E')})" for c in F})
    sh.row(f"{pref}capex", "Скорр. капзатраты", "$ млн", {c: (f"={R('capex26','2026E')}" if c == "2026E" else f"={X('rev',c)}*{A('capex',c)}") for c in F})
    sh.row(f"{pref}nwc", "Приток оборотного капитала", "$ млн", {c: f"=({X('rev',c)}-{X('rev',P(c))})*{R('nwc',c)}" for c in F})
    sh.row(f"{pref}fcff", "FCFF = NOPAT + D&A − капзатраты + Δоборотный капитал", "$ млн", {c: f"={X('nopat',c)}+{X('da',c)}-{X('capex',c)}+{X('nwc',c)}" for c in F}, bold=True)
    sh.row(f"{pref}fcffd", "FCFF к дисконтированию (2026E — только 2П: минус факт 1П26)", "$ млн", {c: (f"={X('fcff',c)}-{R('h1fcff','2026E')}" if c == "2026E" else f"={X('fcff',c)}") for c in F})
    sh.row(f"{pref}t", "Период дисконтирования (лет от 30.06.2026, середина периода)", "лет", {c: TPER[c] for c in F}, fmt="0.00", src="[Д] середина периода; 2П26 = 0,25")
    sh.row(f"{pref}df", "Коэффициент дисконтирования", "x", {c: f"=1/(1+{R('wacc','Значение')})^{X('t',c)}" for c in F}, fmt="0.000")
    sh.row(f"{pref}pv", "PV FCFF", "$ млн", {c: f"={X('fcffd',c)}*{X('df',c)}" for c in F})
    sh.row(f"{pref}eps", "EPS скорр. = (EBIT − чистые %) × (1 − t) / акции", "$", {c: f"=({X('ebit',c)}-{R('netint',c)})*(1-{R('tax','2026E')})/{R('shares','2026E')}" for c in F}, fmt="0.00")
    c = "2032E"
    wacc = R("wacc", "Значение"); g = R("g", "2026E")
    sh.row(f"{pref}sumpv", "Сумма PV FCFF 2П26–2032E", "$ млн", {c: f"=SUM({sh.col('2026E')}{sh.r-2}:{sh.col('2032E')}{sh.r-2})"})
    sh.row(f"{pref}tvg", "TV Гордон = NOPAT×(1+g)×(1−g/RONIC)/(WACC−g)", "$ млн", {c: f"={X('nopat',c)}*(1+{g})*(1-{g}/{R('ronic','2026E')})/({wacc}-{g})"})
    sh.row(f"{pref}mult", "Выходной мультипликатор EV/EBITDAR", "x", {c: "=__EXIT__" + s}, fmt="0.0", src="история UAL: мин / медиана / макс (лист Multiples)")
    sh.row(f"{pref}tvm", "TV по мультипликатору = EV/EBITDAR × EBITDAR 2032E", "$ млн", {c: f"={X('mult',c)}*{X('ebitdar',c)}"})
    sh.row(f"{pref}dftv", "Дисконт TV (6,5 года)", "x", {c: f"=1/(1+{wacc})^6.5"}, fmt="0.000")
    sh.row(f"{pref}evg", "EV (Гордон)", "$ млн", {c: f"={X('sumpv',c)}+{X('tvg',c)}*{X('dftv',c)}"})
    sh.row(f"{pref}evm", "EV (мультипликатор)", "$ млн", {c: f"={X('sumpv',c)}+{X('tvm',c)}*{X('dftv',c)}"})
    sh.row(f"{pref}tvshare", "Доля TV в EV (Гордон)", "%", {c: f"={X('tvg',c)}*{X('dftv',c)}/{X('evg',c)}"}, fmt="0%")
    nd = R("w_netdebt", "Значение"); nol = R("nol", "2026E"); shs = R("shares", "2026E")
    sh.row(f"{pref}psg", "Стоимость акции (Гордон) = (EV − чистый скорр. долг + NOL) / акции", "$", {c: f"=({X('evg',c)}-{nd}+{nol})/{shs}"}, fmt="0.00")
    sh.row(f"{pref}psm", "Стоимость акции (мультипликатор)", "$", {c: f"=({X('evm',c)}-{nd}+{nol})/{shs}"}, fmt="0.00")
    sh.row(f"{pref}ps", "Стоимость акции DCF (взвешенная)", "$", {c: f"={R('w_gordon','2026E')}*{X('psg',c)}+(1-{R('w_gordon','2026E')})*{X('psm',c)}"}, fmt="0.00", bold=True, hl=True)
    sh.row(f"{pref}impl_mult", "Проверка: EV/EBITDAR, заложенный в TV Гордона", "x", {c: f"={X('tvg',c)}/{X('ebitdar',c)}"}, fmt="0.0")
    sh.blank()

for s in SCN:
    dcf_block(d, s, f"{s}_")
d.note("Аренда: операционная аренда включена в долг (EV, WACC); процентная часть арендного платежа возвращена в NOPAT, чтобы избежать двойного учёта.")
d.note("FCFF облагается полной ставкой 22%; фактический налоговый щит NOL ($2,3 млрд) добавлен отдельно как актив (Assumptions). 2026E одинаков во всех сценариях.")

# ============================================================== REVERSE DCF
rv = S("ReverseDCF", "Reverse DCF: какой рост RASM и какой WACC заложены в цену $103,23", DC)
rv.header("Поправка к базовому росту TRASM в 2027–2032 (п.п. в год)", "", "")
DELTAS = [-0.01, -0.005, 0.0, 0.005, 0.01]
rv.row("rv_deltas", "Поправки", "п.п.", {c: v for c, v in zip(YF[:5], DELTAS)}, fmt="+0.0%;-0.0%;0.0%", src="[Д] сетка")
for i, dl in enumerate(DELTAS):
    dref = R("rv_deltas", YF[i])
    dcf_block(rv, "Base", f"RV{i}_", delta_key=dref)
rv.header("Итог reverse DCF", "", "")
pr = R("price", "2026E")
rv.row("rv_val", "Стоимость акции при поправке", "$", {YF[i]: f"={R(f'RV{i}_ps','2032E')}" for i in range(5)}, fmt="0.00")
rv.row("rv_m", "Маржа 2032E при поправке", "%", {YF[i]: f"={R(f'RV{i}_m','2032E')}" for i in range(5)}, fmt="0.0%")
v = [R("rv_val", YF[i]) for i in range(5)]; dl = [R("rv_deltas", YF[i]) for i in range(5)]; mm = [R("rv_m", YF[i]) for i in range(5)]
def interp(ys):
    parts = []
    for i in range(4):
        parts.append(f"IF(AND({pr}>={v[i]},{pr}<={v[i+1]}),{ys[i]}+({pr}-{v[i]})/({v[i+1]}-{v[i]})*({ys[i+1]}-{ys[i]}),")
    return "=" + "".join(parts) + '"вне сетки"' + ")" * 4
rv.row("rv_impl_d", "Поправка к росту TRASM, заложенная в цену", "п.п./год", {"2026E": interp(dl)}, fmt="+0.00%;-0.00%", bold=True, hl=True, src="линейная интерполяция по сетке")
rv.row("rv_impl_m", "Операционная маржа 2032E, заложенная в цену", "%", {"2026E": interp(mm)}, fmt="0.0%", bold=True, hl=True, src="сравнить с mid-cycle маржой (Metrics)")
rv.row("rv_mid", "Справочно: mid-cycle маржа UAL", "%", {"2026E": f"={R('midcycle','2025')}"}, fmt="0.0%")

# ============================================================== MULTIPLES
peers = json.load(open(f"{SCR}/data/peers_ttm_q2_2026.json"))
PC = ["UAL", "DAL", "AAL", "LUV", "ALK", "JBLU"]
mu = S("Multiples", "Сравнительная оценка: пиры (TTM на 30.06.2026, цены на 13.10.2026)", PC + ["Медиана пиров"])
mu.header("Показатель", "Ед.")
mu.section("Вводные (XBRL companyfacts SEC, 10-K 2025 + 10-Q 2К26)")
px = {"UAL": 103.23, "DAL": 79.57, "AAL": 12.50, "LUV": 40.65, "ALK": 37.81, "JBLU": 3.35}
shs = {"UAL": 326.6, "DAL": 657.912, "AAL": 661.970, "LUV": 489.208, "ALK": 111.603, "JBLU": 377.312}
peers["UAL"]["pension"] = 1074; peers["DAL"]["sti"] = 0; peers["LUV"]["sti"] = 0; peers["LUV"]["rent"] = 0; peers["JBLU"]["pension"] = 0
peers["ALK"]["capex"] = 1154
peers["UAL"]["cash"] = 10166
def prow(key, label, unit, field, src, fmt="#,##0"):
    mu.row(key, label, unit, {p: round(peers[p][field], 1) if peers[p][field] is not None else None for p in PC}, fmt=fmt, src=src)
mu.row("p_px", "Цена акции 13.10.2026", "$", px, fmt="0.00", src="[Ф] Yahoo Finance chart API, закрытие 13.10.2026")
mu.row("p_sh", "Акций", "млн", shs, fmt="0.0", src="[Ф] обложки 10-Q (DAL — 10-Q 3К26); UAL — разводнённые 2К26 для согласования с DCF")
prow("p_rev", "Выручка TTM", "$ млн", "rev", "[Ф] XBRL Revenues/RevenueFromContract…: FY2025 + 1П26 − 1П25")
prow("p_op", "Операционная прибыль TTM (GAAP)", "$ млн", "opinc", "[Ф] XBRL OperatingIncomeLoss (включая спецстатьи: у ALK/AAL/JBLU — расходы на интеграцию/реструктуризацию; у UAL — доход $615 млн от sale-leaseback)")
prow("p_da", "D&A TTM", "$ млн", "da", "[Ф] XBRL")
prow("p_rent", "Аренда самолётов TTM", "$ млн", "rent", "[Ф] XBRL AircraftRental; LUV — отдельно не раскрывается (0, флот в собственности)")
prow("p_ni", "Чистая прибыль TTM", "$ млн", "ni", "[Ф] XBRL NetIncomeLoss")
prow("p_cfo", "CFO TTM", "$ млн", "cfo", "[Ф] XBRL")
prow("p_capex", "Капзатраты TTM", "$ млн", "capex", "[Ф] XBRL; DAL = ProductiveAssets FY25 + (FlightEquipment+Other) 1П26 − 1П25; ALK = FlightEquipment + OtherPP&E")
mu.row("p_debt", "Долг + фин. аренда (тек. + долгоср.)", "$ млн", {p: round(peers[p]["debt_c"] + peers[p]["debt_lt"], 1) for p in PC}, src="[Ф] XBRL LongTermDebtAndCapitalLeaseObligations (+Current), 30.06.2026")
mu.row("p_ol", "Опер. аренда (обязательства)", "$ млн", {p: round(peers[p]["ol_c"] + peers[p]["ol_nc"], 1) for p in PC}, src="[Ф] XBRL OperatingLeaseLiability (Current+Noncurrent)")
prow("p_pen", "Пенсии и пострабочие (долгоср.)", "$ млн", "pension", "[Ф] XBRL; UAL — 8-K 15.07.2026; JBLU — не найдено (0)")
mu.row("p_cash", "Денежные средства + краткоср. инвестиции", "$ млн", {p: round(peers[p]["cash"] + (peers[p]["sti"] or 0), 1) for p in PC}, src="[Ф] XBRL; DAL и LUV — краткоср. инвестиции не найдены в XBRL (0)")
prow("p_eq", "Собственный капитал", "$ млн", "equity", "[Ф] XBRL StockholdersEquity")
mu.row("p_fpe_in", "Прогноз EPS FY2026 (консенсус/гайденс)", "$", {"UAL": 7.92, "DAL": 5.35, "AAL": 0.0}, fmt="0.00", src="[Ф] UAL — консенсус Yahoo Finance (~13.10.2026, 23 аналитика); DAL — середина гайденса $5,10–5,60 (09.10.2026); AAL — середина (−0,65…0,65) (23.07.2026); LUV/ALK/JBLU — не найдено")
mu.section("Расчёт (формулы)")
X = lambda k, p: R(k, p)
mu.row("p_mcap", "Капитализация", "$ млн", {p: f"={X('p_px',p)}*{X('p_sh',p)}" for p in PC})
mu.row("p_ev", "EV = капитализация + долг + опер. аренда + пенсии − ден. средства", "$ млн", {p: f"={X('p_mcap',p)}+{X('p_debt',p)}+{X('p_ol',p)}+{X('p_pen',p)}-{X('p_cash',p)}" for p in PC})
mu.row("p_ebitdar", "EBITDAR (GAAP) = опер. прибыль + D&A + аренда самолётов", "$ млн", {p: f"={X('p_op',p)}+{X('p_da',p)}+{X('p_rent',p)}" for p in PC})
def med(k, fmt):
    return f"=MEDIAN({X(k,'DAL')},{X(k,'AAL')},{X(k,'LUV')},{X(k,'ALK')},{X(k,'JBLU')})"
def mrow(key, label, unit, fn, fmt, src="формула", medok=True):
    vals_ = {p: fn(p) for p in PC}
    if medok: vals_["Медиана пиров"] = f"=MEDIAN({mu.col('DAL')}{mu.r}:{mu.col('JBLU')}{mu.r})"
    mu.row(key, label, unit, vals_, fmt=fmt, src=src)
mrow("m_evebitdar", "EV / EBITDAR", "x", lambda p: f"=IF({X('p_ebitdar',p)}>0,{X('p_ev',p)}/{X('p_ebitdar',p)},\"н/п\")", "0.0")
mrow("m_evrev", "EV / выручка", "x", lambda p: f"={X('p_ev',p)}/{X('p_rev',p)}", "0.00")
mrow("m_pe", "P/E TTM (GAAP)", "x", lambda p: f"=IF({X('p_ni',p)}>0,{X('p_mcap',p)}/{X('p_ni',p)},\"н/п\")", "0.0")
mu.row("m_fpe", "P/E FY2026 (консенсус/гайденс)", "x", {p: f"=IF({X('p_fpe_in',p)}>0,{X('p_px',p)}/{X('p_fpe_in',p)},\"н/п\")" for p in ["UAL", "DAL", "AAL"]}, fmt="0.0", src="AAL — н/п (прибыль около нуля)")
mrow("m_fcfy", "FCF yield = (CFO − капзатраты) / капитализация", "%", lambda p: f"=({X('p_cfo',p)}-{X('p_capex',p)})/{X('p_mcap',p)}", "0.0%")
mrow("m_pfcf", "P/FCF", "x", lambda p: f"=IF({X('p_cfo',p)}-{X('p_capex',p)}>0,{X('p_mcap',p)}/({X('p_cfo',p)}-{X('p_capex',p)}),\"н/п\")", "0.0")
mrow("m_pb", "P/B", "x", lambda p: f"=IF({X('p_eq',p)}>0,{X('p_mcap',p)}/{X('p_eq',p)},\"н/п\")", "0.00")
mrow("m_opm", "Операционная маржа TTM (GAAP)", "%", lambda p: f"={X('p_op',p)}/{X('p_rev',p)}", "0.0%")
mrow("m_lev", "Чистый скорр. долг / EBITDAR", "x", lambda p: f"=IF({X('p_ebitdar',p)}>0,({X('p_ev',p)}-{X('p_mcap',p)})/{X('p_ebitdar',p)},\"н/п\")", "0.0")
mu.note("Медиана пиров рассчитана по DAL, AAL, LUV, ALK, JBLU (без UAL); MEDIAN игнорирует «н/п».")
mu.note("Некорректность сравнения: LUV — point-to-point лоукостер (короткие рейсы, без премиум-кабин до 2026); JBLU и ALK — убыточны (интеграция Hawaiian у ALK), их EV/EBITDAR завышен;")
mu.note("AAL — отрицательный капитал (P/B н/п) и прибыль около нуля; DAL — собственный НПЗ (снижает цену топлива ~$0,3–0,4/галл.) и выручка НПЗ в TRASM; у UAL TTM-прибыль включает $615 млн разовых доходов.")
mu.blank()
# UAL own history
HC = ["2019", "2022", "2023", "2024", "2025", "Текущий"]
hh = S.__new__(S); hh.ws = mu.ws; hh.name = "Multiples"; hh.cols = HC; hh.fc = 3; hh.r = mu.r; hh.srccol = mu.srccol
hh.header("История мультипликаторов UAL (на конец года; 2020–21 не показательны — убытки)", "Ед.")
def hcol(c): return "TTM 2К26" if c == "Текущий" else c
hh.row("hm_px", "Цена акции", "$", {c: (f"={R('price','2026E')}" if c == "Текущий" else f"={R('px_ye',c)}") for c in HC}, fmt="0.00")
hh.row("hm_sh", "Акций", "млн", {c: (f"={R('shares','2026E')}" if c == "Текущий" else f"={R('sh_out',c)}") for c in HC}, fmt="0.0")
hh.row("hm_mcap", "Капитализация", "$ млн", {c: f"={R('hm_px',c)}*{R('hm_sh',c)}" for c in HC})
hh.row("hm_nd", "Чистый скорр. долг", "$ млн", {c: f"={R('net_debt',hcol(c))}" for c in HC}, src="2019 — [О] на основе XBRL (определение пенсий уже)")
hh.row("hm_ev", "EV", "$ млн", {c: f"={R('hm_mcap',c)}+{R('hm_nd',c)}" for c in HC})
hh.row("hm_ebitdar", "EBITDAR (скорр. EBITDA + аренда самолётов)", "$ млн", {c: f"={R('ebitdar',hcol(c))}" for c in HC})
hh.row("hm_evebitdar", "EV / EBITDAR", "x", {c: f"={R('hm_ev',c)}/{R('hm_ebitdar',c)}" for c in HC}, fmt="0.0", hl=True)
hh.row("hm_eps", "Скорр. EPS (текущий — TTM)", "$", {c: f"={R('adj_eps',hcol(c))}" for c in HC}, fmt="0.00")
hh.row("hm_pe", "P/E (скорр.)", "x", {c: f"={R('hm_px',c)}/{R('hm_eps',c)}" for c in HC}, fmt="0.0")
hh.row("hm_pb", "P/B", "x", {c: f"={R('hm_mcap',c)}/{R('equity',hcol(c))}" for c in HC}, fmt="0.00")
hh.row("hm_fcfy", "FCF yield (CFO − капзатраты GAAP)", "%", {c: f"={R('fcf',hcol(c))}/{R('hm_mcap',c)}" for c in HC}, fmt="0.0%")
hh.row("hm_evrev", "EV / выручка", "x", {c: f"={R('hm_ev',c)}/{R('rev',hcol(c))}" for c in HC}, fmt="0.00")
mu.r = hh.r
SC = ["EV/EBITDAR", "P/E", "P/B"]
st = S.__new__(S); st.ws = mu.ws; st.name = "Multiples"; st.cols = SC; st.fc = 3; st.r = mu.r; st.srccol = mu.srccol
st.header("Статистика истории UAL (2019, 2022–2025)", "")
hist_cols = ["2019", "2022", "2023", "2024", "2025"]
for key, lab, fn in [("hist_min", "Минимум", "MIN"), ("hist_med", "Медиана", "MEDIAN"), ("hist_max", "Максимум", "MAX")]:
    st.row(key, lab, "x", {"EV/EBITDAR": f"={fn}({','.join(R('hm_evebitdar',c) for c in hist_cols)})",
                           "P/E": f"={fn}({','.join(R('hm_pe',c) for c in hist_cols)})",
                           "P/B": f"={fn}({','.join(R('hm_pb',c) for c in hist_cols)})"}, fmt="0.0")
mu.r = st.r
mu.blank()
IC = ["Значение"]
im = S.__new__(S); im.ws = mu.ws; im.name = "Multiples"; im.cols = IC; im.fc = 3; im.r = mu.r; im.srccol = mu.srccol
im.header("Оценка UAL по мультипликаторам", "")
nd = R("w_netdebt", "Значение"); shs_ = R("shares", "2026E")
im.row("im_peer", "EV/EBITDAR медиана пиров × EBITDAR TTM UAL (GAAP)", "$/акция", {"Значение": f"=({R('m_evebitdar','Медиана пиров')}*{R('p_ebitdar','UAL')}-{nd})/{shs_}"}, fmt="0.00", src="пиры убыточны/искажены — справочно")
im.row("im_hist27", "EV/EBITDAR медиана истории UAL × EBITDAR 2027E (база)", "$/акция", {"Значение": f"=({R('hist_med','EV/EBITDAR')}*{R('Base_ebitdar','2027E')}-{nd}+{R('nol','2026E')})/{shs_}"}, fmt="0.00", src="EBITDAR 2027E — лист DCF, базовый сценарий")
im.row("im_normeps", "Нормализованная EPS = mid-cycle маржа × выручка 2027E (база)", "$", {"Значение": f"=({R('midcycle','2025')}*{R('Base_rev','2027E')}-{R('netint','2027E')})*(1-{R('tax','2026E')})/{shs_}"}, fmt="0.00", src="[О]")
im.row("im_pe_norm", "P/E на нормализованной EPS (текущая цена)", "x", {"Значение": f"={R('price','2026E')}/{R('im_normeps','Значение')}"}, fmt="0.0")
im.row("im_normpe", "Нормализованная EPS × медианный P/E истории", "$/акция", {"Значение": f"={R('im_normeps','Значение')}*{R('hist_med','P/E')}"}, fmt="0.00")
mu.r = im.r
mu.blank()
mu.section("Международные пиры (только контекст; мультипликаторы не рассчитаны)")
mu.note("Air France-KLM (AF.PA) €10,48; IAG (IAG.L) 412,45 пенса; Lufthansa (LHA.DE) €7,10 — котировки Yahoo Finance 14.10.2026 (внутри дня).")
mu.note("Причины не сравнивать напрямую: МСФО (IFRS 16 — вся аренда уже в EBITDA, EBITDAR ≈ EBITDA), другая валюта, госучастие (AF-KLM), структура хабов ЕС и регулирование слотов.")
mu.note("Первичные отчёты за 1П26 в рамках работы не собраны — не найдено; на вывод не влияет, т.к. рекомендация опирается на DCF и собственную историю UAL.")

# patch exit multiple placeholders
for ws in [d.ws, rv.ws]:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("=__EXIT__"):
                s = cell.value.replace("=__EXIT__", "")
                cell.value = "=" + {"Bear": R("hist_min", "EV/EBITDAR"), "Base": R("hist_med", "EV/EBITDAR"), "Bull": R("hist_max", "EV/EBITDAR")}[s]

# ============================================================== SCENARIOS
sc = S("Scenarios", "Сценарии и целевая цена (горизонт 12 месяцев)", SCN + ["Взвешенно"])
sc.header("Показатель", "Ед.")
sc.row("sc_prob", "Вероятность", "%", {s: f"={R('prob_'+s,'2026E')}" for s in SCN} | {"Взвешенно": f"=SUM({sc.col('Bear')}{sc.r}:{sc.col('Bull')}{sc.r})"}, fmt="0%")
sc.row("sc_val", "Стоимость акции DCF", "$", {s: f"={R(s+'_ps','2032E')}" for s in SCN} | {"Взвешенно": "=" + "+".join(f"{sc.col(x)}{sc.r-1}*{sc.col(x)}{sc.r}" for x in SCN)}, fmt="0.00", bold=True, hl=True)
sc.row("sc_up", "Потенциал к цене $103,23", "%", {c: f"={R('sc_val',c)}/{R('price','2026E')}-1" for c in SCN + ["Взвешенно"]}, fmt="+0%;-0%")
sc.row("sc_eps27", "EPS 2027E", "$", {s: f"={R(s+'_eps','2027E')}" for s in SCN}, fmt="0.00")
sc.row("sc_m27", "Опер. маржа 2027E", "%", {s: f"={R(s+'_m','2027E')}" for s in SCN}, fmt="0.0%")
sc.row("sc_m32", "Опер. маржа 2032E (терминал)", "%", {s: f"={R(s+'_m','2032E')}" for s in SCN}, fmt="0.0%")
sc.row("sc_fuel27", "Топливо 2027E", "$/галл.", {s: f"={R(s+'_fpx','2027E')}" for s in SCN}, fmt="0.00")
sc.row("sc_psg", "DCF Гордон", "$", {s: f"={R(s+'_psg','2032E')}" for s in SCN}, fmt="0.00")
sc.row("sc_psm", "DCF выходной мультипликатор", "$", {s: f"={R(s+'_psm','2032E')}" for s in SCN}, fmt="0.00")
sc.blank()
TC = ["Значение"]
tg = S.__new__(S); tg.ws = sc.ws; tg.name = "Scenarios"; tg.cols = TC; tg.fc = 3; tg.r = sc.r; tg.srccol = sc.srccol
tg.header("Итог", "")
tg.row("t_target", "Целевая цена (взвешенная DCF, округл. до $1)", "$", {"Значение": f"=ROUND({R('sc_val','Взвешенно')},0)"}, fmt="0", bold=True, hl=True)
tg.row("t_low", "Нижняя граница диапазона = мин(DCF база, EV/EBITDAR ист., норм. P/E)", "$", {"Значение": f"=ROUND(MIN({R('Base_ps','2032E')},{R('im_hist27','Значение')},{R('im_normpe','Значение')}),0)"}, fmt="0")
tg.row("t_high", "Верхняя граница = макс(взвешенная DCF, EV/EBITDAR ист., норм. P/E)", "$", {"Значение": f"=ROUND(MAX({R('sc_val','Взвешенно')},{R('im_hist27','Значение')},{R('im_normpe','Значение')}),0)"}, fmt="0")
tg.row("t_up", "Потенциал к целевой цене", "%", {"Значение": f"={R('t_target','Значение')}/{R('price','2026E')}-1"}, fmt="+0.0%;-0.0%", bold=True)
tg.row("t_rating", "Рекомендация (правило: Buy > +15%, Sell < −10%)", "", {"Значение": f"=IF({R('t_up','Значение')}>{R('rule_buy','2026E')},\"BUY\",IF({R('t_up','Значение')}<{R('rule_sell','2026E')},\"SELL\",\"HOLD\"))"}, bold=True, hl=True)
tg.row("t_roll", "Справочно: взвешенная DCF через 12 мес. (× (1 + Ke))", "$", {"Значение": f"={R('sc_val','Взвешенно')}*(1+{R('ke','Значение')})"}, fmt="0.00", src="не используется в целевой цене (консервативно); UAL не платит дивиденды")
sc.r = tg.r

# ============================================================== SENSITIVITY
se = wb.create_sheet("Sensitivity")
se["A1"] = "Чувствительность (базовый сценарий)"; se["A1"].font = TITLE
se.column_dimensions["A"].width = 34
for j in range(2, 12): se.column_dimensions[L(j)].width = 11
# (a) WACC x g
se["A3"] = "1) Стоимость акции DCF (метод Гордона), $: WACC × терминальный рост"; se["A3"].font = BOLD
waccs = [0.075, 0.080, 0.085, 0.090, 0.095]; gs = [0.015, 0.020, 0.025, 0.030, 0.035]
se["A4"] = "WACC \\ g"; se["A4"].font = HDR; se["A4"].fill = HFILL
for j, gv in enumerate(gs):
    x = se.cell(4, 2 + j, gv); x.font = Font(name="Arial", size=9, color="0000FF", bold=True); x.number_format = "0.0%"
fcffd = [R("Base_fcffd", c) for c in YF]; tt = [R("Base_t", c) for c in YF]
nop = R("Base_nopat", "2032E"); ron = R("ronic", "2026E"); ndr = R("w_netdebt", "Значение"); nolr = R("nol", "2026E"); shr = R("shares", "2026E")
for i, wv in enumerate(waccs):
    r = 5 + i
    x = se.cell(r, 1, wv); x.font = Font(name="Arial", size=9, color="0000FF", bold=True); x.number_format = "0.0%"
    for j in range(len(gs)):
        wc = f"$A{r}"; gc = f"{L(2+j)}$4"
        pv = "+".join(f"{fcffd[k]}/(1+{wc})^{tt[k]}" for k in range(7))
        tv = f"{nop}*(1+{gc})*(1-{gc}/{ron})/({wc}-{gc})/(1+{wc})^6.5"
        cell = se.cell(r, 2 + j, f"=({pv}+{tv}-{ndr}+{nolr})/{shr}"); cell.number_format = "0"
REG["sens_wg"] = {"r0": 5, "c0": 2}
se["A11"] = "Метод Гордона; FCFF не зависят от WACC и g. Центральная ячейка (8,5% / 2,5%) ≈ значению Гордона в DCF при WACC 8,5%."; se["A11"].font = NOTE
# implied WACC at current price (Gordon, base g)
se["A13"] = "2) Reverse DCF по WACC: стоимость акции (Гордон, g = 2,5%) по сетке WACC"; se["A13"].font = BOLD
wg = [0.07, 0.075, 0.08, 0.085, 0.09, 0.095, 0.10, 0.105, 0.11]
se["A14"] = "WACC"; se["A15"] = "Стоимость, $"
for j, wv in enumerate(wg):
    x = se.cell(14, 2 + j, wv); x.font = Font(name="Arial", size=9, color="0000FF"); x.number_format = "0.0%"
    wc = f"{L(2+j)}$14"; gref = R("g", "2026E")
    pv = "+".join(f"{fcffd[k]}/(1+{wc})^{tt[k]}" for k in range(7))
    tv = f"{nop}*(1+{gref})*(1-{gref}/{ron})/({wc}-{gref})/(1+{wc})^6.5"
    se.cell(15, 2 + j, f"=({pv}+{tv}-{ndr}+{nolr})/{shr}").number_format = "0.0"
prc = R("price", "2026E")
parts = []
for j in range(len(wg) - 1):
    v0, v1 = f"{L(2+j)}15", f"{L(3+j)}15"; w0, w1 = f"{L(2+j)}14", f"{L(3+j)}14"
    parts.append(f"IF(AND({prc}<={v0},{prc}>={v1}),{w0}+({v0}-{prc})/({v0}-{v1})*({w1}-{w0}),")
se["A16"] = "WACC, заложенный в цену"; se["A16"].font = BOLD
se["B16"] = "=" + "".join(parts) + '"вне сетки"' + ")" * (len(wg) - 1); se["B16"].number_format = "0.00%"; se["B16"].fill = KFILL
REG["impl_wacc"] = {"v": "'Sensitivity'!$B$16"}
# (b) fuel x TRASM on 2027E EPS
se["A18"] = "3) EPS 2027E (база), $: цена топлива 2027E × отклонение TRASM 2027E от базы"; se["A18"].font = BOLD
fuels = [2.70, 3.00, 3.30, 3.60, 3.90, 4.20]; drs = [-0.04, -0.02, 0.0, 0.02, 0.04]
se["A19"] = "Топливо $/галл. \\ ΔTRASM"; se["A19"].font = HDR; se["A19"].fill = HFILL
for j, dv in enumerate(drs):
    x = se.cell(19, 2 + j, dv); x.font = Font(name="Arial", size=9, color="0000FF", bold=True); x.number_format = "+0%;-0%;0%"
asm27 = R("Base_asm", "2027E"); tr27 = R("Base_trasm", "2027E"); gpa27 = R("Base_gpa", "2027E"); fp27 = R("Base_fpx", "2027E")
pre27 = R("Base_pre", "2027E"); psr = R("psr", "2027E"); ni27 = R("netint", "2027E"); taxr = R("tax", "2026E"); tpr = R("tp", "2027E")
for i, fv in enumerate(fuels):
    r = 20 + i
    x = se.cell(r, 1, fv); x.font = Font(name="Arial", size=9, color="0000FF", bold=True); x.number_format = "0.00"
    for j in range(len(drs)):
        dc = f"{L(2+j)}$19"; fc = f"$A{r}"
        pre = f"({pre27}+{asm27}*{tr27}/100*{dc}*(1-{tpr})-{asm27}*{gpa27}/1000*({fc}-{fp27}))"
        se.cell(r, 2 + j, f"=({pre}-MAX(0,{pre})*{psr}-{ni27})*(1-{taxr})/{shr}").number_format = "0.00"
se["A26"] = "Перенос топлива в тариф в этой таблице не моделируется (статический эффект); в DCF перенос учитывается."; se["A26"].font = NOTE
# (c) +/-10% fuel, +/-1% RASM on operating income
se["A28"] = "4) Чувствительность операционной прибыли, $ млн (до участия в прибыли и без переноса в тариф)"; se["A28"].font = BOLD
hdrs = ["", "TTM 2К26", "2027E (база)"]
for j, hname in enumerate(hdrs):
    x = se.cell(29, 1 + j, hname); x.font = HDR; x.fill = HFILL
rows4 = [
 ("Топливо +10%", f"=-{R('fuel','TTM 2К26')}*0.1", f"=-{R('Base_fc','2027E')}*0.1"),
 ("Топливо −10%", f"={R('fuel','TTM 2К26')}*0.1", f"={R('Base_fc','2027E')}*0.1"),
 ("RASM +1%", f"={R('rev','TTM 2К26')}*0.01", f"={R('Base_rev','2027E')}*0.01"),
 ("RASM −1%", f"=-{R('rev','TTM 2К26')}*0.01", f"=-{R('Base_rev','2027E')}*0.01"),
 ("Справочно: скорр. опер. прибыль", f"={R('adj_opinc','TTM 2К26')}", f"={R('Base_ebit','2027E')}"),
 ("Топливо ±10% в % от опер. прибыли", f"=B30/B34", f"=C30/C34"),
 ("RASM ±1% в % от опер. прибыли", f"=B32/B34", f"=C32/C34"),
]
for i, (lab, f1, f2) in enumerate(rows4):
    r = 30 + i
    se.cell(r, 1, lab)
    c1 = se.cell(r, 2, f1); c2 = se.cell(r, 3, f2)
    fm = "0%" if "в %" in lab else "#,##0"
    c1.number_format = fm; c2.number_format = fm
REG["s4"] = {"fuel_ttm": "'Sensitivity'!$B$30", "fuel_27": "'Sensitivity'!$C$30", "rasm_ttm": "'Sensitivity'!$B$32", "rasm_27": "'Sensitivity'!$C$32",
             "fuel_pct_ttm": "'Sensitivity'!$B$35", "rasm_pct_ttm": "'Sensitivity'!$B$36", "adj_ttm": "'Sensitivity'!$B$34"}
se.sheet_view.showGridLines = False

# ============================================================== DEBT
db = S("Debt", "Долг, ликвидность, поставки самолётов", ["2026", "2027", "2028", "2029", "2030", "После 2030", "Итого"])
db.header("График денежных обязательств на 31.12.2025, $ млрд", "")
DCOL = ["2026", "2027", "2028", "2029", "2030", "После 2030"]
def drow(key, label, lst, src):
    v = dict(zip(DCOL, lst)); rr = db.r
    v["Итого"] = f"=SUM({db.col('2026')}{rr}:{db.col('После 2030')}{rr})"
    db.row(key, label, "$ млрд", v, fmt="0.0", src=src)
SRCD = "[Ф] 10-K 2025, Liquidity — material cash requirements"
drow("mat_debt", "Долг и проценты", [5.0, 2.8, 2.6, 3.6, 3.0, 9.0], SRCD)
drow("mat_fl", "Финансовая аренда", [0.1, 0.2, 0.1, 0, 0, 0], SRCD)
drow("mat_ol", "Операционная аренда", [0.9, 1.1, 0.9, 0.7, 0.8, 3.6], SRCD)
drow("mat_ofl", "Прочие фин. обязательства", [0.4, 0.6, 0.2, 0.2, 0.2, 2.7], SRCD)
drow("mat_capex", "Капитальные закупки (самолёты и пр.)", [12.6, 5.6, 7.4, 9.0, 8.7, 13.8], SRCD)
rr = db.r
db.row("mat_tot", "Итого долг + аренда", "$ млрд", {c: f"=SUM({db.col(c)}{rr-5}:{db.col(c)}{rr-2})" for c in DCOL + ["Итого"]}, fmt="0.0", bold=True)
db.note("После 31.12.2025: выпущены ноты $1 млрд 2029 (4,875%) и $1 млрд 2031 (5,375%); во 2К26 привлечено $3,7 млрд банковских кредитов, досрочно погашено ~$1 млрд дорогого долга (8-K 15.07.2026).")
db.blank()
LC = ["Значение"]
lq = S.__new__(S); lq.ws = db.ws; lq.name = "Debt"; lq.cols = LC; lq.fc = 3; lq.r = db.r; lq.srccol = db.srccol
lq.header("Ликвидность и рейтинги на 30.06.2026", "")
lq.row("liq", "Доступная ликвидность (вкл. невыбранные кредитные линии)", "$ млн", {"Значение": 19600}, src="[Ф] 8-K 15.07.2026")
lq.row("liq_rev", "Ликвидность / выручка TTM", "%", {"Значение": f"={R('liq','Значение')}/{R('rev','TTM 2К26')}"}, fmt="0%")
lq.row("rcf", "Револьверная линия (не выбрана на 31.12.2025)", "$ млн", {"Значение": 3000}, src="[Ф] 10-K 2025")
lq.row("covn", "Ковенант: минимальная ликвидность", "$ млн", {"Значение": 2000}, src="[Ф] 10-K 2025")
lq.row("rating", "Рейтинги S&P / Moody's / Fitch", "", {"Значение": "BB+ / Ba1 / BB+"}, src="[Ф] 10-Q 2К26: S&P BB+ (позитивный), Moody's Ba1 (стабильный), Fitch BB+ (стабильный); цель UAL — инвестиционный рейтинг в 2026")
lq.row("bb_left", "Остаток программы выкупа акций", "$ млн", {"Значение": 755}, src="[Ф] 10-Q 2К26 (на 09.07.2026); во 2К26 выкупа не было")
db.r = lq.r
db.blank()
EC = ["2022", "2023", "2024", "2025", "TTM 2К26"]
ev = S.__new__(S); ev.ws = db.ws; ev.name = "Debt"; ev.cols = EC; ev.fc = 3; ev.r = db.r; ev.srccol = db.srccol
ev.header("Долговая нагрузка: через EBITDAR и через денежный поток", "x")
ev.row("lv1", "Чистый скорр. долг / EBITDAR", "x", {c: f"={R('nd_ebitdar',c)}" for c in EC}, fmt="0.0")
ev.row("lv2", "Чистый скорр. долг / CFO", "x", {c: f"={R('nd_cfo',c)}" for c in EC}, fmt="0.0")
ev.row("lv3", "Чистый скорр. долг / CFO без прироста ATL и миль", "x", {c: f"={R('nd_cfoex',c)}" for c in EC[:4]}, fmt="0.0")
ev.row("lv4", "Покрытие фикс. платежей (EBITDAR / (% + аренда))", "x", {c: f"={R('covr',c)}" for c in EC}, fmt="0.0")
db.r = ev.r
db.note("Вывод: CFO у авиакомпании завышен притоком авансов (ATL) и нулевыми денежными налогами (NOL), поэтому долг/CFO выглядит мягче; надёжнее EBITDAR с поправкой на аренду, оценённый через цикл.")
db.blank()
FC = ["Контракт 2П26", "Контракт 2027", "Контракт после", "Ожид. 2П26", "Ожид. 2027", "Ожид. после"]
fl = S.__new__(S); fl.ws = db.ws; fl.name = "Debt"; fl.cols = FC; fl.fc = 3; fl.r = db.r; fl.srccol = db.srccol
for i in range(len(FC)): db.ws.column_dimensions[L(3 + i)].width = 13
fl.header("Поставки самолётов: контрактные vs ожидаемые (10-Q 2К26)", "шт.")
SRCF = "[Ф] 10-Q 2К26, Firm Order and Option Aircraft"
for k, lab, lst in [("f787", "Boeing 787", [42, 11, 93, 12, 27, 107]), ("fmax9", "Boeing 737 MAX 9", [63, 0, 0, 40, 23, 0]), ("fmax10", "Boeing 737 MAX 10", [3, 44, 120, 0, 20, 147]),
                    ("fa321", "Airbus A321neo", [8, 3, 100, 6, 5, 100]), ("fxlr", "Airbus A321XLR", [5, 15, 29, 5, 15, 29])]:
    fl.row(k, lab, "шт.", dict(zip(FC, lst)), src=SRCF)
rr = fl.r
fl.row("ftot", "Итого", "шт.", {c: f"=SUM({fl.col(c)}{rr-5}:{fl.col(c)}{rr-1})" for c in FC}, bold=True)
fl.row("fboe", "Boeing: ожидаемо / контрактно (2П26 + 2027)", "%", {"Ожид. 2П26": f"=SUM({fl.col('Ожид. 2П26')}{rr-5}:{fl.col('Ожид. 2027')}{rr-3})/SUM({fl.col('Контракт 2П26')}{rr-5}:{fl.col('Контракт 2027')}{rr-3})"}, fmt="0%", src="формула: доля контрактных поставок Boeing, которую UAL реально ожидает")
db.r = fl.r
db.note("A350: 45 твёрдых заказов без графика поставок (10-K 2025). MAX 10 не сертифицирован — поставки сдвинуты на 2027+.")

# order sheets
order = ["README", "Hist", "Metrics", "Assumptions", "Beta", "WACC", "DCF", "ReverseDCF", "Multiples", "Scenarios", "Sensitivity", "Debt"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != "Arial":
                c.font = Font(name="Arial", size=c.font.size if c.font and c.font.size else 9, bold=c.font.bold if c.font else False,
                              italic=c.font.italic if c.font else False, color=c.font.color if c.font else None)
wb.save(OUT)
json.dump(REG, open(OUT + ".reg.json", "w"), ensure_ascii=False, indent=0)
print("saved", OUT)
