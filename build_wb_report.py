# -*- coding: utf-8 -*-
"""
Строит из недельного детализированного отчёта Wildberries (reportDetailByPeriod,
выгрузка «в Excel» из портала) финальную книгу вида:
    СВОДНЫЕ ТАБЛИЦЫ | ОБЩИЙ ФАЙЛ | МАРТ | АПРЕЛЬ | ...
точь-в-точь как эталонные файлы СТВ_678457557 / СТВ_678457562.

Логика (ровно то, что делалось вручную):
  1. ОБЩИЙ ФАЙЛ  = все строки отчёта, сгруппированные по месяцу «Даты продажи»,
                   со сквозной нумерацией в колонке «№».
  2. <МЕСЯЦ>     = строки этого месяца (те же № и колонки).
  3. СВОДНЫЕ ТАБЛИЦЫ = сводная по «Обоснованию для оплаты» (суммы 7 денежных полей)
                   для ОБЩИЙ ФАЙЛ и для каждого месяца, плюс строка «нетто»,
                   «Итого к оплате» и блок «Поверка».

Итого к оплате = К перечислению(нетто) − доставка − штрафы − хранение − удержания − приёмка,
где нетто = Σ по категориям − 2×Возврат  (т.е. Продажа − Возврат + компенсации).

Запуск:  python3 build_wb_report.py вход_отчёт.xlsx выход.xlsx
"""
import sys, argparse
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.colors import Color
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


def _san(v):
    """убирает управляющие символы, которые Excel не принимает (напр. в кодах маркировки)."""
    if isinstance(v, str):
        return ILLEGAL_CHARACTERS_RE.sub('', v)
    return v

MONEY_FMT = '#,##0.00\\ "₽"'
YELLOW = PatternFill('solid', fgColor='FFFFFF00')
FILL_SRC   = PatternFill('solid', fgColor=Color(theme=7, tint=0.4))   # имя источника
FILL_HDR   = PatternFill('solid', fgColor=Color(theme=7, tint=0.8))   # заголовки "Сумма по полю"
FILL_CHECK = PatternFill('solid', fgColor=Color(theme=6, tint=0.8))   # блок "Поверка"
BOLD = Font(bold=True)

MONTHS_RU = {1:'ЯНВАРЬ',2:'ФЕВРАЛЬ',3:'МАРТ',4:'АПРЕЛЬ',5:'МАЙ',6:'ИЮНЬ',
             7:'ИЮЛЬ',8:'АВГУСТ',9:'СЕНТЯБРЬ',10:'ОКТЯБРЬ',11:'НОЯБРЬ',12:'ДЕКАБРЬ'}

# денежные колонки сводной: (заголовок в данных, заголовок "Сумма по полю ...")
PIVOT_COLS = [
    'Вайлдберриз реализовал Товар (Пр)',
    'К перечислению Продавцу за реализованный Товар',
    'Услуги по доставке товара покупателю',
    'Общая сумма штрафов',
    'Хранение',
    'Удержания',
    'Операции на приемке',
]
COL_OPER = 'Обоснование для оплаты'
COL_DATE = 'Дата продажи'
COL_NUM  = '№'
RETURN_LABEL = 'Возврат'

# ширина столбцов листов данных (как в эталонных файлах) — чтобы не было "#####"
DATA_COL_WIDTHS = {
    'A':10,'B':25,'C':18,'D':19,'E':18,'F':23,'G':100,'H':14,'I':20,'J':17,'K':18,
    'N':16,'O':19,'P':23,'Q':25,'R':20,'S':22,'T':25,'U':35,'W':21,'X':14,'Y':18,
    'Z':16,'AA':33,'AC':30,'AD':40,'AE':27,'AF':23,'AG':27,'AH':18,'AJ':24,'AK':35,
    'AN':16,'AP':24,'AQ':19,'AR':30,'AS':16,'AT':30,'AU':20,'AW':19,'AX':16,'AY':20,
    'AZ':25,'BA':15,'BC':32,'BD':30,'BF':40,'BG':60,'BH':30,'BI':35,'BJ':45,'BK':40,
    'BM':45,'BN':40,'BT':50,'BU':40,'BW':50,'BX':30,
}


def month_of(v):
    """месяц из значения «Даты продажи» (строка 'YYYY-MM-DD' или datetime)."""
    if v is None or v == '':
        return None
    if hasattr(v, 'month'):
        return v.month
    s = str(v)
    try:
        return int(s[5:7])
    except Exception:
        return None


def read_report(path):
    """читает входной отчёт: возвращает (headers, rows) с листа, где есть шапка."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        try:
            hdr = next(it)
        except StopIteration:
            continue
        hdr = list(hdr)
        if COL_OPER in hdr and COL_NUM in hdr:
            rows = [list(r) for r in it if any(c is not None for c in r)]
            wb.close()
            return hdr, rows
    wb.close()
    raise SystemExit('Не нашёл лист с шапкой отчёта (колонки «№» и «Обоснование для оплаты»)')


def pivot(rows, idx):
    """суммы 7 денежных полей по категориям; категории по алфавиту."""
    oper_i = idx[COL_OPER]
    val_i = [idx[c] for c in PIVOT_COLS]
    agg = {}
    for r in rows:
        cat = r[oper_i]
        a = agg.setdefault(cat, [0.0]*7)
        for k, ci in enumerate(val_i):
            v = r[ci]
            if isinstance(v, (int, float)):
                a[k] += v
    return [(cat, agg[cat]) for cat in sorted(agg)]


def write_data_sheet(ws, headers, rows):
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
    for r in rows:
        ws.append([_san(v) for v in r])
    for col, w in DATA_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w


def write_pivot_block(ws, r0, title, cats):
    """пишет блок сводной начиная со строки r0. Возвращает (row_net, row_itogo, next_row)."""
    from openpyxl.utils import get_column_letter
    # шапка
    ws.cell(r0, 1, title).fill = FILL_SRC
    ws.cell(r0, 1).font = BOLD
    for k, name in enumerate(PIVOT_COLS):
        c = ws.cell(r0, 2+k, 'Сумма по полю ' + name)
        c.fill = FILL_HDR; c.font = BOLD
    # категории
    ret_row = None
    first = r0 + 1
    for j, (cat, sums) in enumerate(cats):
        rr = first + j
        ws.cell(rr, 1, cat)
        if cat == RETURN_LABEL:
            ret_row = rr
        for k in range(7):
            c = ws.cell(rr, 2+k, round(sums[k], 2))
            c.number_format = MONEY_FMT
    last = first + len(cats) - 1
    # строка "нетто" (после пустой строки)
    r_net = last + 2
    ws.cell(r_net, 1, 'Продажа')
    for k in range(7):
        col = get_column_letter(2+k)
        f = f'=SUM({col}{first}:{col}{last})'
        if k < 2 and ret_row:                 # реализовал и К перечислению: минус 2×Возврат
            f += f'-2*{col}{ret_row}'
        c = ws.cell(r_net, 2+k, f)
        c.number_format = MONEY_FMT
    ws.cell(r_net, 2).fill = YELLOW            # «оборот» — жёлтым
    # Итого к оплате
    r_it = r_net + 1
    ws.cell(r_it, 1, 'Итого к оплате')
    c = ws.cell(r_it, 2, f'=C{r_net}-D{r_net}-E{r_net}-F{r_net}-G{r_net}-H{r_net}')
    c.number_format = MONEY_FMT
    return r_net, r_it, r_it + 2


def build(headers, rows, out_path):
    idx = {h: i for i, h in enumerate(headers) if h is not None}
    num_i = idx[COL_NUM]; date_i = idx[COL_DATE]

    # группировка по месяцам, сохраняя исходный порядок внутри месяца
    months = []
    by_month = {}
    for r in rows:
        m = month_of(r[date_i])
        if m not in by_month:
            by_month[m] = []; months.append(m)
    months = sorted([m for m in months if m is not None])
    for r in rows:
        by_month[month_of(r[date_i])].append(r)

    # ОБЩИЙ ФАЙЛ = месяцы по порядку, сквозная нумерация №
    ordered = []
    for m in months:
        ordered.extend(by_month[m])
    n = 1
    for r in ordered:
        r[num_i] = n; n += 1

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_piv = wb.create_sheet('СВОДНЫЕ ТАБЛИЦЫ')
    write_data_sheet(wb.create_sheet('ОБЩИЙ ФАЙЛ'), headers, ordered)
    for m in months:
        write_data_sheet(wb.create_sheet(MONTHS_RU[m]), headers, by_month[m])

    # ---- СВОДНЫЕ ТАБЛИЦЫ ----
    # блок ОБЩИЙ ФАЙЛ (строка 1)
    net_o, it_o, nxt = write_pivot_block(ws_piv, 1, 'ОБЩИЙ ФАЙЛ', pivot(ordered, idx))
    # Поверка (заполним ссылки после расчёта месячных блоков)
    r_check = nxt
    # месячные блоки идут ниже Поверки (Поверка = 3 строки + пустая)
    r_month0 = r_check + 4
    rr = r_month0
    net_rows = []
    for m in months:
        net_m, it_m, rr = write_pivot_block(ws_piv, rr, MONTHS_RU[m], pivot(by_month[m], idx))
        net_rows.append((net_m, it_m))
    # Поверка
    ws_piv.cell(r_check, 1, 'Поверка').fill = FILL_CHECK
    ws_piv.cell(r_check, 1).font = BOLD
    cP = ws_piv.cell(r_check+1, 1, 'Продажа'); cP.fill = FILL_CHECK
    cPv = ws_piv.cell(r_check+1, 2, '=' + '+'.join(f'B{nr}' for nr, _ in net_rows))
    cPv.number_format = MONEY_FMT; cPv.fill = FILL_CHECK
    cI = ws_piv.cell(r_check+2, 1, 'Итого к оплате'); cI.fill = FILL_CHECK
    cIv = ws_piv.cell(r_check+2, 2, '=' + '+'.join(f'B{ir}' for _, ir in net_rows))
    cIv.number_format = MONEY_FMT; cIv.fill = FILL_CHECK

    # ширина столбцов сводной: A под подписи, B..H под суммы с "₽" (до миллионов)
    ws_piv.column_dimensions['A'].width = 20
    for col in 'BCDEFGH':
        ws_piv.column_dimensions[col].width = 17

    wb.save(out_path)
    print(f'Готово: {out_path}')
    print(f'  строк всего: {len(ordered)}   месяцы: {", ".join(MONTHS_RU[m] for m in months)}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('input')
    ap.add_argument('output')
    a = ap.parse_args()
    hdr, rows = read_report(a.input)
    build(hdr, rows, a.output)
