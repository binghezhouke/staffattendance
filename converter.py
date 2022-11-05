from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime, timedelta
import sys

'''表格格式不变时通常不需要换,代表有效数据起始的行和列'''
ROW_START = 5
COL_START = 4


def judge(cell):
    start_time = datetime.strptime("10:00", "%H:%M")
    end_time = datetime.strptime("19:00", "%H:%M")
    font_late = Font(color="FF0000")
    font_too_early = Font(color="00FF00")
    font_bonus = Font(color="912CEE")

    bonus_delta = timedelta(hours=11)
    should_give_bonus = False
    s = cell.value
    v = s.split(r";")

    if len(v) == 3:
        # 存在正常的上下班打卡记录
        s = v[0].strip()
        start = datetime.strptime(s, "%H:%M")
        if start > start_time:
            cell.font = font_late
        e = v[1].strip()
        if "次日" in e:
            # 异常数据处理
            e = "23:59"
        end = datetime.strptime(e, "%H:%M")
        if end < end_time:
            cell.font = font_too_early
        if end-start >= bonus_delta:
            should_give_bonus = True
            cell.font = font_bonus
        cell.value = "{}\n{}".format(s, e)
    else:
        if len(v) == 2:
            cell.font = font_late
        cell.value = "{}\n".format(v[0].strip())

    return should_give_bonus


def find_holiday(ws, row_max, col_max):
    holidays_idx = []
    for c in range(COL_START, col_max+1):
        holiday_cnt = len(
            [r for r in range(ROW_START, row_max+1) if "休息" in ws.cell(r, c).value])
        if holiday_cnt >= 1:
            holidays_idx.append(c)
    print("本月休息日为:")
    print([x-3 for x in holidays_idx])
    return holidays_idx


def handle(in_file, out_file):
    wb = load_workbook(in_file)
    ws = wb[wb.sheetnames[0]]
    row_max, col_max = ws.max_row, ws.max_column

    holiday_idx = find_holiday(ws, row_max, col_max)
    for r in range(ROW_START, row_max+1):
        print(ws.cell(r, 1).value)
        bonus_cnt = 0
        for c in range(COL_START, col_max+1):
            cell = ws.cell(r, c)
            should_give_bonus = judge(cell)
            if should_give_bonus and c not in holiday_idx:
                bonus_cnt += 1
        ws.cell(r, col_max+1).value = bonus_cnt

    wb.save(out_file)


if __name__ == '__main__':
    if len(sys.argv) == 3:
        in_file = sys.argv[1]
        out_file = sys.argv[2]
    else:
        in_file = "a.xlsx"
        out_file = "b.xlsx"
    print(in_file, out_file)
    handle(in_file, out_file)
