from datetime import date, datetime
from environs import env
import openpyxl

headers_col_dic = dict([])


def change_dealer_rec(sheet, login, t):
    global headers_col_dic
    for i in range(len(dealers)):
        if dealers[i]['Логин'] == login:
            cell_last_auth = headers_col_dic['Последняя авторизация'] + str(i + 2)
            old_dt = sheet[cell_last_auth].value
            if old_dt is None:
                old_date = date.fromtimestamp(0)
            else:
                old_date = datetime.strptime(old_dt, "%Y.%m.%d %H:%M:%S").date()
            new_date = date.fromtimestamp(t)
            sheet[cell_last_auth] = datetime.fromtimestamp(t).strftime('%Y.%m.%d %H:%M:%S')
            cell_cnt_auth = headers_col_dic['Кол-во авторизаций'] + str(i + 2)
            cnt = sheet[cell_cnt_auth].value
            if cnt is None:
                cnt = 0
            else:
                cnt = int(cnt)
            if old_date != new_date:
                cnt += 1
            sheet[cell_cnt_auth] = cnt


def get_last_auth(dilers: list[dict]):
    last_t = 0
    for dealer in dealers:
        t = dealer.get('Последняя авторизация')
        if t != 'None':
            dt = datetime.strptime(t, "%Y.%m.%d %H:%M:%S")
            unix_t = int(dt.timestamp())
            if last_t < unix_t:
                last_t = unix_t
    return last_t
    # return datetime.fromtimestamp(last_t).strftime('%Y.%m.%d %H:%M:%S')


if __name__ == '__main__':
    env.read_env()
    log_file_name = env("LOGFILE")
    dealers_file_name = env("DEALERSFILE")

    file_changed = False
    workbook = openpyxl.load_workbook(filename=dealers_file_name, read_only=False)
    sheet = workbook.active
    rows = sheet.rows
    headers = tuple(str(cell.value) for cell in next(rows))
    headers_col_dic = {headers[i]: chr(ord('A') + i) for i in range(len(headers))}

    dealers = [{k: v for k, v in zip(headers, tuple(str(cell.value) for cell in row))} for row in rows]
    # for dealer in dealers:
    #     print(dealer)

    last_auth = get_last_auth(dealers)

    with open(log_file_name, 'r') as log:
        while True:
            line = log.readline()
            if not line:
                break
            rec = line.split(maxsplit=5)
            if int(rec[0]) <= last_auth:
                continue
            if 'logged in' in rec[5]:
                change_dealer_rec(sheet, rec[4], int(rec[0]))
                file_changed = True

    if file_changed:
        workbook.save(filename="dealers.xlsx")

