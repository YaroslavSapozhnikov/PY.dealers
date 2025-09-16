from typing import Iterator, IO
from datetime import date, time, datetime, timedelta
import openpyxl
# import time


def iter_excel(file: IO[bytes]) -> Iterator[dict[str, object]]:
    workbook = openpyxl.load_workbook(file)
    rows = workbook.active.rows
    headers = [str(cell.value) for cell in next(rows)]
    for row in rows:
        yield dict(zip(headers, (cell.value for cell in row)))


def last_auth(dilers: list[dict]):
    last_t = 0
    for dealer in dealers:
        t = dealer.get('Последняя авторизация')
        if t != 'None':
            dt = datetime.strptime(t, "%Y.%m.%d %H:%M:%S")
            unix_t = int(dt.timestamp())
            if last_t < unix_t:
                last_t = unix_t
    return datetime.fromtimestamp(last_t).strftime('%Y.%m.%d %H:%M:%S')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    workbook = openpyxl.load_workbook(filename="dealers.xlsx")
    sheet = workbook.active
    rows = sheet.rows
    headers = tuple(str(cell.value) for cell in next(rows))

    dealers = [{k: v for k, v in zip(headers, tuple(str(cell.value) for cell in row))} for row in rows]
    for dealer in dealers:
        print(dealer)

    print(last_auth(dealers))

    # with open('dealers.xlsx', 'rb') as f:
    #     rows = iter_excel(f)
    #     row = next(rows)
    #     print(row)
    #     row = next(rows)
    #     print(row)
