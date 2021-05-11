import xlwings as xw
import time 


EXCEL = 'animation_py.xlsx'
CANTS = [
    0, 30, 60, 75, 60, 30, 0, -30, -60, -75, -60, -30,
    0, 30, 60, 75, 60, 30, 0, -30, -60, -75, -60, -30, 0
    ]
N_CANTS = len(CANTS)


def run():
    wb = xw.Book(EXCEL)
    sheet = wb.sheets[0]
    count = 0
    while (count < N_CANTS):
        sheet.range('F7').value = CANTS[count]
        time.sleep(1.5)
        count += 1


def main():
    run()

if __name__ == "__main__":
    main()