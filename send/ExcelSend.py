#-------------------------------------------------------------------------------
# Author:      ГавриловАВ
# Copyright:   (c) ГавриловАВ 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import xlwings as xw
import requests
import os

def main(numStr):
    def sendReq(name, nDev):
        # функция отправки запроса API Trello
        url = "https://api.trello.com/1/cards"
        # доска КТО '5f7f5031a39ace11ae0cefea'
        query = {
           'key': '04403aaa312416004ec400c73c48a811',
           'token': '51fe87fbdffd3b558b342bc6e6eb43e96c6e6f26b850c25956aaf5212f515ebc',
           'idList': '5f843bbdfa40fd77f31f414c',
           'name' : name,
           'desc' : nDev,
        }
        response = requests.request(
           "POST",
           url,
           params=query
        )
        print(response.text)
    def crt_msg(numStr):
        # функция формирования текста запроса для отправки
        wb = xw.Book.caller()
        cellNdoc, cellProd, cellDev= f"B{numStr}", f"C{numStr}", f"F{numStr}"
        nDoc = f"{int(wb.sheets[1].range(cellNdoc).value)}"
        # наименование изделия
        nProd = f"{wb.sheets[1].range(cellProd).value}"
        # текст отклонения (описание карточки)
        nDev = f"{wb.sheets[1].range(cellDev).value}"
        # заголовок карточки
        name = f"КВР №{nDoc} {nProd}"
        print(f"Заголовок карточки: {name}")
        print(f"Описание: {nDev}")
        sendReq(name, nDev)

    crt_msg(numStr)

# если скрипт запускается напрямую, без использования VBA, используем тестовую
# книгу рядом со скриптом
if __name__ == '__main__':
    wb = xw.Book('Test2.xlsm').set_mock_caller()
    wb = xw.Book.caller()
    #numStr = str(int(wb.sheets[1].range('B12').value))
    numStr = f"{int(wb.sheets[1].range('B12').value)}"
    main(numStr)






