#-------------------------------------------------------------------------------
# Author:      ГавриловАВ
# Copyright:   (c) ГавриловАВ 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import xlwings as xw
import requests
import datetime
import os

def main(numStr):
    def sendReq(name, nDev, due):
        # функция отправки запроса API Trello
        url = "https://api.trello.com/1/cards"
        # доска КТО '5f7f5031a39ace11ae0cefea'
        # ID Пшеничников 5f97f43e32ca2328de00534f
        # ID Карпушкина  5f9a736a41543e48b873059c

        query = {
           'key': '04403aaa312416004ec400c73c48a811',
           'token': '51fe87fbdffd3b558b342bc6e6eb43e96c6e6f26b850c25956aaf5212f515ebc',
           'idList': '5f7f5031a39ace11ae0cefea',
           'name' : name,
           'desc' : nDev,
           'due' : due,
            'idMembers' : '5f97f43e32ca2328de00534f',
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
        # получаем дату завершения
        due = datetime.datetime.now() + datetime.timedelta(days=1)
        print(f"Заголовок карточки: {name}")
        print(f"Описание: {nDev}")
        print(f"Дата завершения: {due}")
        sendReq(name, nDev, due)

    crt_msg(numStr)

# если скрипт запускается напрямую, без использования VBA, используем тестовую
# книгу рядом со скриптом
if __name__ == '__main__':
    wb = xw.Book('Test2.xlsm').set_mock_caller()
    wb = xw.Book.caller()
    #numStr = str(int(wb.sheets[1].range('B12').value))
    numStr = f"{int(wb.sheets[1].range('B12').value)}"
    main(numStr)






