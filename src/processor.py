#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import uuid
import pytz
import xlrd
from xml.dom import minidom
from datetime import datetime

# Устанавливаем стандартную кодировку
reload(sys)
sys.setdefaultencoding('utf8')

# Полезные переменные
local_tz = pytz.timezone('Europe/Moscow')

# Полезные функции
def utc_to_local(utc_dt):
    local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
    return local_tz.normalize(local_dt)

def insert_str(string, str_to_insert, index):
    return string[:index] + str_to_insert + string[index:]

# Открываем рабочий файл
workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_index(0)

# Создаём коренной элемент
doc = minidom.Document()
root = doc.createElement('ns2:annualBalanceF0503730_2015')
doc.appendChild(root)

# Header
header = doc.createElement('header')
root.appendChild(header)

# ID
id = doc.createElement('id')
id.appendChild(doc.createTextNode(str(uuid.uuid4())))
header.appendChild(id)

# createDateTime
createDateTime = doc.createElement('createDateTime')
createDateTime.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
header.appendChild(createDateTime)

# ns2:body
ns2_body = doc.createElement('ns2:body')
root.appendChild(ns2_body)

# ns2:position
ns2_position = doc.createElement('ns2:position')
ns2_body.appendChild(ns2_position)

# positionId
positionId = doc.createElement('positionId')
positionId.appendChild(doc.createTextNode(str(uuid.uuid4())))
ns2_position.appendChild(positionId)

# changeDate
changeDate = doc.createElement('changeDate')
changeDate.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
ns2_position.appendChild(changeDate)

# placer
placer = doc.createElement('placer')
ns2_position.appendChild(placer)

# regNum
placer_regNum = doc.createElement('regNum')
placer_regNum.appendChild(doc.createTextNode('462D1140'))
placer.appendChild(placer_regNum)

# fullName
placer_fullName = doc.createElement('fullName')
placer_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
placer.appendChild(placer_fullName)

# inn
placer_inn = doc.createElement('inn')
placer_inn.appendChild(doc.createTextNode('5022049898'))
placer.appendChild(placer_inn)

# kpp
placer_kpp = doc.createElement('kpp')
placer_kpp.appendChild(doc.createTextNode('502201001'))
placer.appendChild(placer_kpp)

# initiator
initiator = doc.createElement('initiator')
ns2_position.appendChild(initiator)

# regNum
initiator_regNum = doc.createElement('regNum')
initiator_regNum.appendChild(doc.createTextNode('462D1140'))
initiator.appendChild(initiator_regNum)

# fullName
initiator_fullName = doc.createElement('fullName')
initiator_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
initiator.appendChild(initiator_fullName)

# inn
initiator_inn = doc.createElement('inn')
initiator_inn.appendChild(doc.createTextNode('5022049898'))
initiator.appendChild(initiator_inn)

# kpp
initiator_kpp = doc.createElement('kpp')
initiator_kpp.appendChild(doc.createTextNode('502201001'))
initiator.appendChild(initiator_kpp)

# versionNumber
versionNumber = doc.createElement('versionNumber')
versionNumber.appendChild(doc.createTextNode('0'))
ns2_position.appendChild(versionNumber)

# now
now = datetime.now()

# formationPeriod
formationPeriod = doc.createElement('formationPeriod')
formationPeriod.appendChild(doc.createTextNode(str(now.year)))
ns2_position.appendChild(formationPeriod)

# generalData
generalData = doc.createElement('generalData')
ns2_position.appendChild(generalData)

# date
date = doc.createElement('date')
date.appendChild(doc.createTextNode(str(now.year - 1) + "-01-01+03:00"))
generalData.appendChild(date)

# periodicity
periodicity = doc.createElement('periodicity')
periodicity.appendChild(doc.createTextNode('annual'))
generalData.appendChild(periodicity)

# okei
okei = doc.createElement('okei')
generalData.appendChild(okei)

# code
code = doc.createElement('code')
code.appendChild(doc.createTextNode(worksheet.cell(13, 13).value))
okei.appendChild(code)

# symbol
symbol = doc.createElement('symbol')
symbol.appendChild(doc.createTextNode('руб'))
okei.appendChild(symbol)

# okpo
okpo = doc.createElement('okpo')
okpo.appendChild(doc.createTextNode(worksheet.cell(5, 13).value))
generalData.appendChild(okpo)

# inn
inn = doc.createElement('inn')
inn.appendChild(doc.createTextNode(worksheet.cell(6, 13).value))
generalData.appendChild(inn)

# oktmo
oktmo = doc.createElement('oktmo')
generalData.appendChild(oktmo)

# code
code = doc.createElement('code')
code.appendChild(doc.createTextNode(worksheet.cell(8, 13).value))
oktmo.appendChild(code)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('Коломна'))
oktmo.appendChild(name)

# founderAuthority
founderAuthority = doc.createElement('founderAuthority')
generalData.appendChild(founderAuthority)

# regNum
regNum = doc.createElement('regNum')
regNum.appendChild(doc.createTextNode('46200077'))
founderAuthority.appendChild(regNum)

# fullName
fullName = doc.createElement('fullName')
fullName.appendChild(doc.createTextNode('МИНИСТЕРСТВО ОБРАЗОВАНИЯ МОСКОВСКОЙ ОБЛАСТИ'))
founderAuthority.appendChild(fullName)

# founderAuthorityOkpo
founderAuthorityOkpo = doc.createElement('founderAuthorityOkpo')
founderAuthorityOkpo.appendChild(doc.createTextNode(worksheet.cell(9, 13).value))
generalData.appendChild(founderAuthorityOkpo)

#
# НЕФИНАНСОВЫЕ АКТИВЫ
#

#
# Сущность <reportItem> может включать в себя сущность <reportSubItem>. Каждый такой экземпляр сущности в свою очередь
# включает в себя следующие сущности:
#
#    <targetFundsStartYear>	- деятельность с целевыми средствами на начало года
#    <targetFundsEndYear>	- деятельность с целевыми средствами на конец отчётного периода
#    <stateTaskFundsStartYear>	- деятельность по государственному заданию на начало года
#    <stateTaskFundsEndYear>	- деятельность по государственному заданию на конец отчётного периода
#    <revenueFundsStartYear>	- приносящая доход деятельность на начало года
#    <revenueFundsEndYear>	- приносящая доход деятельность на конец отчётного периода
#    <totalStartYear>		- итого на начало года
#    <totalEndYear>		- итого на конец отчётного периода
#

#
# Алгоритм генерации сущностей типа <reportItem> и <reportSubItem>
#
# 1. Сущность <reportItem> может включать в себя сущности <reportSubItem>, которые по своему составу
#    абсолютно идентичны
# 2. Ищем строку "АКТИВ". Вычисляем сдвиг: положение ячейки "АКТИВ" + 3 = первый элемент
# 3. Ищем последний элемент. Его поиск производится следующим образом:
#    ЕСЛИ значениеТекущейЯчейки(x) ПУСТО И значениеТекущейЯчейки(x) + 1 ПУСТО, ТОГДА КОНЕЦ
#    Предыдущая ячейка(y) будет последней
# 4. Теперь мы получили значение range(ВСЕГО_ЯЧЕЕК). Обрабатываем каждую строку:
#    1. ЕСЛИ значениеТекущейЯчейки(x) НЕ ПУСТО:
#       КОД_СТРОКИ    = СТОЛБЕЦ(1)
#       НАЧАЛО_ГОДА_1 = СТОЛБЕЦ(2)
#       НАЧАЛО_ГОДА_2 = СТОЛБЕЦ(3)
#       НАЧАЛО_ГОДА_3 = СТОЛБЕЦ(4)
#       НАЧАЛО_ГОДА_4 = СТОЛБЕЦ(5)
#       КОНЕЦ_ГОДА_1  = СТОЛБЕЦ(2)
#       КОНЕЦ_ГОДА_2  = СТОЛБЕЦ(3)
#       КОНЕЦ_ГОДА_3  = СТОЛБЕЦ(4)
#       КОНЕЦ_ГОДА_4  = СТОЛБЕЦ(5)
#    2. текущаяЯчейка(x) = текущаяЯчейка(x) + 1
#    3. ЕСЛИ значениеТекущейЯчейки(x) НЕ ПУСТО:
#       ...
#

# nonFinancialAssets
nonFinancialAssets = doc.createElement('nonFinancialAssets')
ns2_position.appendChild(nonFinancialAssets)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while not "I. Нефинансовые активы" in elementValue:
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 1

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while True:
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition
savedLastElementPosition = lastElementPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)

    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        if not "Итого по разделу" in str(worksheet.cell(currentPosition, 0).value):
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value)))
            reportItem.appendChild(name)
        else:
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value).replace('\n', ' ')))
            reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к nonFinancialAssets
        nonFinancialAssets.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# Обработка "АКТИВ"
#

# Поиск первого элемента
currentPosition = savedLastElementPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "АКТИВ":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 3

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while not "II. Финансовые активы" in str(worksheet.cell(currentPosition, cell).value):
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)
    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        if not "Итого по разделу" in str(worksheet.cell(currentPosition, 0).value):
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value)))
            reportItem.appendChild(name)
        else:
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value).replace('\n', ' ')))
            reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к nonFinancialAssets
        nonFinancialAssets.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# ФИНАНСОВЫЕ АКТИВЫ
#

#
# Сущность <reportItem> может включать в себя сущность <reportSubItem>. Каждый такой экземпляр сущности в свою очередь
# включает в себя следующие сущности:
#
#    <targetFundsStartYear>	- деятельность с целевыми средствами на начало года
#    <targetFundsEndYear>	- деятельность с целевыми средствами на конец отчётного периода
#    <stateTaskFundsStartYear>	- деятельность по государственному заданию на начало года
#    <stateTaskFundsEndYear>	- деятельность по государственному заданию на конец отчётного периода
#    <revenueFundsStartYear>	- приносящая доход деятельность на начало года
#    <revenueFundsEndYear>	- приносящая доход деятельность на конец отчётного периода
#    <totalStartYear>		- итого на начало года
#    <totalEndYear>		- итого на конец отчётного периода
#

# financialAssets
financialAssets = doc.createElement('financialAssets')
ns2_position.appendChild(financialAssets)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "II. Финансовые активы":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 1

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while True:
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition
savedLastElementPosition = lastElementPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)

    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        name = doc.createElement('name')
        name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value)))
        reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к financialAssets
        financialAssets.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# Обработка "АКТИВ"
#

# Поиск первого элемента
currentPosition = savedLastElementPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "АКТИВ":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 3

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while True:
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)
    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        if not "Итого по разделу" in str(worksheet.cell(currentPosition, 0).value):
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value)))
            reportItem.appendChild(name)
        else:
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value).replace('\n', ' ')))
            reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к financialAssets
        financialAssets.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# ОБЯЗАТЕЛЬСТВА
#

#
# Сущность <reportItem> может включать в себя сущность <reportSubItem>. Каждый такой экземпляр сущности в свою очередь
# включает в себя следующие сущности:
#
#    <targetFundsStartYear>	- деятельность с целевыми средствами на начало года
#    <targetFundsEndYear>	- деятельность с целевыми средствами на конец отчётного периода
#    <stateTaskFundsStartYear>	- деятельность по государственному заданию на начало года
#    <stateTaskFundsEndYear>	- деятельность по государственному заданию на конец отчётного периода
#    <revenueFundsStartYear>	- приносящая доход деятельность на начало года
#    <revenueFundsEndYear>	- приносящая доход деятельность на конец отчётного периода
#    <totalStartYear>		- итого на начало года
#    <totalEndYear>		- итого на конец отчётного периода
#

# commitments
commitments = doc.createElement('commitments')
ns2_position.appendChild(commitments)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while not "III. Обязательства" in elementValue:
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 1

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while not "IV. Финансовый результат" in str(worksheet.cell(currentPosition, cell).value):
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition
savedLastElementPosition = lastElementPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)

    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        if "Итого по разделу" in str(worksheet.cell(currentPosition, 0).value):
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode("Итого по разделу III" + " " + str(worksheet.cell(currentPosition, 0).value).splitlines()[1]))
            reportItem.appendChild(name)
        else:
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value).replace('\n', ' ')))
            reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к commitments
        commitments.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# ФИНАНСОВЫЙ РЕЗУЛЬТАТ
#

#
# Сущность <reportItem> может включать в себя сущность <reportSubItem>. Каждый такой экземпляр сущности в свою очередь
# включает в себя следующие сущности:
#
#    <targetFundsStartYear>	- деятельность с целевыми средствами на начало года
#    <targetFundsEndYear>	- деятельность с целевыми средствами на конец отчётного периода
#    <stateTaskFundsStartYear>	- деятельность по государственному заданию на начало года
#    <stateTaskFundsEndYear>	- деятельность по государственному заданию на конец отчётного периода
#    <revenueFundsStartYear>	- приносящая доход деятельность на начало года
#    <revenueFundsEndYear>	- приносящая доход деятельность на конец отчётного периода
#    <totalStartYear>		- итого на начало года
#    <totalEndYear>		- итого на конец отчётного периода
#

# financialResult
financialResult = doc.createElement('financialResult')
ns2_position.appendChild(financialResult)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while not "IV. Финансовый результат" in elementValue:
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition + 1

# Поиск последнего элемента
currentPosition = firstElementPosition; cell = 0; count = 0
while True:
    for element in range(10):
        value = str(worksheet.cell(currentPosition, cell).value)
        if value.strip():
            count += 1;
        cell += 1
    if count == 0:
        break
    currentPosition += 1; cell = 0; count = 0
lastElementPosition = currentPosition
savedLastElementPosition = lastElementPosition

# Количество элементов
total = lastElementPosition - firstElementPosition

# Текущая позиция
currentPosition = firstElementPosition

# Обработка строк в соответствии с документацией ТФФ ГМУ
for line in range(total):
    # Если элемент находится в ячейке 0, то он - reportItem,
    # а если > 0 - reportSubItem, относящийся к reportItem
    cell = 0
    currentElementValue = str(worksheet.cell(currentPosition, cell).value)

    if str(worksheet.cell(currentPosition, 5).value)[-1] == "0":
        # reportItem
        reportItem = doc.createElement('reportItem')

        # name
        name = doc.createElement('name')
        name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 0).value)))
        reportItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportItem.appendChild(totalEndYear)

        # Добавляем к financialResult
        financialResult.appendChild(reportItem)
    else:
        # reportSubItem
        reportSubItem = doc.createElement('reportSubItem')

        # name
        name = doc.createElement('name'); currentCell = 1

        # Ищем название, попутно исключая пустые ячейки
        while not str(worksheet.cell(currentPosition, currentCell).value).strip():
            currentCell += 1
        try:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
        except IndexError:
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
        reportSubItem.appendChild(name)

        # lineCode
        lineCode = doc.createElement('lineCode')
        lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value)))
        reportSubItem.appendChild(lineCode)

        # targetFundsStartYear
        targetFundsStartYear = doc.createElement('targetFundsStartYear')
        targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsStartYear)

        # targetFundsEndYear
        targetFundsEndYear = doc.createElement('targetFundsEndYear')
        targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(targetFundsEndYear)

        # stateTaskFundsStartYear
        stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
        stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsStartYear)

        # stateTaskFundsEndYear
        stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
        stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0').replace('X', '0')))
        reportSubItem.appendChild(stateTaskFundsEndYear)

        # revenueFundsStartYear
        revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
        revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsStartYear)

        # revenueFundsEndYear
        revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
        revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 12).value).replace('0,00', '0')))
        reportSubItem.appendChild(revenueFundsEndYear)

        # totalStartYear
        totalStartYear = doc.createElement('totalStartYear')
        totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalStartYear)

        # totalEndYear
        totalEndYear = doc.createElement('totalEndYear')
        totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 13).value).replace('0,00', '0')))
        reportSubItem.appendChild(totalEndYear)

        # Добавляем к reportItem
        reportItem.appendChild(reportSubItem)

    # Переходим к следующей строке
    currentPosition += 1

#
# СПРАВКА
#

# Переходим на вторую страницу
worksheet = workbook.sheet_by_index(1)

# reference
reference = doc.createElement('reference')
ns2_position.appendChild(reference)

# Инициализируем переменные
savedFirstElementPosition = 0; total = 0

# Итерируем прогон цикла три раза
for i in range(3):

    # Поиск первого элемента
    currentPosition = savedFirstElementPosition + total; elementValue = str(worksheet.cell(currentPosition, 0).value)
    while not "1" in elementValue:
        currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
    firstElementPosition = currentPosition + 2
    savedFirstElementPosition = firstElementPosition

    # Поиск последнего элемента
    currentPosition = firstElementPosition; cell = 0; count = 0
    while True:
        for element in range(10):
            value = str(worksheet.cell(currentPosition, cell).value)
            if value.strip():
                count += 1;
            cell += 1
        if count == 0:
            break
        currentPosition += 1; cell = 0; count = 0
    lastElementPosition = currentPosition
    savedLastElementPosition = lastElementPosition

    # Количество элементов
    total = lastElementPosition - firstElementPosition
    savedTotal = total

    # Текущая позиция
    currentPosition = firstElementPosition

    # Обработка строк в соответствии с документацией ТФФ ГМУ
    for line in range(total):
        # Если элемент находится в ячейке 0, то он - reportItem,
        # а если > 0 - reportSubItem, относящийся к reportItem
        cell = 0
        currentElementValue = str(worksheet.cell(currentPosition, cell).value)

        if str(worksheet.cell(currentPosition, 3).value)[-1] == "0":
            # reportItem
            reportItem = doc.createElement('reportItem')

            # name
            name = doc.createElement('name')
            name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 1).value)))
            reportItem.appendChild(name)

            # lineCode
            lineCode = doc.createElement('lineCode')
            lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 3).value)))
            reportItem.appendChild(lineCode)

            # targetFundsStartYear
            targetFundsStartYear = doc.createElement('targetFundsStartYear')
            targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 4).value).replace('0,00', '0').replace('X', '0')))
            reportItem.appendChild(targetFundsStartYear)

            # targetFundsEndYear
            targetFundsEndYear = doc.createElement('targetFundsEndYear')
            targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0').replace('X', '0')))
            reportItem.appendChild(targetFundsEndYear)

            # stateTaskFundsStartYear
            stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
            stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value).replace('0,00', '0').replace('X', '0')))
            reportItem.appendChild(stateTaskFundsStartYear)

            # stateTaskFundsEndYear
            stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
            stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0').replace('X', '0')))
            reportItem.appendChild(stateTaskFundsEndYear)

            # revenueFundsStartYear
            revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
            revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0')))
            reportItem.appendChild(revenueFundsStartYear)

            # revenueFundsEndYear
            revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
            revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0')))
            reportItem.appendChild(revenueFundsEndYear)

            # totalStartYear
            totalStartYear = doc.createElement('totalStartYear')
            totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0')))
            reportItem.appendChild(totalStartYear)

            # totalEndYear
            totalEndYear = doc.createElement('totalEndYear')
            totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0')))
            reportItem.appendChild(totalEndYear)

            # Добавляем к reference
            reference.appendChild(reportItem)
        else:
            # reportSubItem
            reportSubItem = doc.createElement('reportSubItem')

            # name
            name = doc.createElement('name'); currentCell = 1

            # Ищем название, попутно исключая пустые ячейки
            while not str(worksheet.cell(currentPosition, currentCell).value).strip():
                currentCell += 1
            try:
                name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value).splitlines()[1]))
            except IndexError:
                name.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, currentCell).value)))
            reportSubItem.appendChild(name)

            # lineCode
            lineCode = doc.createElement('lineCode')
            lineCode.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 3).value)))
            reportSubItem.appendChild(lineCode)

            # targetFundsStartYear
            targetFundsStartYear = doc.createElement('targetFundsStartYear')
            targetFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 4).value).replace('0,00', '0').replace('X', '0')))
            reportSubItem.appendChild(targetFundsStartYear)

            # targetFundsEndYear
            targetFundsEndYear = doc.createElement('targetFundsEndYear')
            targetFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 8).value).replace('0,00', '0').replace('X', '0')))
            reportSubItem.appendChild(targetFundsEndYear)

            # stateTaskFundsStartYear
            stateTaskFundsStartYear = doc.createElement('stateTaskFundsStartYear')
            stateTaskFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 5).value).replace('0,00', '0').replace('X', '0')))
            reportSubItem.appendChild(stateTaskFundsStartYear)

            # stateTaskFundsEndYear
            stateTaskFundsEndYear = doc.createElement('stateTaskFundsEndYear')
            stateTaskFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 9).value).replace('0,00', '0').replace('X', '0')))
            reportSubItem.appendChild(stateTaskFundsEndYear)

            # revenueFundsStartYear
            revenueFundsStartYear = doc.createElement('revenueFundsStartYear')
            revenueFundsStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 6).value).replace('0,00', '0')))
            reportSubItem.appendChild(revenueFundsStartYear)

            # revenueFundsEndYear
            revenueFundsEndYear = doc.createElement('revenueFundsEndYear')
            revenueFundsEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 10).value).replace('0,00', '0')))
            reportSubItem.appendChild(revenueFundsEndYear)

            # totalStartYear
            totalStartYear = doc.createElement('totalStartYear')
            totalStartYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 7).value).replace('0,00', '0')))
            reportSubItem.appendChild(totalStartYear)

            # totalEndYear
            totalEndYear = doc.createElement('totalEndYear')
            totalEndYear.appendChild(doc.createTextNode(str(worksheet.cell(currentPosition, 11).value).replace('0,00', '0')))
            reportSubItem.appendChild(totalEndYear)

            # Добавляем к reportItem
            reportItem.appendChild(reportSubItem)

        # Переходим к следующей строке
        currentPosition += 1

xml_str = doc.toprettyxml(indent="    ")
with open(sys.argv[2], "w") as f:
    f.write(xml_str)

