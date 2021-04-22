#!/usr/bin/env python3
#!print('Hello World!')

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Color
from copy import copy

inputJapanese = openpyxl.load_workbook('input_Farm_Japanese.xlsx')
inputEnglish = openpyxl.load_workbook('input_Farm_English.xlsx')
#rosettaStone = openpyxl.load_workbook('translation_Pairs_Unicode.xlsx')
outputFile = Workbook()

allJapaneseInputSheetNames = inputJapanese.sheetnames
allEnglishInputSheetNames = inputEnglish.sheetnames
#allOutputSheetNames = outputFile.sheetnames

# japanese = []
# english = []

# rosettaSheet = rosettaStone.active
outputSheet = outputFile.active
# outputSheet.title = str(inputFile.sheetnames[0])


# x = 1
# while len(allJapaneseInputSheetNames) > x:
#     allEnglishInputSheetNames[x] = allJapaneseInputSheetNames[x]
#     x += 1


# def rosettaData ():
#     i = 1
#     i2 = 1
#     while rosettaSheet.max_row >= i:
#         japanese.append(rosettaSheet['A'+ str(i)].value)
#         i += 1
#
#     while rosettaSheet.max_row >= i2:
#         english.append(rosettaSheet['B'+ str(i2)].value)
#         i2 += 1
#
# rosettaData()


def scan_Japanese ():
    i = 1
    for sheet in allJapaneseInputSheetNames:
        currentJapaneseSheet = inputJapanese[sheet]

        for sheets in inputJapanese:
            for row in range(1, currentJapaneseSheet.max_row + 1):
                for column in range(1, currentJapaneseSheet.max_column + 1):
                    currentJapaneseCell = currentJapaneseSheet.cell(row=row, column=column)

                    if currentJapaneseCell.value is None:
                        continue
                    elif currentJapaneseCell.value == "":
                        continue
                    else:
                        outputSheet.cell(row=i, column=1, value=currentJapaneseCell.value)
                        i += 1

scan_Japanese()

def scan_English ():
    i = 1
    for sheet in allEnglishInputSheetNames:
        currentEnglishSheet = inputEnglish[sheet]

        for sheets in inputEnglish:
            for row in range(1, currentEnglishSheet.max_row + 1):
                for column in range(1, currentEnglishSheet.max_column + 1):
                    currentEnglishCell = currentEnglishSheet.cell(row=row, column=column)

                    if currentEnglishCell.value is None:
                        continue
                    elif currentEnglishCell.value == "":
                        continue
                    else:
                        outputSheet.cell(row=i, column=2, value=currentEnglishCell.value)
                        i += 1

scan_English()
#

#
#
#
#                     # i = 0
#                     # while i < len(japanese):
#                     #     if currentCell.value == japanese[i]:
#                     #         outputSheet.cell(row=row, column=column, value=english[i])
#                     #         currentFont = currentCell.font
#                     #         outputCell.font = copy(currentFont)
#                     #         break
#                     #
#                     #     else:
#                     #         i += 1
#                     #
#                     # if outputCell.value is None:
#                     #     outputSheet.cell(row=row, column=column, value=currentCell.value)
#                     #     currentFont = currentCell.font
#                     #     outputCell.font = copy(currentFont)
#


outputFile.save('output_Farm.xlsx')
