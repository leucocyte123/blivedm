import time
import datetime
import openpyxl
from openpyxl import workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet import worksheet

from blivedm import GuardBuyMessage

class WorkbookMetadata:
    def __init__(self, filename: str, sheetname: str, head: list):
        self.filename = filename
        self.sheetname = sheetname
        self.head = head

guard_info = WorkbookMetadata(
    '大航海记录.xlsx', 
    '大航海记录', 
    ['弹幕ID', 'B站昵称', 'UID', '日期', '时间', '等级', '数量']
)

guard_bonus = WorkbookMetadata(
    '大航海福利.xlsx',
    '大航海福利',
    ['B站昵称', 'UID', '日期', '礼物', '消耗次数']
)

summary = WorkbookMetadata(
    '大航海信息汇总.xlsx',
    '大航海信息汇总',
    ['B站昵称', 'UID', '累计舰长次数', '已兑换', '剩余次数']
)


def checkFileExist(filename):
    try:
        workbook = openpyxl.load_workbook(filename)
    except:
        workbook = openpyxl.Workbook()
        workbook.save(filename)
    return workbook

# def checkSheetExist(filename, sheetname):
#     workbook = checkFileExist(filename)
#     try:
#         worksheet = workbook[sheetname]
#     except:
#         worksheet = workbook.create_sheet(sheetname)
#         if filename == filename_guard_info and sheetname == sheetname_guard_info:
#             for index, item in enumerate():
#                 worksheet.cell(1, index + 1, item)
#         elif filename == filename_guard_bonus and sheetname == sheetname_guard_bonus:
            
#         elif filename == filename_summary and sheetname == sheetname_summary:
#             pass
#         else:
#             raise 'Undefined filename or sheetname'
#     return worksheet

def updateGuardInfo(message: GuardBuyMessage):
    # Save raw information to the file

    # Create or update the summary information
    
    pass