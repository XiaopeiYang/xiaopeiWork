#!/usr/bin/python3
# -*-coding:utf-8 -*-

# Reference:**********************************************
# @Time    : 6/16/2021 12:20 PM
# @Author  : Gaopeng.Bai
# @File    : excelOperator.py
# @User    : gaope
# @Software: PyCharm
# @Description:
# Reference:**********************************************
# global element
import xlrd
import time


class excelOperation:
    """
    excel operations for the submitting the tickets
    """
    def __init__(self, excelFile, sheetName="Ticket Log") -> object:
        self.workbook = xlrd.open_workbook(excelFile)
        self.ticketSheet = self.workbook.sheet_by_name(sheetName)  # 打开提单页面

    def getMaxRow(self):
        """
        Get the max row number
        """
        print(self.ticketSheet.get_rows())
        return self.ticketSheet.get_rows()

    def OpenTicketExcel(self, start_Row, end_Row):
        """
            read the content of excel file to array list
        :param start_Row:
        :param end_Row:
        :return: array list with excel content
        """
        ticketTable = []  # the container of excel content
        startRow = start_Row - 1
        endRow = end_Row
        for row in range(startRow, endRow):
            Temp = self.ticketSheet.row_values(row)
            Time_Excel = Temp[6] + 1
            Time_Excel = (Time_Excel - 19 - 70 * 365) * \
                86400 - 8 * 3600  # 将excel时间序列值转化为时间戳
            Time_Local = time.localtime(Time_Excel)  # 将时间戳转化为localtime
            Temp[6] = time.strftime(
                "%Y-%m-%d", Time_Local)  # 将localtime转换成提单日期
            ticketTable.append(Temp)  # 时间格式改好，记录入二维数据表
        self.workbook.release_resources()
        return ticketTable
