# Author : Shaik Akbar Basha
# PsNo: 99003734
# Email id: shaik.basha1@ltts.com

"""In this program iam using openpyxl library to get the details of a
particular person from 5 excels files with the help of name,ps no,email id and
writing the data in a master sheet, these excel files are present in a different
directory"""

# import libraries

import xlrd
import xlsxwriter
import openpyxl as op
import os.path
from os import path
import os


# Using Classes and functions and passing mane,PsNo,Email id as parameters

class Read:
    def workbookFunc(self, Name, PsNo, EmailId):
        global fileName
        res = []
        locs = ['sample1.xlsx', 'sample2.xlsx', 'sample3.xlsx', 'sample4.xlsx', 'sample5.xlsx']
        counter = 0
        masterName = "masterSheet.xlsx"

# To check the path existance of master sheet

        FileStatus = path.exists(masterName)
        print(FileStatus)
        
#To check the existance of Excel file in a directory

        for loc in locs:
            for root, dir, files in os.walk("/"):
                if loc in files:
                    fileName = os.path.join(root, loc)
            counter += 1
            wb = xlrd.open_workbook(fileName)
            sheet = wb.sheet_by_index(0)
            for i in range(sheet.nrows):
                if (sheet.row_values(i)[0] == Name and sheet.row_values(i)[1] == PsNo and sheet.row_values(i)[
                    2] == EmailId):
                    print(sheet.row_values(i), type(sheet.row_values(i)))
                    if counter == 1:
                        for i in sheet.row_values(i):
                            if i not in res:
                                res.append(i)
                    else:
                        print(sheet.row_values(i)[3:])
                        for i in sheet.row_values(i)[3:]:
                            res.append(i)
        print(res)
        
# Master sheet is present and writing data to the master sheet

        if FileStatus == True:
            print('Master Sheet is Present')
            wb1 = op.load_workbook(masterName)
            ws = (wb1['Sheet1'])
            ws.append(res)
            wb1.save(masterName)
            wb1.close()

# Create a Mastersheet workbook and add a worksheet.

        else:
            print('Creating master sheet')
            workbook = xlsxwriter.Workbook(masterName)
            worksheet = workbook.add_worksheet()
            header = ['Name', 'PsNo', 'EmailId', 'Company Name', 'Location', 'Position', 'Business Unit',
                      'Floor', 'Joining Date', 'Phone Number', 'College Name', 'Location', 'Degree', 'Branch', 'CGPA',
                      'Grade', 'Passing Year', 'Higher Secondary(12)', 'College Name', 'Maths', 'Physics', 'Chemistry',
                      'Percentage', 'Passing Year', 'Secondary(10)', 'School', 'Location', 'Maths', 'Physics',
                      'Chemistry',
                      'CGPA', 'Gender', 'Marital Status', 'Aadhar Number', 'Pan Number', 'Place Of Birth', 'Pincode',
                      'Age']

# Start from the first cell. Rows and columns are zero indexed.
# Adding headers to the master sheet

            row = 0
            col = 0
            for coloumn in (header):
                worksheet.write(row, col, coloumn)
                col += 1
            
# Writing data to the worksheet

            datarow = 0
            datacol = 0
            for data in (res):
                worksheet.write(datarow + 1, datacol, data)
                datacol += 1
            workbook.close()

#Creating 'p1' object

p1 = Read()

#passing name, psno, email id as agruments in a workbook function

inputs = input("How many inputs: ")
for i in range(int(inputs)):
    name = input("Enter Name: ")
    PSNumber = eval(input("Enter PS Number: "))
    emailId = input("Enter emailId: ")
    p1.workbookFunc(name, PSNumber, emailId)
