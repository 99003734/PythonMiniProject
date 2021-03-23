import xlrd
import xlsxwriter
import openpyxl as op
import os.path
from os import path
import os


class Read:
    def workbookFunc(self, Name, PsNo, EmailId):
        global fileName
        res = []
        locs = ['sample1.xlsx', 'sample2.xlsx', 'sample3.xlsx',
                'sample4.xlsx', 'sample5.xlsx']
        counter = 0
        masterName = "masterSheet.xlsx"
        FileStatus = path.exists(masterName)
        print(FileStatus)
        for loc in locs:
            for root, dir, files in os.walk("/"):
                if loc in files:
                    fileName = os.path.join(root, loc)
            counter += 1
            wb = xlrd.open_workbook(fileName)
            sheet = wb.sheet_by_index(0)
            for i in range(sheet.nrows):
                if (
                    sheet.row_values(i)[0] == Name and
                        sheet.row_values(i)[1] == PsNo and
                        sheet.row_values(i)[2] == EmailId):
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
        if FileStatus is True:
            print('Master Sheet is Present')
            wb1 = op.load_workbook(masterName)
            ws = (wb1['Sheet1'])
            ws.append(res)
            wb1.save(masterName)
            wb1.close()
        else:
            print('Creating master sheet')
            workbook = xlsxwriter.Workbook(masterName)
            worksheet = workbook.add_worksheet()
            header = ['Name', 'PsNo', 'EmailId', 'Company Name', 'Location',
                      'Position', 'Business Unit', 'Floor', 'Joining Date',
                      'Phone Number', 'College Name', 'Location', 'Degree',
                      'Branch', 'CGPA', 'Grade', 'Passing Year',
                      'Higher Secondary(12)', 'College Name', 'Maths',
                      'Physics', 'Chemistry', 'Percentage', 'Passing Year',
                      'Secondary(10)', 'School', 'Location', 'Maths',
                      'Physics', 'Chemistry', 'CGPA', 'Gender',
                      'Marital Status', 'Aadhar Number',
                      'Pan Number', 'Place Of Birth', 'Pincode', 'Age']
            row = 0
            col = 0
            for coloumn in (header):
                worksheet.write(row, col, coloumn)
                col += 1
            datarow = 0
            datacol = 0
            for data in (res):
                worksheet.write(datarow + 1, datacol, data)
                datacol += 1
            workbook.close()
p1 = Read()
inputs = input("How many inputs: ")
for i in range(int(inputs)):
    name = input("Enter Name: ")
    PSNumber = eval(input("Enter PS Number: "))
    emailId = input("Enter emailId: ")
    p1.workbookFunc(name, PSNumber, emailId)
