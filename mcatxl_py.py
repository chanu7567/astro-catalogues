import xlrd
from datetime import time
import os

mcatxlFile = "mcatxl.xls"
basedir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(basedir, mcatxlFile)

continueProgram = ''
mNumber = 0

class RangeError(Exception):
    pass

# class NullUserInputError(Exception):
#     pass

workBook = xlrd.open_workbook(file_path)

def getObjectType(type):
    global workBook
    sheet = workBook.sheet_by_index(1)

while continueProgram != 'q' or continueProgram != 'Q':
        
    sheet = workBook.sheet_by_index(0)

    try:
        mNumber = int(input("Enter the Messier Number : "))
        # if mNumber is None:
        #     raise NullUserInputError("Enter a Messier Number!")
        if mNumber <= 0 or mNumber > 110:
            raise RangeError("The Messier Number should be in between 0 and 111")
    except ValueError:
        print("Invalid Input!!!")
    except RangeError:
        print("The Messier Number should be in between 0 and 111")
    # except NullUserInputError:
    #     print("Enter a Messier Number!")
    else:
        columnNames = sheet.row_values(0)
        dataRow = sheet.row_values(mNumber)
        print("--------------------------------------------------------------")
        for columnNumber in range(sheet.ncols):
            if dataRow[columnNumber] == '':
                dataRow[columnNumber] = 'N/A'
            if columnNames[columnNumber] == 'RA':
                timeInt = int(dataRow[columnNumber] * 24 * 3600)
                dataRow[columnNumber] = time(timeInt//3600, (timeInt%3600)//60, timeInt%60)
            if columnNames[columnNumber] == 'M_Num' or columnNames[columnNumber] == 'NGC_Num':
                dataRow[columnNumber] = dataRow[columnNumber] // 1
                if columnNames[columnNumber] == 'M_Num':
                    columnNames[columnNumber] = 'Messier Number'
                if columnNames[columnNumber] == 'NGC_Num':
                    columnNames[columnNumber] = 'NGC Number'
            print(columnNames[columnNumber], ':\t', dataRow[columnNumber])
        print("--------------------------------------------------------------")
        #print(*columnNames, sep='\n')
        #print(*dataRow, sep='\n')
        #input("\nPress return to continue")

    print('\nEnter q to exit program. Press any other key to continue')
    continueProgram = input()
    if continueProgram == 'q' or continueProgram == 'Q':
        break;

print("--------------------------------------------------------------")
print("--------------------------------------------------------------")
input("Program Terminated - Press Return to Exit Window")
