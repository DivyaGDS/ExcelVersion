# Converts a Microsoft Excel 2007+ file into plain text
# for comparison using git diff
#
# Instructions for setup:
# 1. Place this file in a folder
# 2. Add the following line to the global .gitconfig:
#     [diff "zip"]
#   	    binary = True
#	    textconv = python c:/path/to/git_diff_xlsx.py
# 3. Add the following line to the repository's .gitattributes
#    *.xlsx diff=zip
# 4. Now, typing [git diff] at the prompt will produce text versions
# of Excel .xlsx files
#
# Copyright William Usher 2013
# Contact: w.usher@ucl.ac.uk
#

import subprocess
import xlrd as xl
import sys
import unicodedata

def parse(infile,outfile):
    """
    Converts an Excel file into text
    Returns a formatted text file for comparison using git diff.
    """
    # print ("infile print..",infile)
    book = xl.open_workbook(infile)

    num_sheets = book.nsheets
    temp_list = []
    # print("initial temp o..",temp_list[0])
    # print("initial temp 1..",temp_list[1])
    # for x in infile:
    #     #  print("inline..",infile)
    #      temp_list.append(infile)
         

    # print("temp o..",temp_list[0])
    # print("temp 1..",temp_list[1])
    # print("temp length..",len(temp_list))
   

#   print "File last edited by " + book.user_name + "\n"
    outfile.write("File last edited by " + book.user_name + "\n")
    # subprocess.call(['java', '-jar', 'C://Users//UB217ZA//Downloads//excel-version-control//target//excel-version-control-1.0.jar', "ConvertToCSV"])


    def get_cells(sheet, rowx, colx):
        try:
            value = str(sheet.cell_value(rowx, colx))
            # print("value is..",value)
        except:
            value = ''
        if (rowx,colx) in sheet.cell_note_map.keys():
            value += ' <<' + unicodedata.normalize('NFKD', str(sheet.cell_note_map[rowx,colx].text)).encode('ascii', 'ignore').decode('ascii') + '>>'
            # print("final value is",value)
        return value

    # # loop over worksheets

    for index in range(0,num_sheets):
        # find non empty cells
        sheet = book.sheet_by_index(index)
        outfile.write("=================================\n")
        sheetname = unicodedata.normalize('NFKD', str(sheet.name)).encode('ascii', 'ignore').decode('ascii')
        outfile.write("Sheet: " + sheetname + "[ " + str(sheet.nrows) + " , " + str(sheet.ncols) + " ]\n")
        outfile.write("=================================\n")
        for row in range(0,sheet.nrows):
            for col in range(0,sheet.ncols):
                content = get_cells(sheet, row, col)
                if content != "":
                    output = '    ' + xl.cellname(row,col) + ':\n        '
                    output += unicodedata.normalize('NFKD', str(content)).encode('ascii', 'ignore').decode('ascii')
                    output += '\n'
                    outfile.write(output)
        print ("\n")

# output cell address and contents of cell
def main():
    print("In main..")
    args = sys.argv[1:]
    print('args 1..',args)
    
    if len(args) != 1:
        print ('usage: python git_diff_xlsx.py infile.xlsx')
        sys.exit(-1)
    outfile = sys.stdout
    parse(args[0],outfile)

if __name__ == '__main__':
    main()
    