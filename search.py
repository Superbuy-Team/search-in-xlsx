#!/usr/bin/python

import os, sys
import xlrd

def find(path, word):

    l = []

    d = os.listdir(path)

    for file in d:

        filename = str(path) +'/' + str(file)

        print ('Finding in %s' %file)

        if filename.endswith('.xlsx'):

            wb = xlrd.open_workbook(filename)
            ws = wb.sheet_by_index(0)

            for i, row in enumerate(range(ws.nrows)):
                for j, col in enumerate(range(ws.ncols)):
                    if str(word) in str(ws.cell_value(i, j)):
                        l.append((file,row,col))
    if l:
        print ('Word %s found %d times in:' %(word,len(l)))

        for fn, row, col in l:
            print ('File: %s, row: %s ,column: %s' %(fn,row,col))        
    else:
        print ('Word %s not found' %word) 


if __name__ == "__main__":
    try:
        find (sys.argv[1],sys.argv[2])
    except IndexError:
        print('\tExecute: python searchpy <path> <word>')
        print('\tEg: python searchpy /home/user/files/ Fox')