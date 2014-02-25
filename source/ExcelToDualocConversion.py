#!/usr/bin/python
######################################################################
#
#   file: ExcelToDualocConversion.py
#   author: Ken Zyma
#
#   Function class to convert excel format spec'd in readme.txt to
#       DUALOC input.
#
#   modification history:
#       Feb. 2014- created (author: Ken Zyma)
#
######################################################################

import xlrd
import ntpath
import os

'''
Return Constants
0: excecuted normally, file converted
-1:error opening file
1: error opening/creaing output file
3: excel sheet format is incorrect
'''

class ExcelToDualocConversion():

    def __call__(self,filePath,sheetIndex=0):
        #get data from excel workbook
        excelWorkbook = self.openExcelWorkbook(filePath)
        
        #if excelWorkbook fails, return -1 and allow client to
        #try a new file.
        if(excelWorkbook == -1):
            return -1
        
        excelDataSet = excelWorkbook.sheet_by_index(sheetIndex)

        try:
            #get column and row data from excel sheet
            numOfColumns=int(excelDataSet.cell(0,0).value)
            numOfRows=int(excelDataSet.cell(0,1).value)
            #get paramter data from excel sheet
            nrhs=int(excelDataSet.cell(0,2).value)
            jscn=int(excelDataSet.cell(0,3).value)
            iout=int(excelDataSet.cell(0,4).value)
            jdcyc=int(excelDataSet.cell(0,5).value)
            depth=int(excelDataSet.cell(0,6).value)
            maxIterations=int(excelDataSet.cell(0,7).value)

            #fixed cost data array
            a_fixedCost=[]
            for i in range(1,(numOfColumns+1)):
                a_fixedCost.append(int(excelDataSet.cell(1,i).value))

            #note that python dic is unordered, so you must
            #order the data below(least->greatest) before transfering
            #it to DUALOC format
            a_colCoverData=[]
            for row in range(3,numOfRows+3):
                d_colCoverDataPerRow={}
                for col in range(1,(numOfColumns+1)):
                    if(excelDataSet.cell(row,col).value!=''):
                        d_colCoverDataPerRow.update({col:(int(excelDataSet.cell(row,col).value))})
                #ordered list of dicts
                a_colCoverData.append(d_colCoverDataPerRow)

        except:
            print "Incorrect excel file format, see readme.txt"
            return 3

        #output data to .dat file
        #open will create file if one does not exist (with w+)
        #mirrorInFileName= ntpath.basename(filePath)
        #get rid of excel extension
        mirrorInFileName=os.path.splitext(filePath)[0]

        try:
            outputFile = open(mirrorInFileName+"-"+excelDataSet.name+'.dat', 'w+')
        except IOError:
            return 1
        
        #line 1
        outputFile.write('{}'.format(numOfColumns))
        outputFile.write('   ')
        outputFile.write('{}'.format(numOfRows))
        outputFile.write('    ')
        outputFile.write('{}'.format(nrhs))
        outputFile.write('     ')
        outputFile.write('{}'.format(jscn))
        outputFile.write('     ')
        outputFile.write('{}'.format(iout))
        outputFile.write('     ')
        outputFile.write('{}'.format(jdcyc))
        outputFile.write('   ')
        outputFile.write('{}'.format(depth))
        outputFile.write(' ')
        outputFile.write('{}'.format(maxIterations))
        outputFile.write('\n')
        #all column/row data
        for i in range(numOfRows):
            outputFile.write('    ')
            outputFile.write('{}'.format(i+1))
            outputFile.write('      ')
            outputFile.write('{}'.format(len(a_colCoverData[i])))
            outputFile.write('\n')
            #a_rowValuesSorted=sorted(a_colCoverData[i].values())
            #put all dictionary key/value pairs into a list
            a_rowValuesSorted = a_colCoverData[i].items()
            a_rowValuesSorted = sorted(a_rowValuesSorted, key=lambda rowCover:rowCover[1])
            outputFile.write('  ')
            for j in a_rowValuesSorted:
                #next line should be safe, python docs garentee order will not change as long as
                #the dict is not being modified as its going on (multi-threaded)
                #outputFile.write('{}'.format(a_colCoverData[i].keys()[a_colCoverData[i].values().index(j)]))
                outputFile.write('{}'.format(j[0]))
                outputFile.write('  ')
                outputFile.write('{}'.format(j[1]))
                outputFile.write('   ')
            outputFile.write('\n')
        outputFile.write(' ')
        #fixed cost row
        for i in range(numOfColumns):
            outputFile.write('{}'.format(a_fixedCost[i]))
            outputFile.write('    ')

        outputFile.close()
        return 0

    def openExcelWorkbook(self,filePath):
        try:
            #do not need to explicity close file, xlrd handles this.
            f=xlrd.open_workbook(filePath)
        except IOError:
            print 'file',filePath,'failed to open'
            f=-1
        except xlrd.XLRDError:
            print 'file',filePath,'is an unsupported format, see xlrd documentation'+\
                  'https://github.com/python-excel/xlrd'
            f=-1
        finally:
            return f






