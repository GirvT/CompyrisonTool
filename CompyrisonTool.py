# -*- coding: utf-8 -*-
"""
Created on Mon Mar 12 10:01:01 2019

@author: Girvan Tse
I have regrets tbh
"""
import re
from PySimpleGUI import Text, FileBrowse, FolderBrowse, Input, OK, Cancel, Window, Popup, Button, Column
from CompySame import compareAll
from pandas import ExcelWriter, DataFrame, read_excel, concat
from xlrd import XLRDError

layout1 = [[Text('How many files would you like to compare?')],
           [Input('#', key='numFiles')],
           [Button('Next',key='toLayout2'), Button('Cancel', key='exit')]]

layout2 = []

layout3 = [[Text('Where would you like the export to be put')],
           [Input('Path'), FolderBrowse(key = 'exportPath')],
           [Button('Finish',key='export'), Button('Cancel', key='exit')]]

def validate(file):
    try:
        testParam = read_excel(file[0], 
                               sheet_name = file[1])
        testParam[file[2]]
    except FileNotFoundError:
        return 1
    except XLRDError:
        return 1
    except KeyError:
        return 1
    return 0

def Same(dataINframes, dataINlist):
    dataOUTframes = list()
    dataOUTlist = list()
    for indx in range(0, len(dataINframes)):
        dataOUTlist.append(list())
    for match in compareAll(*dataINlist):
        for indx in range(0, len(match)):
            dataOUTlist[indx].append(dataINframes[indx].values.tolist()[match[indx]])
    for indx in range(0, len(dataINframes)):
        dataOUTframes.append(DataFrame(dataOUTlist[indx], columns = list(dataINframes[indx])))
    for indx in range(0, len(dataOUTframes)):
        dataOUTframes[indx] = dataOUTframes[indx].drop_duplicates()
    return dataOUTframes

def Different(dataINframes, sameEntries):
    dataOUTframes = list()
    for indx in range(0, len(dataINframes)):
        dataOUTframes.append((dataINframes[indx].merge(sameEntries[indx],indicator = True, how='left').loc[lambda x : x['_merge']!='both']).drop_duplicates())
    return dataOUTframes

window = Window('CompyrisonTool').Layout(layout1)
runTool = True
while runTool:
    event, values = window.Read()
    if (event is None or
        event == 'exit'):
        runTool = False

    if (event == 'toLayout2'):
        try:
            values['numFiles'] = int(values['numFiles'])
            layout2.append([Text('Comparing ' + str(values['numFiles']) + ' files')])
            column = []
            for i in range(0, values['numFiles']):
                column.append([Text(i + 1), Input('Path', size=(49,0)),
                               FileBrowse(file_types=(("Excel Workbook", "*.xlsx"),
                                                      ("All Files", "*.*")))])
                column.append([Input('Sheet Name', size=(25,0)),
                            Input('Indicator Name', size=(25,0))])
            layout2.append([Column(column, scrollable=True, vertical_scroll_only=True, size=(450, 400))])
            layout2.append([Button('Next',key='toLayout3'), Button('Cancel', key='exit')])
            window.Close()
            window = Window('CompyrisonTool').Layout(layout2)
        except:
            Popup('Please enter a interger value')

    if (event == 'toLayout3'):
        fileList = list()
        fileItem = list()
        for i in range(0, len(values)):
            fileItem.append(values[i])
            if len(fileItem) == 3:
                fileList.append(fileItem)
                fileItem = []
        
        dataINframes = list()
        dataINlist = list()
        try:
            for arg in fileList:
                    dataINframes.append(read_excel(arg[0], sheet_name = arg[1]))
                    dataINlist.append(dataINframes[-1][arg[2]].tolist())
        except:
                Popup('Your path, sheet, or indicator is incorrect')
        toExport = list()
        sameResult = Same(dataINframes, dataINlist)
        diffResult = Different(dataINframes, sameResult)
        for df in sameResult:
            df.reset_index(drop = True, inplace = True)
        for df in diffResult:
            df.reset_index(drop = True, inplace = True)
        toExport.append(sameResult)
        toExport.append(diffResult)
        window.Close()
        window = Window('CompyrisonTool').Layout(layout3)

    if (event == 'export'):
        writer = ExcelWriter(values['exportPath'] + "/OUTPUT.xlsx",
                             engine = 'xlsxwriter')
        for same in range(0, len(toExport[0])):
            header = fileList[same][0].rsplit('/', 1)[-1][:-5] + fileList[same][1] + fileList[same][2]
            if len(header) > 27:
                header = header[(len(header) - 27):]
            sameHeader = header + 'SAME'
            toExport[0][same].to_excel(writer, sheet_name = sameHeader)
        for diff in range(0, len(toExport[1])):
            header = fileList[diff][0].rsplit('/', 1)[-1][:-5] + fileList[diff][1] + fileList[diff][2]
            if len(header) > 27:
                header = header[(len(header) - 27):]
            diffHeader = header + 'DIFF'
            toExport[1][diff].to_excel(writer, sheet_name = diffHeader)
        try:
            writer.save()
            writer.close()
            Popup("Success!",
                ("Successfully completed operation, see " +
                values['exportPath'] +
                "/OUTPUT.xlsx for output"))
        except:
            Popup('Unable to save, incorrect path or you have the file open')
        runTool = False
window.Close()