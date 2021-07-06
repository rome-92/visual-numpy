# Copyright Román U. Martínez
# Distributed under the terms of the GNU General Public License

# --------------------------------------------------------------------
#    This file is part of Visual Numpy.
#
#    Visual Numpy is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    Visual Numpy is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with Visual Numpy.  If not, see <https://www.gnu.org/licenses/>.
# --------------------------------------------------------------------

from PySide6.QtCore import (QAbstractTableModel, QMimeData, QByteArray, QDataStream, QIODevice,
                            Qt,QItemSelectionModel, QModelIndex)
import globals_
import traceback
import numpy as np


class MyModel(QAbstractTableModel):
    def __init__(self,parent=None):
        '''Initialize the model'''
        super().__init__(parent)
        #creates a data type consisting of a unicode string of lenght 32 and a 128 bit complex number
        dataType = np.dtype('U32,D')
        #the array that will store the data is a 2 dimensional array
        self.dataContainer = np.zeros((52,52),dtype=dataType)
        self.formulas = []
        self.ftoapply = []
        self.formulaSnap = []
        self.highlight = None
        self.domainHighlight = {}
        self.alignmentDict = {}
        self.fonts = {}
        self.background = {}
        self.foreground = {}
        self.history = []
        self.history.append((self.dataContainer.copy(),self.formulas.copy(),self.alignmentDict.copy(),
            self.fonts.copy(),self.foreground.copy(),self.background.copy()))
        self.thousandsSep = True

    def enableThousandsSep(self):
        '''Enables thousands separator'''
        self.thousandsSep = True

    def disableThousandsSep(self):
        '''Disables thousands separator'''
        self.thousandsSep = False

    def mimeTypes(self):
        '''Intended to provide drop of csv data and app specific data'''
        return ['text/csv','application/octet-stream']

    def mimeData(self,indexes,flag='mouse'):
        '''Creates the MimeData object app specific'''
        mimeData = QMimeData()
        encodedData = QByteArray()
        stream = QDataStream(encodedData,QIODevice.WriteOnly)
        rows = []
        columns = []
        textMatrix = ''
        formulas = []
        for index in indexes:
            for f in self.formulas:
                if f.addressRow == index.row() and f.addressColumn == index.column():
                    formulas.append(f)
                    break
            rows.append(index.row())
            columns.append(index.column())
        clip = self.dataContainer[min(rows):max(rows)+1,min(columns):max(columns)+1]['f0']
        for arrayRow in clip:
            textMatrix += ','.join(arrayRow)
            textMatrix += '\n'
        textMatrix = textMatrix[:-1]
        stream.writeQString(textMatrix)
        if flag == 'mouse':
            currentIndexRow = str(self.parent().currentIndex().row())
            currentIndexColumn = str(self.parent().currentIndex().column())
        topLeftIndexRow = str(min(rows))
        topLeftIndexColumn = str(min(columns))
        if flag == 'keepTopIndex':
            currentIndexRow = topLeftIndexRow
            currentIndexColumn = topLeftIndexColumn
        stream.writeString(currentIndexRow)
        stream.writeString(currentIndexColumn)
        stream.writeString(topLeftIndexRow)
        stream.writeString(topLeftIndexColumn)
        for f in formulas:
            stream.writeString(str(f.addressRow)+','+str(f.addressColumn))
        mimeData.setData('application/octet-stream',encodedData)
        return mimeData

    def dropMimeData(self,data,action,row,column,parent):
        '''Handles drop action from app specific mime type'''
        if data.hasFormat('application/octet-stream'):
            encodedData = data.data('application/octet-stream')
            stream = QDataStream(encodedData,QIODevice.ReadOnly)
            data = []
            matrix = []
            while not stream.atEnd():
                data.append(stream.readString())
            textMatrix,rowFromData,columnFromData,topLeftRow,topLeftColumn = data[:5]
            rowFromData = int(rowFromData)
            columnFromData = int(columnFromData)
            topLeftRow = int(topLeftRow)
            topLeftColumn = int(topLeftColumn)
            textMatrix = textMatrix.splitlines()
            for textMatrixRow in textMatrix:
                matrix.append(textMatrixRow.split(sep=','))
            dropArray = np.array(matrix)
            dropRow = parent.row()
            dropColumn = parent.column()
            newRowDifference = dropRow - rowFromData
            newColumnDifference = dropColumn - columnFromData
            newRow = topLeftRow + newRowDifference
            newColumn = topLeftColumn + newColumnDifference
            dropArrayShape = dropArray.shape
            dropArrayRows = dropArrayShape[0]
            dropArrayColumns = dropArrayShape[1]
            self.formulaSnap = self.formulas.copy()
            formulasAddresses = set()
            if action == Qt.MoveAction:
                if len(data) > 5:
                    formulas2move = data[5:]
                    for f in formulas2move:
                        address = eval(f)
                        row = address[0]
                        column = address[1]
                        for modelFormula in self.formulas:
                            if modelFormula.addressRow == row and modelFormula.addressColumn == column:
                                try:
                                    modelFormula.addressRow = row + newRowDifference
                                    modelFormula.addressColumn = column + newColumnDifference
                                    self.checkForCircularRef(modelFormula,newRowDifference,newColumnDifference)
                                    formulasAddresses.add((modelFormula.addressRow,modelFormula.addressColumn))
                                except Exception as e:
                                    traceback.print_tb(e.__traceback__)
                                    print(e)
                                    modelFormula.addressRow = row
                                    modelFormula.addressColumn = column
                for row in range(topLeftRow,topLeftRow+dropArrayRows):
                    for column in range(topLeftColumn,topLeftColumn+dropArrayColumns):
                        self.setData(self.index(row,column),'',formulaTriggered=True)
            for row,y in zip(dropArray,range(newRow,newRow+dropArrayRows)):
                for column,x in zip(row,range(newColumn,newColumn+dropArrayColumns)):
                    if (y,x) in formulasAddresses:
                        self.setData(self.index(y,x),column,formulaTriggered=True)
                    else:
                        self.setData(self.index(y,x),column,formulaTriggered='ERASE')
            selectionModel = self.parent().selectionModel()
            selectionModel.clear()
            for row_z in  range(newRow,newRow+dropArrayRows):
                for column_z in range(newColumn,newColumn+dropArrayColumns):
                    selectionModel.select(self.index(row_z,column_z),QItemSelectionModel.Select)
            if self.ftoapply:
                self.updateModel_()
            return True

    def checkForCircularRef(self,formula,*deltas):
        '''Checks if there's a circular reference if formula is to be dropped'''
        rowsDomain = []
        columnsDomain = []
        newRowDifference = deltas[0]
        newColumnDifference = deltas[1]
        for index in formula.domain:
            rowsDomain.append(index[0])
            columnsDomain.append(index[1])
        topLeftDomain = [min(rowsDomain),min(columnsDomain)]
        bottomRightDomain = [max(rowsDomain),max(columnsDomain)]
        topLeftDomain[0] = topLeftDomain[0]+newRowDifference
        topLeftDomain[1] = topLeftDomain[1]+newColumnDifference
        bottomRightDomain[0] = bottomRightDomain[0]+newRowDifference
        bottomRightDomain[1] = bottomRightDomain[1]+newColumnDifference
        oldDomain = formula.domain
        newDomain = []
        for row in range(topLeftDomain[0],bottomRightDomain[0]+1):
            for column in range(topLeftDomain[1],bottomRightDomain[1]+1):
                newDomain.append((row,column))
        formulaIndexesSet = set(formula.indexes)
        newDomainSet = set(newDomain)
        if formulaIndexesSet.intersection(newDomainSet):
            raise CircularReferenceError(formula.addressRow,formula.addressColumn)
        formula.domain = newDomain

    def updateModel_(self):
        '''Executes formulas if any are pending of updating'''
        for f in self.ftoapply:
            if f in self.formulas:
                self.parent().parent().calculate(f.text,f.addressRow,f.addressColumn,flag=True)
        self.ftoapply.clear()

    def columnCount(self,parent=QModelIndex()):
        return self.dataContainer.shape[1]

    def rowCount(self,parent=QModelIndex()):
        return self.dataContainer.shape[0]

    def insertRows(self,row,count,parent=QModelIndex()):
        '''Inserts one row'''
        self.dataContainer.resize((self.rowCount()+1,self.columnCount()))
        self.beginInsertRows(parent,row,row)
        self.endInsertRows()
        return True

    def insertColumns(self,column,count,parent=QModelIndex()):
        '''Inserts one column'''
        self.dataContainer.resize((self.rowCount(),self.columnCount()+1))
        self.beginInsertColumns(parent,column,column)
        self.endInsertColumns()
        return True

    def getAlphanumeric(self,column,row):
        '''Gets the alphanumeric coordinate for the corresponding cell'''
        if column < 26:
            letter = globals_.ALPHABET[column]
            return letter+str(row+1)
        elif column < 702:
            letter1 = globals_.ALPHABET[column//26-1]
            letter2 = globals_.ALPHABET[column-(column//26*26)]
            return letter1+letter2+str(row+1)
        else:
            letter1 = globals_.ALPHABET[(column-26)//676-1]
            letter2 = globals_.ALPHABET[(column-((column-26)//676*676))//26-1]
            letter3 = globals_.ALPHABET[column-(676+(column-676)//26*26)]
            return letter1+letter2+letter3+str(row+1)

    def data(self,index,role=Qt.DisplayRole):
        '''Returns the appropiate data for the corresponding role'''
        if role == Qt.DisplayRole:
            returnValue = self.dataContainer\
                    [index.row(),index.column()]
            try:
                #checks if the value is a number
                complex(returnValue['f0']) 
            except Exception:
                #if it's not a number returns the corresponding string
                return returnValue['f0']  
            else:
                #if it's a number it first checks if it has an imaginary part
                if returnValue['f1'].imag == 0: 
                    #if not it returns only the real part in string form
                    if returnValue['f1'].real.is_integer():
                        if self.thousandsSep:
                            return '{:,d}'.format(int(returnValue['f1'].real))
                        else:
                            return '{:d}'.format(int(returnValue['f1'].real))
                    else:
                        if self.thousandsSep:
                            return '{:,.8f}'.format(returnValue['f1'].real)
                        else:
                            return '{:.8f}'.format(returnValue['f1'].real)
                else:
                    #if it does have an imaginary part it returns the real and im part without parenteses
                    return returnValue['f0'].strip('()')
        if role == Qt.BackgroundRole:
            if self.highlight:
                if index.row() == self.highlight[0][0]:
                    if index.column() == self.highlight[0][1]:
                        return self.highlight[1]
            elif self.domainHighlight:
                if (index.row(),index.column()) in self.domainHighlight:
                    return self.domainHighlight[index.row(),index.column()]
            else:
                return self.background.get((index.row(),index.column()),globals_.defaultBackground)
        if role == Qt.FontRole:
            return self.fonts.get((index.row(),index.column()),globals_.defaultFont)
        if role == Qt.TextAlignmentRole:
            return self.alignmentDict.get((index.row(),index.column()),int(Qt.AlignLeft|Qt.AlignVCenter))
        if role == Qt.ForegroundRole:
            return self.foreground.get((index.row(),index.column()),globals_.defaultForeground)

    def setData(self,index,value,role=Qt.EditRole,formulaTriggered=False):
        '''Sets the appropiate data for the corresponding role'''
        if index.row() > self.dataContainer.shape[0]-1:
            self.insertRow(self.rowCount())
        if index.column() > self.dataContainer.shape[1]-1:
            self.insertColumn(self.columnCount())
        if role == Qt.EditRole:
            if str(value) == self.dataContainer[index.row(),index.column()]['f0']: return True
            try:
                #checks if the value is a number
                endValue = complex(value) 
                #if it's a number the corresponding assignments are made
                self.dataContainer[index.row(),index.column()]['f0'] = value 
                self.dataContainer[index.row(),index.column()]['f1'] = endValue
            except Exception:
                if value != None:
                    #if it's NaN then the number field 'f1' is set to zero
                    self.dataContainer[index.row(),index.column()]['f0'] = value 
                    self.dataContainer[index.row(),index.column()]['f1'] = 0
                else:
                    self.dataContainer[index.row(),index.column()]['f0'] = ''
                    self.dataContainer[index.row(),index.column()]['f1'] = 0
            try:
                assert self.formulas
            except AssertionError:
                return True
            if formulaTriggered == 'ERASE':
                for f in self.formulas:
                    if index.row() == f.addressRow and index.column() == f.addressColumn:
                        self.formulas.remove(f)
                        break
                for f in self.formulaSnap.copy():
                    for i in f.indexes:
                        if index.row() == i[0] and index.column() == i[1]:
                            self.ftoapply.append(f)
                            self.formulaSnap.remove(f)
                            break
            elif formulaTriggered == True:
                for f in self.formulaSnap.copy():
                    for i in f.indexes:
                        if index.row() == i[0] and index.column() == i[1]:
                            self.ftoapply.append(f)
                            self.formulaSnap.remove(f)
                            break
            elif formulaTriggered == False:
                for f in self.formulas:
                    if index.row() == f.addressRow and index.column() == f.addressColumn:
                        self.formulas.remove(f)
                        break
                for f in self.formulas:
                    for i in f.indexes:
                        if index.row() == i[0] and index.column() == i[1]:
                            self.ftoapply.append(f)
                            break
                self.updateModel_()
            return True
        elif role == Qt.BackgroundRole:
            if globals_.formula_mode:
                self.highlight = ((index.row(),index.column()),value)
                return True
            elif globals_.domainHighlight:
                self.domainHighlight[index.row(),index.column()]=value
                return True
            else:
                self.background[index.row(),index.column()] = value
                return True
        elif role == Qt.TextAlignmentRole:
            self.alignmentDict[index.row(),index.column()] = value
            return True
        elif role == Qt.FontRole:
            self.fonts[index.row(),index.column()] = value
            return True
        elif role == Qt.ForegroundRole:
            self.foreground[index.row(),index.column()] = value
            return True
    
    def flags(self,index):
        if index.isValid():
            return Qt.ItemFlags(Qt.ItemIsSelectable|Qt.ItemIsEditable|
                    Qt.ItemIsEnabled|Qt.ItemIsDragEnabled|Qt.ItemIsDropEnabled)

    def supportedDropActions(self):
        return Qt.MoveAction|Qt.CopyAction

    def supportedDragActions(self):
        return Qt.MoveAction|Qt.CopyAction

    def headerData(self,section,orientation,role=Qt.DisplayRole):
        '''Returns the appropiate letter or letters for the corresponding column'''
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if section < 26:
                    return globals_.ALPHABET[section]
                elif section < 702:
                    firstLetter = globals_.ALPHABET[section//26-1]
                    secondLetter = globals_.ALPHABET[section-26*(section//26)]
                    return firstLetter+secondLetter
                else:
                    firstLetter = globals_.ALPHABET[((section-26)//676-1)]
                    secondLetter = globals_.ALPHABET[(section-((section-26)//676*676))//26-1]
                    thirdLetter = globals_.ALPHABET[section-(676+((section-676)//26*26))]
                    return firstLetter+secondLetter+thirdLetter
        elif orientation == Qt.Vertical:
            if role == Qt.DisplayRole:
                return section+1


class CircularReferenceError(Exception):
    def __init__(self,row,column):
        self.formulaRow = row
        self.formulaColumn = column
    def __str__(self):
        errorInfo = "Formula in row {}, column {}, creates a circular reference".format(self.formulaRow,
                self.formulaColumn)
        return errorInfo

