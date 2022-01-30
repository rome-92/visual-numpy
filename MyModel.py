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

import copy
import weakref

from PySide6.QtCore import (
    QAbstractTableModel, QMimeData, QByteArray,
    QDataStream, QIODevice,
    Qt, QItemSelectionModel, QModelIndex
    )

import globals_


class MyModel(QAbstractTableModel):
    def __init__(self, parent=None):
        """Initialize the model"""
        super().__init__(parent)
        self.dataContainer = {}
        self.rows = 52
        self.columns = 52
        self.formulas = {}
        self.ftoapply = weakref.WeakSet()
        self.formulaSnap = weakref.WeakSet()
        self.highlight = None
        self.domainHighlight = {}
        self.alignmentDict = {}
        self.fonts = {}
        self.background = {}
        self.foreground = {}
        self.history = []
        self.history.append((
            self.dataContainer.copy(),
            copy.deepcopy(self.formulas),
            self.alignmentDict.copy(),
            self.fonts.copy(),
            self.foreground.copy(),
            self.background.copy()
            ))
        self.thousandsSep = True

    def enableThousandsSep(self):
        """Enable thousands separator"""
        self.thousandsSep = True

    def disableThousandsSep(self):
        """Disable thousands separator"""
        self.thousandsSep = False

    def mimeTypes(self):
        """Provide drop of csv data and app specific data"""
        return ['text/csv', 'application/octet-stream']

    def mimeData(self, indexes, flag='mouse'):
        """Create the MimeData object app specific"""
        mimeData = QMimeData()
        encodedData = QByteArray()
        stream = QDataStream(encodedData, QIODevice.WriteOnly)
        rows = []
        columns = []
        for index in indexes:
            rows.append(index.row())
            columns.append(index.column())
        if flag == 'mouse':
            currentIndexRow = str(self.parent().currentIndex().row())
            currentIndexColumn = str(self.parent().currentIndex().column())
        topLeftIndexRow = str(min(rows))
        topLeftIndexColumn = str(min(columns))
        if flag == 'keepTopIndex':
            currentIndexRow = topLeftIndexRow
            currentIndexColumn = topLeftIndexColumn
        bottomRightIndexRow = str(max(rows))
        bottomRightIndexColumn = str(max(columns))
        stream.writeString(currentIndexRow)
        stream.writeString(currentIndexColumn)
        stream.writeString(topLeftIndexRow)
        stream.writeString(topLeftIndexColumn)
        stream.writeString(bottomRightIndexRow)
        stream.writeString(bottomRightIndexColumn)
        mimeData.setData('application/octet-stream', encodedData)
        return mimeData

    def dropMimeData(self, data, action, row, column, parent):
        """Handle drop action from app specific mime type"""
        if data.hasFormat('application/octet-stream'):
            encodedData = data.data('application/octet-stream')
            stream = QDataStream(encodedData, QIODevice.ReadOnly)
            data = []
            while not stream.atEnd():
                data.append(stream.readString())
            rowFromData, columnFromData, topRow, \
                leftColumn, bottomRow, rightColumn = data
            rowFromData = int(rowFromData)
            columnFromData = int(columnFromData)
            topRow = int(topRow)
            leftColumn = int(leftColumn)
            bottomRow = int(bottomRow)
            rightColumn = int(rightColumn)
            dropRow = parent.row()
            dropColumn = parent.column()
            newRowDifference = dropRow - rowFromData
            newColumnDifference = dropColumn - columnFromData
            newBottomRow = bottomRow + newRowDifference
            newBottomColumn = rightColumn + newColumnDifference
            selectionModel = self.parent().selectionModel()
            selectionModel.clearSelection()
            self.formulaSnap.update(self.formulas.values())
            if newRowDifference > 0:
                if newRowDifference >= (bottomRow - topRow + 1):
                    beginRow = topRow
                    endRow = bottomRow + 1
                    stepRow = 1
                else:
                    beginRow = bottomRow
                    endRow = topRow - 1
                    stepRow = -1
            else:
                beginRow = topRow
                endRow = bottomRow + 1
                stepRow = 1
            if newColumnDifference > 0:
                if newColumnDifference >= (rightColumn - leftColumn + 1):
                    beginColumn = leftColumn
                    endColumn = rightColumn + 1
                    stepColumn = 1
                else:
                    beginColumn = rightColumn
                    endColumn = leftColumn - 1
                    stepColumn = -1
            else:
                beginColumn = leftColumn
                endColumn = rightColumn + 1
                stepColumn = 1
            for row in range(beginRow, endRow, stepRow):
                for column in range(beginColumn, endColumn, stepColumn):
                    movedIndex = self.index(
                        row + newRowDifference,
                        column + newColumnDifference
                        )
                    self.setData(
                        movedIndex,
                        self.dataContainer.get((row, column), ''),
                        mode='m'
                        )
                    if action == Qt.MoveAction:
                        if f := self.formulas.get((row, column), None):
                            possibleF = self.parent().Formula(f)
                            try:
                                self.checkForCircularRef(
                                    possibleF,
                                    newRowDifference,
                                    newColumnDifference
                                    )
                            except Exception as e:
                                print(e)
                            else:
                                del self.formulas[row, column]
                                self.formulas[
                                        possibleF.addressRow,
                                        possibleF.addressColumn] = possibleF
                        self.setData(
                            self.index(row, column),
                            '',
                            mode='m', erase='n'
                            )
                    selectionModel.select(
                        movedIndex,
                        QItemSelectionModel.Select
                        )
            if self.ftoapply:
                main = self.parent().parent()
                order = main.topologicalSort(self.ftoapply)
                main.executeOrder(order)
                self.ftoapply.clear()
            self.dataChanged.emit(
                self.index(topRow, leftColumn),
                self.index(newBottomRow, newBottomColumn)
                )
            self.formulaSnap.clear()
            return True

    def checkForCircularRef(self, formula, *deltas):
        """Check circular reference when formula is to be dropped"""
        rowsDomain = []
        columnsDomain = []
        newRowDifference = deltas[0]
        newColumnDifference = deltas[1]
        for index in formula.domain:
            rowsDomain.append(index[0])
            columnsDomain.append(index[1])
        topLeftDomain = [min(rowsDomain), min(columnsDomain)]
        bottomRightDomain = [max(rowsDomain), max(columnsDomain)]
        topLeftDomain[0] = topLeftDomain[0] + newRowDifference
        topLeftDomain[1] = topLeftDomain[1] + newColumnDifference
        bottomRightDomain[0] = bottomRightDomain[0] + newRowDifference
        bottomRightDomain[1] = bottomRightDomain[1] + newColumnDifference
        newDomain = []
        for row in range(
                topLeftDomain[0], bottomRightDomain[0] + 1):
            for column in range(
                    topLeftDomain[1], bottomRightDomain[1] + 1):
                newDomain.append((row, column))
        formulaIndexesSet = set(formula.indexes)
        newDomainSet = set(newDomain)
        if formulaIndexesSet.intersection(newDomainSet):
            raise CircularReferenceError(
                formula.addressRow, formula.addressColumn
                )
        formula.precedence.clear()
        for f_ in self.formulas.values():
            f_Domain = set(f_.domain)
            f_Indexes = set(f_.indexes)
            if newDomainSet.intersection(f_Indexes):
                formula.precedence.add(f_)
                f_.subsequent.add(formula)
            if formulaIndexesSet.intersection(f_Domain):
                f_.precedence.add(formula)
        if formula.precedence.intersection(formula.subsequent):
            raise CircularReferenceError(
                formula.addressRow,
                formula.addressColumn
                )
        self.parent().circularReferenceCheck(formula.precedence, formula)
        formula.domain = newDomain
        formula.addressRow = formula.addressRow + newRowDifference
        formula.addressColumn = formula.addressColumn + newColumnDifference

    def columnCount(self, parent=QModelIndex()):
        """Return number of columns"""
        return self.columns

    def rowCount(self, parent=QModelIndex()):
        """Return number of rows"""
        return self.rows

    def insertRows(self, row, count, parent=QModelIndex()):
        """Insert specified rows"""
        self.beginInsertRows(parent, row, row + count - 1)
        self.rows += count
        self.endInsertRows()
        return True

    def insertColumns(self, column, count, parent=QModelIndex()):
        """Insert specified columns"""
        self.beginInsertColumns(parent, column, column + count - 1)
        self.columns += count
        self.endInsertColumns()
        return True

    def getAlphanumeric(self, column, row):
        """Get alphanumeric coordinate for the corresponding cell"""
        if column < 26:
            letter = globals_.ALPHABET[column]
            return letter+str(row + 1)
        elif column < 702:
            letter1 = globals_.ALPHABET[column//26-1]
            letter2 = globals_.ALPHABET[column-(column//26*26)]
            return letter1 + letter2 + str(row + 1)
        else:
            letter1 = globals_.ALPHABET[(column-26)//676-1]
            letter2 = globals_.ALPHABET[
                (column-((column-26)//676*676))//26-1
                ]
            letter3 = globals_.ALPHABET[
                column-(676+(column-676)//26*26)
                ]
            return letter1 + letter2 + letter3 + str(row + 1)

    def data(self, index, role=Qt.DisplayRole):
        """Return the appropiate data for the corresponding role"""
        if role == Qt.DisplayRole:
            returnValue = self.dataContainer.get(
                (index.row(), index.column()),
                ''
                )
            try:
                returnValue = complex(returnValue)
                if returnValue.imag == 0:
                    if returnValue.real.is_integer():
                        if self.thousandsSep:
                            return '{:,d}'.format(int(returnValue.real))
                        else:
                            return '{:d}'.format(int(returnValue.real))
                    else:
                        if self.thousandsSep:
                            return '{:,.8f}'.format(returnValue.real)
                        else:
                            return '{:.8f}'.format(returnValue.real)
                else:
                    return str(returnValue).strip('()')
            except ValueError:
                return returnValue
        if role == Qt.BackgroundRole:
            if self.highlight:
                if index.row() == self.highlight[0][0]:
                    if index.column() == self.highlight[0][1]:
                        return self.highlight[1]
            elif self.domainHighlight:
                if (index.row(), index.column()) in self.domainHighlight:
                    return self.domainHighlight[
                        index.row(), index.column()
                        ]
            else:
                return self.background.get(
                    (index.row(), index.column()),
                    globals_.defaultBackground
                    )
        if role == Qt.FontRole:
            return self.fonts.get(
                (index.row(), index.column()),
                globals_.defaultFont
                )
        if role == Qt.TextAlignmentRole:
            return self.alignmentDict.get(
                (index.row(), index.column()),
                int(Qt.AlignLeft | Qt.AlignVCenter)
                )
        if role == Qt.ForegroundRole:
            return self.foreground.get(
                (index.row(), index.column()),
                globals_.defaultForeground
                )

    def setData(
            self, index, value,
            role=Qt.EditRole, *, mode='s', erase='y'):
        """Set the appropiate data for the corresponding role"""
        if index.row() > self.rowCount() - 1:
            self.insertRows(self.rowCount(), 1)
        if index.column() > self.columnCount() - 1:
            self.insertColumns(self.columnCount(), 1)
        if role == Qt.EditRole:
            if str(value) == self.data(index):
                return True
            if value != '':
                self.dataContainer[(index.row(), index.column())] = value
            else:
                del self.dataContainer[index.row(), index.column()]
            try:
                assert self.formulas
            except AssertionError:
                return True
            if mode == 'm':
                if erase == 'y':
                    if f := self.formulas.get(
                            (index.row(), index.column()), None):
                        del self.formulas[index.row(), index.column()]
                for f in self.formulaSnap.copy():
                    if (index.row(), index.column()) in f.indexes:
                        self.ftoapply.add(f)
                        self.formulaSnap.remove(f)
            elif mode == 's':
                if erase == 'y':
                    if f := self.formulas.get(
                            (index.row(), index.column()), None):
                        del self.formulas[index.row(), index.column()]
                for f in self.formulas.values():
                    if (index.row(), index.column()) in f.indexes:
                        self.ftoapply.add(f)
                if self.ftoapply:
                    main = self.parent().parent()
                    order = main.topologicalSort(self.ftoapply)
                    main.executeOrder(order)
                    self.ftoapply.clear()
            return True
        elif role == Qt.BackgroundRole:
            if globals_.formula_mode:
                self.highlight = ((index.row(), index.column()), value)
                return True
            elif globals_.domainHighlight:
                self.domainHighlight[index.row(), index.column()] = value
                return True
            else:
                self.background[index.row(), index.column()] = value
                return True
        elif role == Qt.TextAlignmentRole:
            self.alignmentDict[index.row(), index.column()] = value
            return True
        elif role == Qt.FontRole:
            self.fonts[index.row(), index.column()] = value
            return True
        elif role == Qt.ForegroundRole:
            self.foreground[index.row(), index.column()] = value
            return True

    def flags(self, index):
        """Return allowed flags for model"""
        if index.isValid():
            return Qt.ItemFlags(
                    Qt.ItemIsSelectable
                    | Qt.ItemIsEditable
                    | Qt.ItemIsEnabled
                    | Qt.ItemIsDragEnabled
                    | Qt.ItemIsDropEnabled
                    )

    def supportedDropActions(self):
        """Return supported drop actions"""
        return Qt.MoveAction | Qt.CopyAction

    def supportedDragActions(self):
        """Return supported drag actions"""
        return Qt.MoveAction | Qt.CopyAction

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """Return the appropiate letters for the corresponding column"""
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
                    secondLetter = \
                        globals_.ALPHABET[
                            (section-((section-26)//676*676))//26-1]
                    thirdLetter = \
                        globals_.ALPHABET[section-(676+((section-676)//26*26))]
                    return firstLetter+secondLetter+thirdLetter
        elif orientation == Qt.Vertical:
            if role == Qt.DisplayRole:
                return section+1


class CircularReferenceError(Exception):
    def __init__(self, row, column):
        self.formulaRow = row
        self.formulaColumn = column

    def __str__(self):
        errorInfo = \
            "(row {}, column {}), creates a circular reference".format(
                self.formulaRow,
                self.formulaColumn
                )
        return errorInfo
