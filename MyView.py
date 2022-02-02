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

from PySide6.QtCore import Qt, QEvent, QPoint, QRect, QSize, QTimer
from PySide6.QtWidgets import (
    QTableView, QAbstractItemView,
    QApplication, QWidget
    )
from PySide6.QtGui import QColor, QPainter, QPen, QBrush

from MyModel import CircularReferenceError
import globals_


class MyView(QTableView):
    def __init__(self, parent=None):
        """Initialize widgets and actions for context menu"""
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.ActionsContextMenu)
        self.overlay = Overlay(self)
        self.setMouseTracking(True)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.addRowBool = True
        self.addColumnBool = True
        self.vScrollBar = self.verticalScrollBar()
        self.hScrollBar = self.horizontalScrollBar()
        self.vScrollBar.actionTriggered.connect(self.addRow_)
        self.vScrollBar.actionTriggered.connect(
            lambda:QTimer.singleShot(0, self.overlay.createRect))
        self.hScrollBar.actionTriggered.connect(self.addColumn_)
        self.hScrollBar.actionTriggered.connect(
            lambda:QTimer.singleShot(0, self.overlay.createRect))
        self.vScrollBar.sliderMoved.connect(self.disableAddRow)
        self.hScrollBar.sliderMoved.connect(self.disableAddColumn)
        self.vScrollBar.sliderReleased.connect(self.enableAddRow)
        self.hScrollBar.sliderReleased.connect(self.enableAddColumn)
        self.horizontalHeader().sectionResized.connect(
            self.overlay.createRect)
        self.verticalHeader().sectionResized.connect(self.overlay.createRect)

    def eventFilter(self, obj, event):
        """Filter events from various objects"""
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Equal:
                globals_.formula_mode = True
                parentWidget = self.parentWidget()
                parentWidget.commandLineEdit.setFocus(Qt.ShortcutFocusReason)
                parentWidget.commandLineEdit.insert('=')
                self.model().setData(
                    self.currentIndex(),
                    QColor(199, 196, 26),
                    role=Qt.BackgroundRole
                    )
                return True
            elif event.key() == Qt.Key_Left:
                self.event(event)
                return True
            elif event.key() == Qt.Key_Up:
                self.event(event)
                return True
            elif event.key() == Qt.Key_Down:
                self.event(event)
                return True
            elif event.key() == Qt.Key_Right:
                self.event(event)
                return True
            elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
                if obj.metaObject().className() == 'CommandLineEdit':
                    return False
                currentIndex = self.currentIndex()
                rowOfCurrentIndex = currentIndex.row()
                columnOfCurrentIndex = currentIndex.column()
                newCurrentIndex = self.model().index(
                    rowOfCurrentIndex+1,
                    columnOfCurrentIndex
                    )
                self.setCurrentIndex(newCurrentIndex)
                return True
            else:
                return False
        else:
            return False

    def selectionChanged(self, selected, deselected):
        """Handle proper response for selection changes"""
        selectionModel = self.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        if not globals_.formula_mode:
            commandLineEdit = self.parent().commandLineEdit
            commandLineEdit.clear()
            if len(selectedIndexes) == 1:
                commandLineEdit.currentIndex = selectedIndexes[0]
                for action in self.actions():
                    if action.objectName() == 'merge':
                        action.setDisabled(True)
                if self.model().domainHighlight:
                    self.model().domainHighlight.clear()
                    globals_.domainHighlight = False
                index = selectedIndexes[0]
                font = self.model().fonts.get(
                    (index.row(), index.column()),
                    globals_.defaultFont
                    )
                size = font.pointSizeF()
                if size.is_integer():
                    size = int(size)
                self.parent().pointSize.setCurrentText(str(size))
                self.parent().fontsComboBox.setCurrentFont(font)
                if font.bold():
                    self.parent().boldAction.setChecked(True)
                else:
                    self.parent().boldAction.setChecked(False)
                if font.italic():
                    self.parent().italicAction.setChecked(True)
                else:
                    self.parent().italicAction.setChecked(False)
                if font.underline():
                    self.parent().underlineAction.setChecked(True)
                else:
                    self.parent().underlineAction.setChecked(False)
                alignment = self.model().data(
                    index,
                    role=Qt.TextAlignmentRole
                    )
                hexAlignment = hex(alignment)
                if hexAlignment[-1] == '1':
                    self.parent().alignL.setChecked(True)
                elif hexAlignment[-1] == '2':
                    self.parent().alignR.setChecked(True)
                else:
                    self.parent().alignC.setChecked(True)
                if hexAlignment[-2] == '2':
                    self.parent().alignU.setChecked(True)
                elif hexAlignment[-2] == '4':
                    self.parent().alignD.setChecked(True)
                else:
                    self.parent().alignM.setChecked(True)
                if f := self.model().formulas.get(
                        (index.row(), index.column()), None):
                    self.parent().commandLineEdit.setText('='+f.text)
                    globals_.domainHighlight = True
                    for d in f.domain:
                        coloredIndex = self.model().index(d[0], d[1])
                        self.model().setData(
                            coloredIndex,
                            QColor(119, 242, 178),
                            role=Qt.BackgroundRole
                            )
            else:
                for action in self.actions():
                    if action.objectName() == 'merge':
                        action.setEnabled(True)
        elif globals_.formula_mode:
            if len(selectedIndexes) == 1:
                commandLineEdit = self.parent().commandLineEdit
                cursorPosition = commandLineEdit.cursorPosition()
                model = self.model()
                text0 = commandLineEdit.text()[:cursorPosition]
                text1 = commandLineEdit.text()[cursorPosition:]
                row = self.currentIndex().row()
                column = self.currentIndex().column()
                if globals_.REGEXP3.search(text0):
                    newText = globals_.REGEXP3.sub(
                        model.getAlphanumeric(column, row),
                        text0
                        )
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif globals_.REGEXP4.search(text0):
                    newText = globals_.REGEXP4.sub(
                        model.getAlphanumeric(column, row),
                        text0
                        )
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0 == '=':
                    newText = text0 + model.getAlphanumeric(column, row)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.endswith(('*', '+', '-', '/', '(', ',')):
                    newText = text0 + model.getAlphanumeric(column, row)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                commandLineEdit.setFocus(Qt.OtherFocusReason)
                commandLineEdit.deselect()
            else:
                rows = []
                columns = []
                commandLineEdit = self.parent().commandLineEdit
                cursorPosition = commandLineEdit.cursorPosition()
                model = self.model()
                text0 = commandLineEdit.text()[:cursorPosition]
                text1 = commandLineEdit.text()[cursorPosition:]
                for index in selectedIndexes:
                    rows.append(index.row())
                    columns.append(index.column())
                topLeftIndex = model.index(min(rows), min(columns))
                bottomRightIndex = model.index(max(rows), max(columns))
                alphanumeric1 = model.getAlphanumeric(
                    topLeftIndex.column(),
                    topLeftIndex.row()
                    )
                alphanumeric2 = model.getAlphanumeric(
                    bottomRightIndex.column(),
                    bottomRightIndex.row()
                    )
                if globals_.REGEXP3.search(text0):
                    newText = globals_.REGEXP3.sub(
                        '['+alphanumeric1+':'+alphanumeric2+']',
                        text0
                        )
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif globals_.REGEXP4.search(text0):
                    newText = globals_.REGEXP4.sub(
                        '['+alphanumeric1+':'+alphanumeric2+']',
                        text0
                        )
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0 == '=':
                    newText = \
                        text0 + '[' + alphanumeric1 + ':' + alphanumeric2 + ']'
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.endswith(('*', '+', '-', '/', '(', ',')):
                    newText = \
                        text0 + '[' + alphanumeric1 + ':' + alphanumeric2 + ']'
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                commandLineEdit.setFocus(Qt.OtherFocusReason)
                commandLineEdit.deselect()
        super().selectionChanged(selected, deselected)
        self.overlay.createRect()

    def createFormula(self, text, arrayRanges, scalars, domain):
        """Check formula integrity and call the formula constructor"""
        indexes = []
        precedence = weakref.WeakSet()
        subsequent = weakref.WeakSet()
        for array in arrayRanges:
            rowLimit1 = array[0][0]
            rowLimit2 = array[1][0] + 1
            columnLimit1 = array[0][1]
            columnLimit2 = array[1][1] + 1
            for row in range(rowLimit1, rowLimit2):
                for column in range(columnLimit1, columnLimit2):
                    indexes.append((row, column))
        if scalars:
            indexes += scalars
        domainIndexes = []
        resultIndexRow = domain['resultIndexRow']
        resultRows = domain['resultRows']
        resultIndexColumn = domain['resultIndexColumn']
        resultColumns = domain['resultColumns']
        address = resultIndexRow, resultIndexColumn
        for row in range(resultIndexRow, resultIndexRow+resultRows):
            for column in range(
                    resultIndexColumn,
                    resultIndexColumn
                    + resultColumns):
                domainIndexes.append((row, column))
        possibleF = Formula(
            text,
            address,
            indexes,
            domainIndexes,
            precedence,
            subsequent
            )
        indexesSet = set(indexes)
        domainIndexesSet = set(domainIndexes)
        if indexesSet.intersection(domainIndexesSet):
            raise CircularReferenceError(resultIndexRow, resultIndexColumn)
        if self.model().formulas:
            for f_ in self.model().formulas.values():
                formulaIndexSet = set(f_.indexes)
                formulaDomainSet = set(f_.domain)
                if domainIndexesSet.intersection(formulaIndexSet):
                    possibleF.precedence.add(f_)
                    f_.subsequent.add(possibleF)
                if indexesSet.intersection(formulaDomainSet):
                    possibleF.subsequent.add(f_)
                    f_.precedence.add(possibleF)
                if possibleF.precedence.intersection(possibleF.subsequent):
                    raise CircularReferenceError(
                            resultIndexRow,
                            resultIndexColumn
                            )
            self.circularReferenceCheck(possibleF.precedence, possibleF)
            if f_ := self.model().formulas.get(
                    (resultIndexRow, resultIndexColumn), None):
                if f_.text != text:
                    self.model().formulas[resultIndexRow, resultIndexColumn]\
                        = possibleF
            else:
                self.model().formulas[resultIndexRow, resultIndexColumn] \
                    = possibleF
        else:
            self.model().formulas[resultIndexRow, resultIndexColumn] \
                = possibleF

    def circularReferenceCheck(self, precedence, possibleF):
        """Check for circular references"""
        if possibleF in precedence:
            raise CircularReferenceError(
                possibleF.addressRow,
                possibleF.addressColumn
                )
        else:
            for f in precedence:
                if f.precedence:
                    self.circularReferenceCheck(f.precedence, possibleF)

    def startDrag(self, supportedActions):
        """Begin dragging operation"""
        if globals_.drag:
            super().startDrag(Qt.MoveAction)
            if globals_.historyIndex != -1:
                hIndex = globals_.historyIndex + len(self.model().history) + 1
                self.model().history = self.model().history[:hIndex]
            globals_.drag = False

    def dropEvent(self, event):
        """Initialize the drop action with move or copy action"""
        data = event.mimeData()
        index = self.indexAt(event.pos())
        if event.keyboardModifiers() == Qt.ControlModifier:
            self.model().dropMimeData(data, Qt.CopyAction, -1, -1, index)
        else:
            self.model().dropMimeData(data, Qt.MoveAction, -1, -1, index)

    def enableAddRow(self):
        """Set corresponding boolean attribute"""
        self.addRowBool = True

    def disableAddRow(self):
        """Set corresponding boolean attribute"""
        self.addRowBool = False

    def enableAddColumn(self):
        """Set corresponding boolean attribute"""
        self.addColumnBool = True

    def disableAddColumn(self):
        """Set corresponding boolean attribute"""
        self.addColumnBool = False

    def addRow_(self, action):
        """Add one row to model and view"""
        if self.addRowBool:
            sliderPosition = self.vScrollBar.sliderPosition()
            if sliderPosition == self.vScrollBar.maximum():
                self.model().insertRows(self.model().rowCount(), 1)
                self.vScrollBar.setMaximum(self.model().rowCount())
                self.vScrollBar.setSliderPosition(self.vScrollBar.maximum())

    def addColumn_(self, action):
        """Add one column to model and view"""
        if self.model().columnCount() < 18278:
            if self.addColumnBool:
                sliderPosition = self.hScrollBar.sliderPosition()
                if sliderPosition == self.hScrollBar.maximum():
                    self.model().insertColumns(self.model().columnCount(), 1)
                    self.hScrollBar.setSliderPosition(
                        self.hScrollBar.maximum()
                        )

    def saveToHistory(self):
        """Save current model and formulas state up to 5 instances"""
        if len(self.model().history) == 5:
            self.model().history = self.model().history[1:]
        data = self.model().dataContainer.copy()
        formulas = copy.deepcopy(self.model().formulas)
        align = self.model().alignmentDict.copy()
        fonts = self.model().fonts.copy()
        foreground = self.model().foreground.copy()
        background = self.model().background.copy()
        self.model().history.append((
            data, formulas, align,
            fonts, foreground, background
            ))
        globals_.historyIndex = -1

    def redo(self):
        """Basic redo functionality"""
        if globals_.historyIndex == -1:
            return
        globals_.historyIndex += 1
        model = self.model().history[globals_.historyIndex]
        data = model[0]
        formulas = model[1]
        alignments = model[2]
        fonts = model[3]
        foreground = model[4]
        background = model[5]
        self.model().formulas = copy.deepcopy(formulas)
        self.model().dataContainer = data.copy()
        self.model().alignmentDict = alignments.copy()
        self.model().fonts = fonts.copy()
        self.model().foreground = foreground.copy()
        self.model().background = background.copy()
        startIndex = self.model().index(0, 0)
        endIndex = self.model().index(
            self.model().rowCount()-1,
            self.model().columnCount()-1
            )
        self.model().dataChanged.emit(startIndex, endIndex)

    def undo(self):
        """Basic undo functionality"""
        if globals_.historyIndex + len(self.model().history) == 0:
            return
        globals_.historyIndex -= 1
        model = self.model().history[globals_.historyIndex]
        data = model[0]
        formulas = model[1]
        alignments = model[2]
        fonts = model[3]
        foreground = model[4]
        background = model[5]
        self.model().formulas = copy.deepcopy(formulas)
        self.model().dataContainer = data.copy()
        self.model().alignmentDict = alignments.copy()
        self.model().fonts = fonts.copy()
        self.model().foreground = foreground.copy()
        self.model().background = background.copy()
        startIndex = self.model().index(0, 0)
        endIndex = self.model().index(
            self.model().rowCount()-1,
            self.model().columnCount()-1
            )
        self.model().dataChanged.emit(startIndex, endIndex)

    def keyPressEvent(self, event):
        """Handle keyboard interaction over the view"""
        if event.key() == Qt.Key_Delete or event.key() == Qt.Key_Backspace:
            if globals_.historyIndex != -1:
                hIndex = \
                    globals_.historyIndex + len(self.model().history) + 1
                self.model().history = self.model().history[:hIndex]
            selectionModel = self.selectionModel()
            selectedIndexes = selectionModel.selectedIndexes()
            rows = []
            columns = []
            self.model().formulaSnap.update(self.model().formulas.values())
            for selIndex in selectedIndexes:
                self.model().setData(selIndex, '', mode='m')
                rows.append(selIndex.row())
                columns.append(selIndex.column())
            if self.model().ftoapply:
                main = self.parent()
                order = main.topologicalSort(self.model().ftoapply)
                main.executeOrder(order)
                self.model().ftoapply.clear()
            minRow = min(rows)
            minColumn = min(columns)
            maxRow = max(rows)
            maxColumn = max(columns)
            minIndex = self.model().index(minRow, minColumn)
            maxIndex = self.model().index(maxRow, maxColumn)
            self.model().dataChanged.emit(minIndex, maxIndex)
            self.saveToHistory()
        elif event.modifiers() == Qt.ControlModifier:
            if event.key() == Qt.Key_Z:
                self.undo()
            elif event.key() == Qt.Key_G:
                self.auxMethod()
            elif event.key() == Qt.Key_R:
                self.redo()
            elif event.key() == Qt.Key_Return:
                if selected := self.selectionModel().selectedIndexes():
                    index2copy = selected[0]
                    data2copy = self.model().dataContainer[
                        index2copy.row(),
                        index2copy.column()
                        ]
                    self.model().formulaSnap.update(
                        self.model().formulas.values()
                        )
                    for ind in selected[1:]:
                        self.model().setData(
                            ind, data2copy, mode='m'
                            )
                    if self.model().ftoapply:
                        main = self.parent()
                        order = main.topologicalSort(self.ftoapply)
                        main.executeOrder(order)
                        self.model().ftoapply.clear()
                    self.saveToHistory()
                    self.model().dataChanged.emit(selected[0], selected[-1])
            else:
                super().keyPressEvent(event)
        elif event.key() == Qt.Key_F6:
            for f in self.model().formulas.values():
                print(f, f.precedence, f.subsequent)
        else:
            super().keyPressEvent(event)
        self.model().formulaSnap.clear()

    def auxMethod(self):
        for f in self.model().formulas.values():
            subList = [s.text for s in f.subsequent]
            precList = [s.text for s in f.precedence]
            print(f'{f.text}: sub-> {subList}')
            print(f'{f.text}: prec-> {precList}')

    def mouseMoveEvent(self, event):
        """Track mouse position and checks for dragging handles"""
        posX = event.x() + self.verticalHeader().width()
        posY = event.y() + self.horizontalHeader().height()
        if self.overlay.checkIfContains(QPoint(posX, posY)):
            QApplication.setOverrideCursor(Qt.OpenHandCursor)
        else:
            QApplication.restoreOverrideCursor()
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        """Check if a drag operation is to take place"""
        posX = event.x() + self.verticalHeader().width()
        posY = event.y() + self.horizontalHeader().height()
        globals_.drag = self.overlay.checkIfContains(QPoint(posX, posY))
        super().mousePressEvent(event)

    def resizeEvent(self, event):
        """Update overlay size to view's size"""
        super().resizeEvent(event)
        try:
            assert self.overlay
            self.overlay.setGeometry(self.rect())
        except AssertionError:
            pass


class Overlay(QWidget):
    """Provide visual aesthetic for selections and drag functionality"""
    def __init__(self, parent=None):
        super(Overlay, self).__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.rect_ = None
        self.auxRects = []

    def createRect(self):
        parent = self.parent()
        selectionModel = parent.selectionModel()
        selected = selectionModel.selectedIndexes()
        rows = []
        columns = []
        offsetX = parent.verticalHeader().width()
        offsetY = parent.horizontalHeader().height()
        self.auxRects.clear()
        if len(selected) == 1:
            index = selected[0]
            rect_ = parent.visualRect(index)
            rect_.translate(offsetX, offsetY)
            self.rect_ = rect_
            auxRight = QRect(
                self.rect_.topRight(),
                QSize(-5, self.rect_.height())
                )
            self.auxRects.append(auxRight)
            auxLeft = QRect(
                self.rect_.topLeft(),
                QSize(5, self.rect_.height())
                )
            self.auxRects.append(auxLeft)
            auxTop = QRect(
                self.rect_.topLeft(),
                QSize(self.rect_.width(), 5)
                )
            self.auxRects.append(auxTop)
            auxBottom = QRect(
                self.rect_.bottomLeft(),
                QSize(self.rect_.width(), -5)
                )
            self.auxRects.append(auxBottom)
        elif len(selected) > 1:
            for index in selected:
                rows.append(index.row())
                columns.append(index.column())
            topLeftIndex = parent.model().index(min(rows), min(columns))
            bottomRightIndex = parent.model().index(max(rows), max(columns))
            height = bottomRightIndex.row() - topLeftIndex.row() + 1
            width = bottomRightIndex.column() - topLeftIndex.column() + 1
            if height * width == len(selected):
                topLeftCorner = parent.visualRect(topLeftIndex).topLeft()
                bottomRightCorner = \
                    parent.visualRect(bottomRightIndex).bottomRight()
                rect_ = QRect(topLeftCorner, bottomRightCorner)
                rect_.translate(offsetX, offsetY)
                self.rect_ = rect_
                auxRight = QRect(
                    self.rect_.topRight(),
                    QSize(-5, self.rect_.height())
                    )
                self.auxRects.append(auxRight)
                auxLeft = QRect(
                    self.rect_.topLeft(),
                    QSize(5, self.rect_.height())
                    )
                self.auxRects.append(auxLeft)
                auxTop = QRect(
                    self.rect_.topLeft(),
                    QSize(self.rect_.width(), 5)
                    )
                self.auxRects.append(auxTop)
                auxBottom = QRect(
                    self.rect_.bottomLeft(),
                    QSize(self.rect_.width(), -5)
                    )
                self.auxRects.append(auxBottom)
            else:
                self.rect_ = None
        self.update()

    def checkIfContains(self, pos):
        """Check if mouse is over aux rects"""
        for r in self.auxRects:
            if r.contains(pos):
                return True
        return False

    def sizeHint(self):
        """Default size"""
        return self.parent().rect()

    def paintEvent(self, event):
        """Paint widget"""
        if self.rect_:
            painter = QPainter()
            painter.begin(self)
            pen = QPen(Qt.green)
            brush = QBrush(Qt.transparent)
            pen.setWidth(2)
            painter.setPen(pen)
            painter.setBrush(brush)
            painter.drawRect(self.rect_)
            painter.end()


class Formula():
    def __init__(self, *args):
        """Constructor for Formula object"""
        if len(args) > 1:
            self.text = args[0]
            self.addressRow = args[1][0]
            self.addressColumn = args[1][1]
            self.indexes = tuple(args[2])
            self.domain = tuple(args[3])
            self.precedence = args[4]
            self.subsequent = args[5]
        elif len(args) == 1:
            self.text = args[0].text
            self.addressRow = args[0].addressRow
            self.addressColumn = args[0].addressColumn
            self.indexes = args[0].indexes
            self.domain = args[0].domain
            self.precedence = args[0].precedence
            self.subsequent = args[0].subsequent
        weakref.finalize(self, print, 'Formula {} killed'.format(self.text))

    def __repr__(self):
        return self.text
