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

from PySide6.QtCore import Qt, QEvent, QPoint, QRect, QSize
from PySide6.QtWidgets import QTableView, QAbstractItemView, QApplication, QWidget
from PySide6.QtGui import QAction, QColor, QPainter, QPen, QBrush
from MyModel import CircularReferenceError
import globals_


class MyView(QTableView):
    def __init__(self,parent=None):
        '''Initialization of parameters and actions for context menu'''
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
        self.hScrollBar.actionTriggered.connect(self.addColumn_)
        self.vScrollBar.sliderMoved.connect(self.disableAddRow)
        self.hScrollBar.sliderMoved.connect(self.disableAddColumn)
        self.vScrollBar.sliderReleased.connect(self.enableAddRow)
        self.hScrollBar.sliderReleased.connect(self.enableAddColumn)
        self.horizontalHeader().sectionResized.connect(self.overlay.createRect)
        self.verticalHeader().sectionResized.connect(self.overlay.createRect)
        copy = QAction('Copy', self)
        copy.setShortcut('Ctrl+C')
        copy.setStatusTip('Copy selected')
        copy.triggered.connect(self.copyAction)
        cut = QAction('Cut', self)
        cut.setShortcut('Ctrl+X')
        cut.setStatusTip('Cut selected')
        cut.triggered.connect(self.cutAction)
        paste = QAction('Paste', self)
        paste.setShortcut('Ctrl+V')
        paste.setStatusTip('Paste from local')
        paste.triggered.connect(self.pasteAction)
        paste.setDisabled(True)
        merge = QAction('Merge cells',self)
        merge.setObjectName('merge')
        merge.setStatusTip('Merge selected cells')
        merge.triggered.connect(self.mergeCells)
        unmerge = QAction('Unmerge cells',self)
        unmerge.setObjectName('unmerge')
        unmerge.setStatusTip('Unmerge selected cells')
        unmerge.triggered.connect(self.unmergeCells)
        self.addActions((copy,cut,paste,merge,unmerge))


    def eventFilter(self,obj,event):
        '''Filters events from various objects intended to provide expected response'''
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Equal:
                globals_.formula_mode = True
                parentWidget = self.parentWidget()
                parentWidget.commandLineEdit.setFocus(Qt.ShortcutFocusReason)
                parentWidget.commandLineEdit.insert('=')
                self.model().setData(self.currentIndex(),QColor(199,196,26),
                        role=Qt.BackgroundRole)
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
                newCurrentIndex = self.model().index(rowOfCurrentIndex+1,
                        columnOfCurrentIndex)
                self.setCurrentIndex(newCurrentIndex)
                return True
            else:
                return False
        else:
            return False

    def mergeCells(self):
        '''Basic merge cell funcionality'''
        selectionModel = self.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        rows = []
        columns = []
        for index in selectedIndexes:
            rows.append(index.row())
            columns.append(index.column())
        topLeftIndex = self.model().index(min(rows),min(columns))
        bottomRightIndex = self.model().index(max(rows),max(columns))
        height = bottomRightIndex.row() - topLeftIndex.row() + 1
        width = bottomRightIndex.column() - topLeftIndex.column() + 1
        if height * width == len(selectedIndexes):
            self.setSpan(topLeftIndex.row(),topLeftIndex.column(),height,width)

    def unmergeCells(self):
        '''Basic unmerge cell funcionality'''
        selectionModel = self.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        if len(selectedIndexes) > 1:
            rows = []
            columns = []
            for index in selectedIndexes:
                rows.append(index.row())
                columns.append(index.column())
            topLeftIndex = self.model().index(min(rows),min(columns))
            bottomRightIndex = self.model().index(max(rows),max(columns))
            height = bottomRightIndex.row() - topLeftIndex.row() + 1
            width = bottomRightIndex.column() - topLeftIndex.column() + 1
            if height * width == len(selectedIndexes):
                self.setSpan(topLeftIndex.row(),topLeftIndex.column(),1,1)

        

    def selectionChanged(self,selected,deselected):
        '''Handles proper response under corresponding state for selection changes'''
        selectionModel = self.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        if not globals_.formula_mode: 
            commandLineEdit = self.parent().commandLineEdit
            commandLineEdit.currentIndex = self.currentIndex()
            commandLineEdit.clear()
            if len(selectedIndexes) == 1:
                for action in self.actions():
                    if action.objectName() == 'merge':
                        action.setDisabled(True)
                if self.model().domainHighlight:
                    self.model().domainHighlight.clear()
                    globals_.domainHighlight = False
                index = self.currentIndex()
                font = self.model().fonts.get((index.row(),index.column()),globals_.defaultFont)
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
                alignment = self.model().data(index,role=Qt.TextAlignmentRole)
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
                for f in self.model().formulas:
                    if f.addressRow == index.row() and f.addressColumn == index.column():
                        self.parent().commandLineEdit.setText('='+f.text)
                        globals_.domainHighlight = True
                        for d in f.domain:
                            coloredIndex = self.model().index(d[0],d[1])
                            self.model().setData(coloredIndex,QColor(119,242,178),role=Qt.BackgroundRole)
                        break
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
                    newText = globals_.REGEXP3.sub(model.getAlphanumeric(column,row),text0)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif globals_.REGEXP4.search(text0):
                    newText = globals_.REGEXP4.sub(model.getAlphanumeric(column,row),text0)
                    finalText = newText +text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.startswith('='):
                    newText = text0 + model.getAlphanumeric(column,row)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.endswith(('*','+','-','/','(',',')):
                    newText = text0 + model.getAlphanumeric(column,row)
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
                topLeftIndex = model.index(min(rows),min(columns)) 
                bottomRightIndex = model.index(max(rows),max(columns))
                alphanumeric1 = model.getAlphanumeric(topLeftIndex.column(),topLeftIndex.row())
                alphanumeric2 = model.getAlphanumeric(bottomRightIndex.column(),bottomRightIndex.row())
                if globals_.REGEXP3.search(text0):
                    newText = globals_.REGEXP3.sub('['+alphanumeric1+':'+alphanumeric2+']',text0)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif globals_.REGEXP4.search(text0):
                    newText = globals_.REGEXP4.sub('['+alphanumeric1+':'+alphanumeric2+']',text0)
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition  += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.startswith('='):
                    newText = text0 + '[' + alphanumeric1 + ':' + alphanumeric2 + ']'
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                elif text0.endswith(('*','+','-','/','(',',')):
                    newText = text0 + '[' + alphanumeric1 + ':' + alphanumeric2 + ']'
                    finalText = newText + text1
                    commandLineEdit.clear()
                    commandLineEdit.setText(finalText)
                    if text1:
                        cursorPosition += len(newText) - len(text0)
                        commandLineEdit.setCursorPosition(cursorPosition)
                commandLineEdit.setFocus(Qt.OtherFocusReason)
                commandLineEdit.deselect()
        super().selectionChanged(selected,deselected)
        self.overlay.createRect()

    def createFormula(self,text,arrayRanges,scalars,domain):
        '''Checks for formula integrity and calls the formula constructor'''
        indexes = []
        for array in arrayRanges:
            rowLimit1 = array[0][0]
            rowLimit2 = array[1][0]+1
            columnLimit1 = array[0][1]
            columnLimit2 = array[1][1]+1
            for row in range(rowLimit1,rowLimit2):
                for column in range(columnLimit1,columnLimit2):
                    indexes.append((row,column))
        if scalars:
            indexes += scalars
        domainIndexes = []
        resultIndexRow = domain['resultIndexRow']
        resultRows = domain['resultRows']
        resultIndexColumn = domain['resultIndexColumn']
        resultColumns = domain['resultColumns']
        address = resultIndexRow,resultIndexColumn
        for row in range(resultIndexRow,resultIndexRow+resultRows):
            for column in range(resultIndexColumn,resultIndexColumn+resultColumns):
                domainIndexes.append((row,column))
        indexesSet = set(indexes)
        domainIndexesSet = set(domainIndexes)
        if indexesSet.intersection(domainIndexesSet):
            raise CircularReferenceError(resultIndexRow,resultIndexColumn)
        if self.model().formulas:
            for f_ in self.model().formulas:
                formulaIndexSet = set(f_.indexes)
                if domainIndexesSet.intersection(formulaIndexSet):
                    self.circularReferenceCheck(indexesSet,f_)
            for f_ in self.model().formulas:
                if f_.addressRow == resultIndexRow and f_.addressColumn == resultIndexColumn:
                    if f_.text != text:
                        self.model().formulas.remove(f_)
                        self.model().formulas.append(Formula(text,address,indexes,domainIndexes))
                        break
                    else:
                        break
            else:
                self.model().formulas.append(Formula(text,address,indexes,domainIndexes))
        else:
            self.model().formulas.append(Formula(text,address,indexes,domainIndexes))

    def circularReferenceCheck(self,subject,match):
        '''Checks for possible circular references which are not allowed by design'''
        formulaDomainSet = set(match.domain)
        if formulaDomainSet.intersection(subject):
            index = self.parent().commandLineEdit.currentIndex
            raise CircularReferenceError(index.row(),index.column())
        else:
            for f in self.model().formulas:
                formulaIndexSet = set(f.indexes)
                if formulaDomainSet.intersection(formulaIndexSet):
                    self.circularReferenceCheck(subject,f)

    def startDrag(self,supportedActions):
        '''Begins dragging operation'''
        if globals_.drag:
            super().startDrag(Qt.MoveAction)
            if globals_.historyIndex != -1:
                hIndex = globals_.historyIndex + len(self.model().history) + 1
                self.model().history = self.model().history[:hIndex]
            self.saveToHistory()
            globals_.drag = False

    def dropEvent(self, event):
        '''Initialize the drop action with the corresponding move or copy action'''
        data = event.mimeData()
        index = self.indexAt(event.pos())
        if event.keyboardModifiers() == Qt.ControlModifier:
            self.model().dropMimeData(data,Qt.CopyAction,-1,-1,index)
        else:
            self.model().dropMimeData(data,Qt.MoveAction,-1,-1,index)
        
    def enableAddRow(self):
        '''Sets corresponding boolean attribute'''
        self.addRowBool = True
    
    def disableAddRow(self):
        '''Sets corresponding boolean attribute'''
        self.addRowBool = False

    def enableAddColumn(self):
        '''Sets corresponding boolean attribute'''
        self.addColumnBool = True

    def disableAddColumn(self):
        '''Sets corresponding boolean attribute'''
        self.addColumnBool = False
    
    def addRow_(self,action):
        '''Adds one row to model and view'''
        if self.addRowBool:
            sliderPosition = self.vScrollBar.sliderPosition()
            if sliderPosition == self.vScrollBar.maximum():
                self.model().insertRows(self.model().rowCount(),1)
                self.vScrollBar.setMaximum(self.model().rowCount())
                self.vScrollBar.setSliderPosition(self.vScrollBar.maximum())

    def addColumn_(self,action):
        '''Adds one column to model and view'''
        #Currently view and model are capped to 18278 columns
        #This is temporary while a proper algorithm is developed that handles
        #any amount of columns that returns a proper alphanumeric code
        if self.model().columnCount() < 18278:
            if self.addColumnBool:
                sliderPosition = self.hScrollBar.sliderPosition()
                if sliderPosition == self.hScrollBar.maximum():
                    self.model().insertColumns(self.model().columnCount(),1)
                    self.hScrollBar.setSliderPosition(self.hScrollBar.maximum())

    def saveToHistory(self):
        '''Saves current model and formulas state up to 5 instances'''
        if len(self.model().history) == 5:
            self.model().history = self.model().history[1:]
        data = self.model().dataContainer.copy()
        align = self.model().alignmentDict.copy()
        fonts = self.model().fonts.copy()
        foreground = self.model().foreground.copy()
        background = self.model().background.copy()
        self.model().history.append((data,self.model().formulas.copy(),align,fonts,foreground,background))
        globals_.historyIndex = -1

    def redo(self):
        '''Basic redo functionality'''
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
        self.model().formulas = formulas.copy()
        self.model().dataContainer = data.copy()
        self.model().alignmentDict = alignments.copy()
        self.model().fonts = fonts.copy()
        self.model().foreground = foreground.copy()
        self.model().background = background.copy()
        rows,columns = data.shape
        startIndex = self.model().index(0,0)
        endIndex = self.model().index(rows-1,columns-1)
        self.model().dataChanged.emit(startIndex,endIndex)

    def undo(self):
        '''Basic undo functionality'''
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
        self.model().formulas = formulas.copy()
        self.model().dataContainer = data.copy()
        self.model().alignmentDict = alignments.copy()
        self.model().fonts = fonts.copy()
        self.model().foreground = foreground.copy()
        self.model().background = background.copy()
        rows,columns = data.shape
        startIndex = self.model().index(0,0)
        endIndex = self.model().index(rows-1,columns-1)
        self.model().dataChanged.emit(startIndex,endIndex)

    def keyPressEvent(self,event):
        '''Handles keyboard interaction over the view to provide expected behaviour'''
        if event.key() == Qt.Key_Delete or event.key() == Qt.Key_Backspace: 
            if globals_.historyIndex != -1:
                hIndex = globals_.historyIndex + len(self.model().history) + 1
                self.model().history = self.model().history[:hIndex]
            selectionModel = self.selectionModel()
            selectedIndexes = selectionModel.selectedIndexes()
            rows = []
            columns = []
            self.model().formulaSnap = self.model().formulas.copy()
            for selIndex in selectedIndexes:
                self.model().setData(selIndex,'',formulaTriggered='ERASE')
                rows.append(selIndex.row())
                columns.append(selIndex.column())
            if self.model().ftoapply:
                self.model().updateModel_()
            minRow = min(rows)
            minColumn = min(columns)
            maxRow = max(rows)
            maxColumn = max(columns)
            minIndex = self.model().index(minRow,minColumn)
            maxIndex = self.model().index(maxRow,maxColumn)
            self.model().dataChanged.emit(minIndex,maxIndex)
            self.saveToHistory()
        elif event.modifiers() == Qt.ControlModifier:
            if event.key() == Qt.Key_Z:
                self.undo()
            elif event.key() == Qt.Key_R:
                self.redo()
            elif event.key() == Qt.Key_Return:
                if selected := self.selectionModel().selectedIndexes():
                    index2copy = selected[0]
                    data2copy = self.model().dataContainer[index2copy.row(),index2copy.column()]['f0']
                    self.model().formulaSnap = self.model().formulas.copy()
                    for ind in selected[1:]:
                        self.model().setData(ind,data2copy,formulaTriggered='ERASE')
                    if self.model().ftoapply:
                        self.model().updateModel_()
                    self.model().dataChanged.emit(selected[0],selected[-1])
            else:
                super().keyPressEvent(event)
        else:
            super().keyPressEvent(event)


    def mouseMoveEvent(self,event):
        '''Tracks mouse position and checks for dragging handles'''
        posX = event.x() + self.verticalHeader().width()
        posY = event.y() + self.horizontalHeader().height()
        if self.overlay.checkIfContains(QPoint(posX,posY)):
            QApplication.setOverrideCursor(Qt.OpenHandCursor)
        else:
            QApplication.restoreOverrideCursor()
        super().mouseMoveEvent(event)

    def mousePressEvent(self,event):
        '''Checks if a drag operation is to take place'''
        posX = event.x() + self.verticalHeader().width()
        posY = event.y() + self.horizontalHeader().height()
        globals_.drag = self.overlay.checkIfContains(QPoint(posX,posY))
        super().mousePressEvent(event)

    def resizeEvent(self,event):
        """Updates overlay size to view's size"""
        super().resizeEvent(event)
        try:
            assert self.overlay
            self.overlay.setGeometry(self.rect())
        except AssertionError:
            pass

    def copyAction(self):
        '''Basic copy action funcionality'''
        self.pasteMode = Qt.CopyAction
        selectionModel = self.selectionModel()
        selected = selectionModel.selectedIndexes()
        self.mimeDataToPaste = self.model().mimeData(selected,flag='keepTopIndex')
        for action in self.actions():
            if action.iconText() == 'Paste':
                action.setDisabled(False)

    def cutAction(self):
        '''Basic cut action funcionality'''
        self.pasteMode = Qt.MoveAction
        selectionModel = self.selectionModel()
        selected = selectionModel.selectedIndexes()
        self.mimeDataToPaste = self.model().mimeData(selected,flag='keepTopIndex')
        for action in self.actions():
            if action.iconText() == 'Paste':
                action.setDisabled(False)

    def pasteAction(self):
        '''Basic paste action functionality'''
        try:
            assert self.mimeDataToPaste
        except AssertionError:
            return
        parent = self.currentIndex()
        self.model().dropMimeData(self.mimeDataToPaste,self.pasteMode,-1,-1,parent)
        if self.pasteMode == Qt.MoveAction:
            self.mimeDataToPaste = None


class Overlay(QWidget):
    '''Provides visual aesthetic for selections and drag functionality'''
    def __init__(self,parent=None):
        super(Overlay,self).__init__(parent)
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
            rect_.translate(offsetX,offsetY)
            self.rect_ = rect_
            auxRight = QRect(self.rect_.topRight(),QSize(-5,self.rect_.height()))
            self.auxRects.append(auxRight)
            auxLeft = QRect(self.rect_.topLeft(),QSize(5,self.rect_.height()))
            self.auxRects.append(auxLeft)
            auxTop = QRect(self.rect_.topLeft(),QSize(self.rect_.width(),5))
            self.auxRects.append(auxTop)
            auxBottom = QRect(self.rect_.bottomLeft(),QSize(self.rect_.width(),-5))
            self.auxRects.append(auxBottom)
        elif len(selected) > 1:
            for index in selected:
                rows.append(index.row())
                columns.append(index.column())
            topLeftIndex = parent.model().index(min(rows),min(columns))
            bottomRightIndex = parent.model().index(max(rows),max(columns))
            height = bottomRightIndex.row() - topLeftIndex.row() + 1
            width = bottomRightIndex.column() - topLeftIndex.column() + 1
            if height * width == len(selected):
                topLeftCorner = parent.visualRect(topLeftIndex).topLeft()
                bottomRightCorner = parent.visualRect(bottomRightIndex).bottomRight()
                rect_ = QRect(topLeftCorner,bottomRightCorner)
                rect_.translate(offsetX,offsetY)
                self.rect_ = rect_
                auxRight = QRect(self.rect_.topRight(),QSize(-5,self.rect_.height()))
                self.auxRects.append(auxRight)
                auxLeft = QRect(self.rect_.topLeft(),QSize(5,self.rect_.height()))
                self.auxRects.append(auxLeft)
                auxTop = QRect(self.rect_.topLeft(),QSize(self.rect_.width(),5))
                self.auxRects.append(auxTop)
                auxBottom = QRect(self.rect_.bottomLeft(),QSize(self.rect_.width(),-5))
                self.auxRects.append(auxBottom)
            else:
                self.rect_ = None
        self.update()

    def checkIfContains(self,pos):
        '''Checks if mouse is over aux rects'''
        for r in self.auxRects:
            if r.contains(pos):
                return True
        return False

    def sizeHint(self):
        return self.parent().rect()
    
    def paintEvent(self,event):
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
    def __init__(self,text,address,indexes,domain):
        '''Constructor params: text->str, address->tuple, indexes->list, domain->list'''
        self.text = text
        self.addressRow = address[0]
        self.addressColumn = address[1]
        self.indexes = indexes
        self.domain = domain
