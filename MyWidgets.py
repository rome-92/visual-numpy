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

import platform
import numbers
import traceback
import random
import csv
import pickle
import copy
import weakref

from PySide6.QtCore import QTimer, QSize, QEvent, Qt, Signal
from PySide6.QtWidgets import (
    QMainWindow, QLineEdit, QToolBar,
    QLabel, QFileDialog, QMessageBox,
    QFontComboBox, QComboBox, QColorDialog,
    QPushButton
    )
from PySide6.QtGui import (
    QAction, QGuiApplication,
    QActionGroup, QFont,
    QPixmap, QBrush, QColor)
from PySide6 import __version__ as PYSIDE6_VERSION
from PySide6.QtCore import __version__ as QT_VERSION
import numpy as np

from MyView import MyView
from MyModel import MyModel
from MyDelegate import MyDelegate
import rcIcons
import globals_

version = '3.0.0-alpha.5'
MAGIC_NUMBER = 0x2384E
FILE_VERSION = 3


class MainWindow(QMainWindow):
    styleSheet = """
    MyView {selection-background-color: rgba(173, 255, 115, 10%);
    selection-color: black;}"""
    currentFile = None

    def __init__(self, parent=None):
        """Initialize  MainWindow widgets and actions"""
        super().__init__(parent)
        self.commandLineEdit = CommandLineEdit()
        self.view = MyView(self)
        self.view.setModel(MyModel(self.view))
        self.view.setItemDelegate(MyDelegate(self.view))
        self.setCentralWidget(self.view)
        self.fontColor = QPushButton(self)
        self.fontColor.setObjectName('FontColor')
        self.fontColor.color_ = (240, 240, 240, 255)
        self.fontColor.setIcon(QPixmap(':text_color.png'))
        self.fontColor.setIconSize(QSize(25, 25))
        self.fontColor.setToolTip('Text color')
        self.backgroundColor = QPushButton(self)
        self.backgroundColor.setIcon(QPixmap(':background_color.png'))
        self.backgroundColor.setIconSize(QSize(25, 25))
        self.backgroundColor.setToolTip('Background color')
        self.backgroundColor.setObjectName('BackgroundColor')
        self.backgroundColor.color_ = (240, 240, 240, 255)
        self.fontColor.clicked.connect(self.showColorDialog)
        self.backgroundColor.clicked.connect(self.showColorDialog)
        newFile = QAction('&New File', self)
        newFile.setShortcut('Ctrl+N')
        newFile.setStatusTip('Create new file')
        newFile.triggered.connect(self.createNew)
        exportFile = QAction('&Export', self)
        exportFile.setShortcut('Ctrl+E')
        exportFile.setStatusTip('Export File to csv')
        exportFile.triggered.connect(self.fileExport)
        saveFile = QAction('&Save', self)
        saveFile.setShortcut('Ctrl+S')
        saveFile.setStatusTip('Save File')
        saveFile.triggered.connect(self.saveFile)
        saveFileAs = QAction('Save &As', self)
        saveFileAs.setShortcut('Shift+Ctrl+S')
        saveFileAs.setStatusTip('Save File As')
        saveFileAs.triggered.connect(self.saveFileAs)
        loadFile = QAction('&Open', self)
        loadFile.setShortcut('Ctrl+O')
        loadFile.setStatusTip('Load .vnp file')
        loadFile.triggered.connect(self.loadFile)
        importFile = QAction('&Import', self)
        importFile.setShortcut('Ctrl+I')
        importFile.setStatusTip('Import csv file')
        importFile.triggered.connect(self.importFile)
        thsndsSep = QAction('Thousands separator', self)
        thsndsSep.setStatusTip('Enable/Disable thousands separator')
        thsndsSep.setCheckable(True)
        thsndsSep.setChecked(True)
        thsndsSep.triggered.connect(self.setThousandsSep)
        self.alignL = QAction('Align left', self)
        self.alignL.setStatusTip('Align text to the left')
        self.alignL.setIcon(QPixmap(':align_left.png'))
        self.alignL.setCheckable(True)
        self.alignL.setChecked(True)
        self.alignL.triggered.connect(self.alignLeft)
        self.alignC = QAction('Align center', self)
        self.alignC.setStatusTip('Align text to the center')
        self.alignC.setIcon(QPixmap(':align_center.png'))
        self.alignC.setCheckable(True)
        self.alignC.setChecked(False)
        self.alignC.triggered.connect(self.alignCenter)
        self.alignR = QAction('Align right', self)
        self.alignR.setStatusTip('Align to the right')
        self.alignR.setIcon(QPixmap(':align_right.png'))
        self.alignR.setCheckable(True)
        self.alignR.setChecked(False)
        self.alignR.triggered.connect(self.alignRight)
        self.alignU = QAction('Align top', self)
        self.alignU.setStatusTip('Align text on top')
        self.alignU.setIcon(QPixmap(':align_top.png'))
        self.alignU.setCheckable(True)
        self.alignU.setChecked(False)
        self.alignU.triggered.connect(self.alignUp)
        self.alignM = QAction('Vertically center', self)
        self.alignM.setStatusTip('Align text to the center vertically')
        self.alignM.setIcon(QPixmap(':vertically_center.png'))
        self.alignM.setCheckable(True)
        self.alignM.setChecked(True)
        self.alignM.triggered.connect(self.alignMiddle)
        self.alignD = QAction('Align bottom', self)
        self.alignD.setStatusTip('Align text to bottom')
        self.alignD.setIcon(QPixmap(':align_bottom.png'))
        self.alignD.setCheckable(True)
        self.alignD.setChecked(False)
        self.alignD.triggered.connect(self.alignDown)
        self.boldAction = QAction('Bold', self)
        self.boldAction.setStatusTip('Bold font')
        self.boldAction.setIcon(QPixmap(':bold.png'))
        self.boldAction.setCheckable(True)
        self.boldAction.triggered.connect(self.bold)
        self.italicAction = QAction('Italic', self)
        self.italicAction.setStatusTip('Italicize font')
        self.italicAction.setIcon(QPixmap(':italic.png'))
        self.italicAction.setCheckable(True)
        self.italicAction.triggered.connect(self.italic)
        self.underlineAction = QAction('Underline', self)
        self.underlineAction.setStatusTip('Underline font')
        self.underlineAction.setIcon(QPixmap(':underline.png'))
        self.underlineAction.setCheckable(True)
        self.underlineAction.triggered.connect(self.underline)
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
        merge = QAction('Merge cells', self)
        merge.setObjectName('merge')
        merge.setStatusTip('Merge selected cells')
        merge.triggered.connect(self.mergeCells)
        unmerge = QAction('Unmerge cells', self)
        unmerge.setObjectName('unmerge')
        unmerge.setStatusTip('Unmerge selected cells')
        unmerge.triggered.connect(self.unmergeCells)
        saveArrayAs = QAction('Save A&rray', self)
        saveArrayAs.setShortcut('Ctrl+J')
        saveArrayAs.setStatusTip('Save Array in .npy format')
        saveArrayAs.triggered.connect(self.saveArrayAs)
        self.view.addActions((copy, cut, paste, merge, unmerge, saveArrayAs))
        self.alignmentGroup1 = QActionGroup(self)
        self.alignmentGroup1.addAction(self.alignL)
        self.alignmentGroup1.addAction(self.alignC)
        self.alignmentGroup1.addAction(self.alignR)
        self.alignmentGroup2 = QActionGroup(self)
        self.alignmentGroup2.addAction(self.alignU)
        self.alignmentGroup2.addAction(self.alignM)
        self.alignmentGroup2.addAction(self.alignD)
        about = QAction('&About', self)
        about.setStatusTip('Show about information')
        about.triggered.connect(self.helpAbout)
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(newFile)
        fileMenu.addAction(loadFile)
        fileMenu.addSeparator()
        fileMenu.addAction(saveFile)
        fileMenu.addAction(saveFileAs)
        fileMenu.addSeparator()
        fileMenu.addAction(exportFile)
        fileMenu.addAction(importFile)
        formatMenu = mainMenu.addMenu('For&mat')
        formatMenu.addAction(thsndsSep)
        helpMenu = self.menuBar().addMenu('&Help')
        helpMenu.addAction(about)
        toolBar = QToolBar('Command Toolbar')
        toolBar.setAllowedAreas(Qt.TopToolBarArea)
        toolBar.setMovable(False)
        commandLabel = QLabel('Evaluate :')
        self.commandLineEdit.installEventFilter(self.view)
        toolBar.addWidget(commandLabel)
        toolBar.addWidget(self.commandLineEdit)
        self.addToolBar(Qt.TopToolBarArea, toolBar)
        self.addToolBarBreak(Qt.TopToolBarArea)
        formatToolbar = QToolBar('Format Toolbar')
        formatToolbar.setAllowedAreas(Qt.TopToolBarArea)
        formatToolbar.setMovable(False)
        self.fontsComboBox = QFontComboBox()
        self.pointSize = QComboBox()
        self.pointSize.addItems(globals_.POINT_SIZES)
        self.pointSize.setMaxVisibleItems(10)
        self.pointSize.setStyleSheet("""QComboBox {combobox-popup: 0;}""")
        self.pointSize.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pointSize.setCurrentIndex(6)
        self.fontsComboBox.activated.connect(self.updateFont)
        self.pointSize.activated.connect(self.updateFont)
        formatToolbar.addWidget(self.fontsComboBox)
        formatToolbar.addWidget(self.pointSize)
        formatToolbar.addWidget(self.fontColor)
        formatToolbar.addWidget(self.backgroundColor)
        formatToolbar.addSeparator()
        formatToolbar.addActions([self.alignL, self.alignC, self.alignR])
        formatToolbar.addSeparator()
        formatToolbar.addActions([self.alignU, self.alignM, self.alignD])
        formatToolbar.addSeparator()
        formatToolbar.addActions([
            self.boldAction,
            self.italicAction,
            self.underlineAction
            ])
        self.addToolBar(Qt.TopToolBarArea, formatToolbar)
        self.commandLineEdit.returnCommand.connect(self.calculate)
        self.statusBar()
        self.setStyleSheet(self.styleSheet)
        globals_.currentFont = QFont(self.fontsComboBox.currentFont())
        globals_.defaultFont = QFont(globals_.currentFont)
        globals_.defaultForeground = QBrush(Qt.black)
        globals_.defaultBackground = QBrush(Qt.white)
        QTimer.singleShot(0, self.center)

    def center(self):
        """Center MainwWindow on the screen"""
        qr = self.frameGeometry()
        cp = QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def sizeHint(self):
        """Set default size"""
        return QSize(1280, 720)

    def createNew(self):
        """Create a new file and clear history"""
        self.view.model().dataContainer = {}
        self.view.model().formulas.clear()
        self.view.model().alignmentDict.clear()
        self.view.model().fonts.clear()
        self.view.model().foreground.clear()
        self.view.model().background.clear()
        self.view.model().history.clear()
        self.view.model().history.append((
            self.view.model().dataContainer.copy(),
            copy.deepcopy(self.view.model().formulas),
            self.view.model().alignmentDict.copy(),
            self.view.model().fonts.copy(),
            self.view.model().foreground.copy(),
            self.view.model().background.copy()
            ))

    def importFile(self, file=None):
        """Import csv files"""
        if not file:
            name, notUsed = QFileDialog.getOpenFileName(
                self,
                'Import File',
                '',
                'csv files (*.csv)'
                )
        else:
            name = file
        if name:
            try:
                with open(name, encoding='latin', newline='') as myFile:
                    reader = csv.reader(myFile, dialect='excel')
                    self.view.model().dataContainer = {}
                    self.view.model().formulas.clear()
                    self.view.model().alignmentDict.clear()
                    self.view.model().fonts.clear()
                    self.view.model().foreground.clear()
                    self.view.model().background.clear()
                    self.view.model().history.clear()
                    self.view.model().history.append((
                        self.view.model().dataContainer.copy(),
                        copy.deepcopy(self.view.model().formulas),
                        self.view.model().alignmentDict.copy(),
                        self.view.model().fonts.copy(),
                        self.view.model().foreground.copy(),
                        self.view.model().background.copy()
                        ))
                    for rowNumber, row in enumerate(reader):
                        for columnNumber, column in enumerate(row):
                            index = \
                                self.view.model().createIndex(
                                    rowNumber,
                                    columnNumber
                                    )
                            self.view.model().setData(
                                index, column,
                                mode='a'
                                )
                    self.view.model().dataChanged.emit(
                        self.view.model().index(0, 0),
                        index
                        )
                MainWindow.currentFile = name
                info = name+' was succesfully imported'
                self.statusBar().showMessage(info, 5000)
                self.view.saveToHistory()
            except Exception as e:
                print(e)
                info = 'There was an error importing '+name
                self.statusBar().showMessage(info, 5000)

    def saveFile(self):
        """Save file into .vnp format"""
        if MainWindow.currentFile:
            name = MainWindow.currentFile
            model = self.view.model().dataContainer
            alignment = self.view.model().alignmentDict
            fonts = self.view.model().fonts.copy()
            foreground = self.view.model().foreground.copy()
            background = self.view.model().background.copy()
            self.encodeFonts(fonts)
            self.encodeColors(foreground)
            self.encodeColors(background)
            formulas = copy.deepcopy(self.view.model().formulas)
            self.prepareFormulas(formulas)
            with open(name, 'wb') as myFile:
                pickle.dump(MAGIC_NUMBER, myFile)
                pickle.dump(FILE_VERSION, myFile)
                pickle.dump(model, myFile)
                pickle.dump(formulas, myFile)
                pickle.dump(alignment, myFile)
                pickle.dump(fonts, myFile)
                pickle.dump(foreground, myFile)
                pickle.dump(background, myFile)
            info = name+' was succesfully saved'
            self.statusBar().showMessage(info, 5000)
        else:
            self.saveFileAs()

    def prepareFormulas(self, f):
        """Substitute weakrefs objects for pickling"""
        for k in f:
            f[k].subsequent = {
                (v.addressRow, v.addressColumn) for v in f[k].subsequent}
            f[k].precedence = {
                (v.addressRow, v.addressColumn) for v in f[k].precedence}
        return f

    def saveFileAs(self, name):
        """Save file into .vnp format"""
        if not name:
            name, notUsed = QFileDialog.getSaveFileName(
                self,
                'Save File',
                '',
                'vnp files (*.vnp)'
                )
        if name:
            name = name.replace('.vnp', '')
            model = self.view.model().dataContainer
            alignment = self.view.model().alignmentDict
            fonts = self.view.model().fonts.copy()
            foreground = self.view.model().foreground.copy()
            background = self.view.model().background.copy()
            self.encodeFonts(fonts)
            self.encodeColors(foreground)
            self.encodeColors(background)
            formulas = copy.deepcopy(self.view.model().formulas)
            self.prepareFormulas(formulas)
            with open(name+'.vnp', 'wb') as myFile:
                pickle.dump(model, myFile)
                pickle.dump(formulas, myFile)
                pickle.dump(alignment, myFile)
                pickle.dump(fonts, myFile)
                pickle.dump(foreground, myFile)
                pickle.dump(background, myFile)
            info = name + ' was succesfully saved'
            self.statusBar().showMessage(info, 5000)
            MainWindow.currentFile = name

    def encodeFonts(self, fonts):
        """Encode font objects so they can be pickled"""
        for i, f in fonts.items():
            fonts[i] = f.toString()

    def encodeColors(self, brushes):
        """Encode colors so they can be pickled"""
        for i, b in brushes.items():
            color = b.color().name()
            brushes[i] = color

    def decodeFonts(self, fonts):
        """Decode fonts so they can be used"""
        for i, f in fonts.items():
            fonts[i] = QFont()
            fonts[i].fromString(f)

    def decodeColors(self, colors):
        """Decode colors so they can be used"""
        for i, c in colors.items():
            brush = QBrush(QColor(c))
            colors[i] = brush

    def copyAction(self):
        """Basic copy action funcionality"""
        self.pasteMode = Qt.CopyAction
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        self.mimeDataToPaste = self.view.model().mimeData(
            selected,
            flag='keepTopIndex'
            )
        for action in self.view.actions():
            if action.iconText() == 'Paste':
                action.setDisabled(False)

    def cutAction(self):
        """Basic cut action funcionality"""
        self.pasteMode = Qt.MoveAction
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        self.mimeDataToPaste = self.view.model().mimeData(
            selected,
            flag='keepTopIndex'
            )
        for action in self.actions():
            if action.iconText() == 'Paste':
                action.setDisabled(False)

    def pasteAction(self):
        """Basic paste action functionality"""
        try:
            assert self.mimeDataToPaste
        except AssertionError:
            return
        parent = self.view.currentIndex()
        self.view.model().dropMimeData(
            self.mimeDataToPaste,
            self.pasteMode, -1, -1, parent
            )
        if self.pasteMode == Qt.MoveAction:
            self.mimeDataToPaste = None

    def mergeCells(self):
        """Basic merge cell funcionality"""
        selectionModel = self.view.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        rows = []
        columns = []
        for index in selectedIndexes:
            rows.append(index.row())
            columns.append(index.column())
        topLeftIndex = self.view.model().index(min(rows), min(columns))
        bottomRightIndex = self.view.model().index(max(rows), max(columns))
        height = bottomRightIndex.row() - topLeftIndex.row() + 1
        width = bottomRightIndex.column() - topLeftIndex.column() + 1
        if height * width == len(selectedIndexes):
            self.view.setSpan(
                topLeftIndex.row(),
                topLeftIndex.column(),
                height,
                width
                )

    def unmergeCells(self):
        """Basic unmerge cell funcionality"""
        selectionModel = self.view.selectionModel()
        selectedIndexes = selectionModel.selectedIndexes()
        if len(selectedIndexes) > 1:
            rows = []
            columns = []
            for index in selectedIndexes:
                rows.append(index.row())
                columns.append(index.column())
            topLeftIndex = self.view.model().index(min(rows), min(columns))
            bottomRightIndex = self.view.model().index(max(rows), max(columns))
            height = bottomRightIndex.row() - topLeftIndex.row() + 1
            width = bottomRightIndex.column() - topLeftIndex.column() + 1
            if height * width == len(selectedIndexes):
                self.view.setSpan(
                    topLeftIndex.row(),
                    topLeftIndex.column(),
                    1,
                    1
                    )

    def saveArrayAs(self):
        """Save array into .npy array format"""
        name, notUsed = QFileDialog.getSaveFileName(
            self,
            'Save Array',
            '',
            'npy files(*.npy)'
            )
        if name:
            name = name.replace('.npy', '')
            selected = self.view.selectedIndexes()
            rows = []
            columns = []
            for index in selected:
                rows.append(index.row())
                columns.append(index.column())
            topRow = min(rows)
            leftColumn = min(columns)
            bottomRow = max(rows)
            rightColumn = max(columns)
            height = bottomRow - topRow + 1
            width = rightColumn - leftColumn + 1
            if height * width == len(selected):
                array = np.zeros((height, width), dtype=np.complex_)
                for y, row in enumerate(range(topRow, bottomRow+1)):
                    for x, column in enumerate(
                            range(leftColumn, rightColumn+1)):
                        try:
                            number = complex(
                                self.view.model().data(
                                    self.view.model().index(row, column)))
                        except ValueError:
                            info = 'There was an error while saving array'
                            self.statusBar().showMessage(info, 5000)
                            return
                        array[y, x] = number
                np.save(name, array)
                info = name + ' was succesfully saved'
                self.statusBar().showMessage(info, 5000)

    def loadFile(self, file=None):
        """Load .vnp format"""
        if not file:
            name, notUsed = QFileDialog.getOpenFileName(
                self, 'Load File',
                '', 'vnp files (*.vnp)'
                )
        else:
            name = file
        if name:
            try:
                with open(name, 'rb') as myFile:
                    magic = pickle.load(myFile)
                    if magic != MAGIC_NUMBER:
                        raise IOError('File type not recognized')
                    fVer = pickle.load(myFile)
                    if fVer != FILE_VERSION:
                        raise IOError('File version not supported')
                    loadedModel = pickle.load(myFile)
                    formulas = pickle.load(myFile)
                    alignment = pickle.load(myFile)
                    fonts = pickle.load(myFile)
                    foreground = pickle.load(myFile)
                    background = pickle.load(myFile)
                    self.decodeFonts(fonts)
                    self.decodeColors(foreground)
                    self.decodeColors(background)
                    self.view.model().dataContainer = loadedModel
                    self.view.model().alignmentDict = alignment
                    self.view.model().fonts = fonts
                    self.view.model().foreground = foreground
                    self.view.model().background = background
                    self.view.model().formulas.clear()
                    self.view.model().history.clear()
                    self.rebuildFormulas(formulas)
                    self.view.model().formulas = formulas
                    rows = (max(v[0] for v in loadedModel.keys()))
                    columns = (max(v[1] for v in loadedModel.keys()))
                    currentRows = self.view.model().rowCount()
                    currentColumns = self.view.model().columnCount()
                    if (rowsToAdd := (rows + 1) - currentRows) > 0:
                        self.view.model().insertRows(currentRows, rowsToAdd)
                    if (columnsToAdd := (columns + 1) - currentColumns) > 0:
                        self.view.model().insertColumns(
                            currentColumns, columnsToAdd
                            )
                    self.view.model().dataChanged.emit(
                        self.view.model().index(0, 0),
                        self.view.model().index(rows, columns)
                        )
                    self.view.model().history.append((
                        self.view.model().dataContainer.copy(),
                        copy.deepcopy(self.view.model().formulas),
                        self.view.model().alignmentDict.copy(),
                        self.view.model().fonts.copy(),
                        self.view.model().foreground.copy(),
                        self.view.model().background.copy()
                        ))
                MainWindow.currentFile = name
                info = name + ' was succesfully loaded'
                self.statusBar().showMessage(info, 5000)
                self.view.saveToHistory()
            except Exception as e:
                traceback.print_tb(e.__traceback__)
                print(e)
                info = 'There was an error loading '+name
                self.statusBar().showMessage(info, 5000)

    def rebuildFormulas(self, f):
        """Rebuild formulas from loaded file"""
        for k, v in f.items():
            sub = v.subsequent
            prec = v.precedence
            v.subsequent = weakref.WeakSet()
            v.precedence = weakref.WeakSet()
            v.subsequent.update({f[idx] for idx in sub})
            v.precedence.update({f[idx] for idx in prec})

    def fileExport(self, file=None):
        """Export file into .csv format"""
        if not file:
            name, notUsed = QFileDialog.getSaveFileName(
                self,
                'Save File',
                '',
                'csv file (*.csv)'
                )
        else:
            name = file
        if name:
            name = name.replace('.csv', '')
            model = self.view.model()
            rows = []
            columns = []
            for index in self.view.model().dataContainer.keys():
                rows.append(index[0])
                columns.append(index[1])
            bottomRow = max(rows)
            rightColumn = max(columns)
            try:
                with open(name+'.csv', 'w', newline='') as csvFile:
                    writer = csv.writer(
                        csvFile,
                        dialect='excel',
                        delimiter=','
                        )
                    for row in range(0, bottomRow+1):
                        rowToAdd = []
                        for column in range(0, rightColumn+1):
                            data = model.data(
                                model.index(row, column))
                            rowToAdd.append(data)
                        writer.writerow(rowToAdd)
                info = name + ' was succesfully exported'
                self.statusBar().showMessage(info, 5000)
            except Exception as e:
                traceback.print_tb(e.__traceback__)
                print(e)
                info = 'There was an error exporting ' + name
                self.statusBar().showMessage(info, 5000)

    def setThousandsSep(self):
        """Simple enable/disable thousands separator"""
        if self.sender().isChecked():
            self.view.model().enableThousandsSep()
        else:
            self.view.model().disableThousandsSep()

    def alignLeft(self):
        """Left align text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup2.actions():
            if action.isChecked():
                if action.text() == 'Align top':
                    vertical = Qt.AlignTop
                elif action.text() == 'Vertically center':
                    vertical = Qt.AlignVCenter
                else:
                    vertical = Qt.AlignBottom
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(Qt.AlignLeft | vertical)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def alignCenter(self):
        """Center text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup2.actions():
            if action.isChecked():
                if action.text() == 'Align top':
                    vertical = Qt.AlignTop
                elif action.text() == 'Vertically center':
                    vertical = Qt.AlignVCenter
                else:
                    vertical = Qt.AlignBottom
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(Qt.AlignHCenter | vertical)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def alignRight(self):
        """Right align text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup2.actions():
            if action.isChecked():
                if action.text() == 'Align top':
                    vertical = Qt.AlignTop
                elif action.text() == 'Vertically center':
                    vertical = Qt.AlignVCenter
                else:
                    vertical = Qt.AlignBottom
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(Qt.AlignRight | vertical)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def alignUp(self):
        """Top align text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup1.actions():
            if action.isChecked():
                if action.text() == 'Align left':
                    horizontal = Qt.AlignLeft
                elif action.text() == 'Align center':
                    horizontal = Qt.AlignHCenter
                else:
                    horizontal = Qt.AlignRight
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(horizontal | Qt.AlignTop)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def alignMiddle(self):
        """Center text vertically from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup1.actions():
            if action.isChecked():
                if action.text() == 'Align left':
                    horizontal = Qt.AlignLeft
                elif action.text() == 'Align center':
                    horizontal = Qt.AlignHCenter
                else:
                    horizontal = Qt.AlignRight
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(horizontal | Qt.AlignVCenter)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def alignDown(self):
        """Bottom align text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for action in self.alignmentGroup1.actions():
            if action.isChecked():
                if action.text() == 'Align left':
                    horizontal = Qt.AlignLeft
                elif action.text() == 'Align center':
                    horizontal = Qt.AlignHCenter
                else:
                    horizontal = Qt.AlignRight
        for i in selected:
            model.alignmentDict[i.row(), i.column()] = \
                int(horizontal | Qt.AlignBottom)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def updateFont(self):
        """Update fonts from selected text to current font"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        font = self.fontsComboBox.currentFont()
        globals_.currentFont = QFont(font)
        pointSize = float(self.pointSize.currentText())
        globals_.currentFont.setPointSizeF(pointSize)
        for index in selected:
            self.view.model().setData(
                index, globals_.currentFont,
                role=Qt.FontRole
                )
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def bold(self):
        """Set bold to True for text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(model.fonts.get(
                (index.row(), index.column()),
                globals_.defaultFont)
                )
            if self.sender().isChecked():
                font.setBold(True)
            else:
                font.setBold(False)
            model.setData(index, font, role=Qt.FontRole)
        model.dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def italic(self):
        """Set text to italic from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(model.fonts.get(
                (index.row(), index.column()),
                globals_.defaultFont)
                )
            if self.sender().isChecked():
                font.setItalic(True)
            else:
                font.setItalic(False)
            model.setData(index, font, role=Qt.FontRole)
        model.dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def underline(self):
        """Underline text from selected cells"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(
                model.fonts.get((
                    index.row(), index.column()),
                    globals_.defaultFont
                    )
                )
            if self.sender().isChecked():
                font.setUnderline(True)
            else:
                font.setUnderline(False)
            model.setData(index, font, role=Qt.FontRole)
        model.dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def showColorDialog(self):
        """Show color dialog for color selection"""
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        color = QColorDialog.getColor()
        if color.isValid():
            self.sender().color_ = color.getRgb()
        else:
            return
        fontR = str(self.fontColor.color_[0])
        fontG = str(self.fontColor.color_[1])
        fontB = str(self.fontColor.color_[2])
        fontA = str(self.fontColor.color_[3])
        backgroundR = str(self.backgroundColor.color_[0])
        backgroundG = str(self.backgroundColor.color_[1])
        backgroundB = str(self.backgroundColor.color_[2])
        backgroundA = str(self.backgroundColor.color_[3])
        styleSheet = """
        MyView {selection-background-color: rgba(173, 255, 115, 25);
        selection-color: black;}
        QPushButton#FontColor {background-color: rgba(%s, %s, %s, %s);}
        QPushButton#BackgroundColor {background-color: rgba(%s, %s, %s, %s);}
        """ % (
            fontR, fontG, fontB, fontA,
            backgroundR, backgroundG, backgroundB, backgroundA)
        self.setStyleSheet(styleSheet)
        model = self.view.model()
        if not selected:
            return
        for index in selected:
            if self.sender().objectName() == 'FontColor':
                model.setData(index, QBrush(color), role=Qt.ForegroundRole)
            else:
                if globals_.domainHighlight:
                    globals_.domainHighlight = False
                model.setData(index, QBrush(color), role=Qt.BackgroundRole)
        self.view.model().dataChanged.emit(selected[0], selected[-1])
        self.view.saveToHistory()

    def helpAbout(self):
        """Show about dialog"""
        QMessageBox.about(
            self,
            "About Visual Numpy",
            """
            <b>Visual Numpy</b> version {0}
            <p>Copyright &copy; 2021 Román U. Martínez
            <p>Simple spreadsheet application "numpy enabled".
            <p>License: GNU General Public License v3.
            <p>Python {1} - Qt {2} - PySide {3} on {4}
            <p><a href="https://github.com/rome-92/visual-numpy">
            github.com/rome-92/visual-numpy</a>""".format(
                version, platform.python_version(),
                QT_VERSION, PYSIDE6_VERSION,
                platform.system()
                )
            )

    def getCoord(self, index):
        """Get row, column (y, x) coordinates from alphanumeric coord"""
        letters = globals_.LETTERS_REG_EXP.search(index).group()
        numbers = globals_.NUMBERS_REG_EXP.search(index).group()
        if len(letters) < 2:
            column = globals_.ALPHABET.index(letters)
        elif len(letters) == 2:
            column = (1+globals_.ALPHABET.index(letters[0]))*26
            column += globals_.ALPHABET.index(letters[1])
        elif len(letters) == 3:
            column = (1+globals_.ALPHABET.index(letters[0]))*676+26
            column += globals_.ALPHABET.index(letters[1])*26
            column += globals_.ALPHABET.index(letters[2])
        row = int(numbers)-1
        return row, column

    def calculate(self, text, *resultIndex, flag=False):
        """Format string into python executable code and evaluate"""
        print(text)
        arrays = globals_.REGEXP1.findall(text)
        coords = []
        numpyArrayList = []
        for array in arrays:
            topLeft, bottomRight = globals_.REGEXP2.findall(array)
            r1, c1 = self.getCoord(topLeft)
            r2, c2 = self.getCoord(bottomRight)
            coords.append(((r1, c1), (r2, c2)))
            rows = [r1, r2+1]
            columns = [c1, c2+1]
            height = rows[1] - rows[0]
            width = columns[1] - columns[0]
            array = np.zeros((height, width), np.complex_)
            for y, row in enumerate(range(rows[0], rows[1])):
                for x, column in enumerate(range(columns[0], columns[1])):
                    element = self.view.model().dataContainer.get(
                        (row, column),
                        0
                        )
                    try:
                        element = complex(element)
                    except Exception as e:
                        print(e)
                        return
                    array[y, x] = element
            numpyArrayList.append(array)
        idPool = set()
        while len(idPool) != len(numpyArrayList):
            idPool.add(
                'x_'
                + str(random.randint(0, 255))
                + str(random.randint(0, 255))
                )
        commandExecutable = globals_.REGEXP1.sub('{}', text)
        idPool = list(idPool)
        for i1, i2 in zip(idPool, numpyArrayList):
            exec(i1+'= i2')
        commandExecutable = commandExecutable.format(*idPool)
        scalars = globals_.REGEXP2.findall(commandExecutable)
        scalarsIndexes = []
        scalarNumbers = []
        for scalar in scalars:
            r, c = self.getCoord(scalar)
            scalarsIndexes.append((r, c))
            number = self.view.model().dataContainer.get((r, c), '0')
            try:
                complex(number)
            except Exception as e:
                print(e, number)
                return
            scalarNumbers.append(number)
        commandExecutable = globals_.REGEXP2.sub('{}', commandExecutable)
        commandExecutable = commandExecutable.format(*scalarNumbers)
        try:
            result = eval(commandExecutable)
        except Exception as e:
            print(e)
            return
        if not flag:
            if globals_.historyIndex != -1:
                hIndex = \
                    globals_.historyIndex + len(self.view.model().history) + 1
                self.view.model().history = self.view.model().history[:hIndex]
        domainDict = {}
        if type(result) is np.ndarray:
            resultIndexRow = resultIndex[0]
            resultIndexColumn = resultIndex[1]
            if len(result.shape) == 2:
                resultRows = result.shape[0]
                resultColumns = result.shape[1]
                domainDict['resultIndexRow'] = resultIndexRow
                domainDict['resultIndexColumn'] = resultIndexColumn
                domainDict['resultRows'] = resultRows
                domainDict['resultColumns'] = resultColumns
                if not flag:
                    try:
                        self.view.createFormula(
                            text, coords, scalarsIndexes, domainDict
                            )
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                for line, rY in zip(
                        result,
                        range(resultIndexRow, resultIndexRow+resultRows)):
                    for dE, cX in zip(
                            line,
                            range(
                                resultIndexColumn,
                                resultIndexColumn
                                + resultColumns)):
                        ind = self.view.model().createIndex(rY, cX)
                        self.view.model().setData(
                            ind,
                            globals_.currentFont,
                            role=Qt.FontRole
                            )
                        self.view.model().setData(
                            ind,
                            dE,
                            mode='a'
                            )
                startIndex = self.view.model().index(
                    resultIndexRow,
                    resultIndexColumn
                    )
                endIndex = self.view.model().index(
                    resultIndexRow+resultRows-1,
                    resultIndexColumn+resultColumns-1)
            elif len(result.shape) == 1:
                resultRows = result.shape[0]
                resultColumns = 1
                domainDict['resultIndexRow'] = resultIndexRow
                domainDict['resultIndexColumn'] = resultIndexColumn
                domainDict['resultRows'] = resultRows
                domainDict['resultColumns'] = resultColumns
                if not flag:
                    try:
                        self.view.createFormula(
                            text,
                            coords,
                            scalarsIndexes,
                            domainDict)
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                for row, ry in zip(
                        result,
                        range(resultIndexRow, resultIndexRow+resultRows)):
                    ind = self.view.model().createIndex(ry, resultIndexColumn)
                    self.view.model().setData(
                        ind,
                        globals_.currentFont,
                        role=Qt.FontRole
                        )
                    self.view.model().setData(ind, row, mode='a')
                startIndex = self.view.model().index(
                    resultIndexRow,
                    resultIndexColumn
                    )
                endIndex = self.view.model().index(
                    resultIndexRow+resultRows-1,
                    resultIndexColumn)
            else:
                return
            self.view.model().dataChanged.emit(startIndex, endIndex)
            self.commandLineEdit.clearFocus()
        elif isinstance(result, numbers.Number):
            resultIndexRow = resultIndex[0]
            resultIndexColumn = resultIndex[1]
            domainDict['resultIndexRow'] = resultIndexRow
            domainDict['resultIndexColumn'] = resultIndexColumn
            domainDict['resultRows'] = 1
            domainDict['resultColumns'] = 1
            if not flag:
                try:
                    self.view.createFormula(
                        text,
                        coords,
                        scalarsIndexes,
                        domainDict
                        )
                except Exception as e:
                    traceback.print_tb(e.__traceback__)
                    print(e)
                    return
            startIndex = self.view.model().createIndex(
                resultIndexRow,
                resultIndexColumn
                )
            endIndex = startIndex
            self.view.model().setData(
                startIndex,
                globals_.currentFont,
                role=Qt.FontRole
                )
            self.view.model().setData(
                startIndex,
                result,
                mode='a'
                )
            self.view.model().dataChanged.emit(startIndex, endIndex)
            self.commandLineEdit.clearFocus()
        currentFormula = self.view.model().formulas[
            resultIndex[0],
            resultIndex[1]
            ]
        if not flag:
            ordered = self.topologicalSort(currentFormula.precedence)
            self.executeOrder(ordered)
            self.view.model().ftoapply.clear()
            self.view.saveToHistory()
        else:
            self.view.saveToHistory()
        self.view.setFocus()

    def executeOrder(self, formulas):
        """Execute formulas in order"""
        for f in formulas:
            self.calculate(
                f.text,
                f.addressRow,
                f.addressColumn,
                flag=True
                )

    def topologicalSort(self, formulas):
        """Create ordered list of formulas"""
        marked = set()
        ordered = []

        def dfs(node):
            marked.add(node)
            for n in node.precedence:
                if n not in marked:
                    dfs(n)
            ordered.append(node)

        for f in formulas:
            if f not in marked:
                dfs(f)
        ordered = list(reversed(ordered))
        return ordered


class CommandLineEdit(QLineEdit):
    """Handle expression and emit corresponding signals"""
    returnCommand = Signal(str, int, int)

    def event(self, event):
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                if globals_.formula_mode:
                    self.clearFocus()
                    globals_.formula_mode = False
                    self.parent().parent().view.setFocus(
                        Qt.ShortcutFocusReason
                        )
                    self.parent().parent().view.model().highlight = None
                    return True
                else:
                    return super().event(event)
            elif event.key() == Qt.Key_Return or \
                    event.key() == Qt.Key_Enter:
                if self.text().startswith("="):
                    globals_.formula_mode = False
                    eval_ = self.text().partition("=")[2]
                    row = self.currentIndex.row()
                    column = self.currentIndex.column()
                    self.parent().parent().view.model().highlight = None
                    self.returnCommand.emit(eval_, row, column)
                    return True
                else:
                    return super().event(event)
            else:
                return super().event(event)
        else:
            return super().event(event)

    def focusInEvent(self, event):
        """Default focus in behaviour"""
        super().focusInEvent(event)
        globals_.formula_mode = True

    def focusOutEvent(self, event):
        """Default focus out behaviour"""
        super().focusOutEvent(event)
        globals_.formula_mode = False
