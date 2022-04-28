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

from PySide6.QtCore import (
    QTimer, QSize,
    QEvent, Qt, Signal
    )
from PySide6.QtWidgets import (
    QMainWindow, QLineEdit, QToolBar,
    QLabel, QFileDialog, QMessageBox,
    QFontComboBox, QComboBox, QColorDialog,
    QPushButton, QVBoxLayout, QWidget,
    QGridLayout, QGraphicsScene, QGraphicsView,
    QSplitter, QStackedWidget, QCheckBox
    )
from PySide6.QtGui import (
    QAction, QGuiApplication,
    QActionGroup, QFont,
    QPixmap, QBrush, QColor
    )
from PySide6 import __version__ as PYSIDE6_VERSION
from PySide6.QtCore import __version__ as QT_VERSION
try:
    from matplotlib.backends.backend_qtagg import (
        FigureCanvas,
        NavigationToolbar2QT as NavigationToolbar
        )
    from matplotlib.figure import Figure
except ImportError:
    PLOT = False
else:
    PLOT = True
import numpy as np

from MyView import MyView
from MyModel import MyModel
from MyDelegate import MyDelegate
import rcIcons
import globals_

version = '3.0.0-alpha.12'
MAGIC_NUMBER = 0x2384E
FILE_VERSION = 4


class MainWindow(QMainWindow):
    styleSheet = """
    MyView {selection-background-color: rgba(173, 255, 115, 10%);
    selection-color: black;}"""
    currentFile = None

    def __init__(self, parent=None):
        """Initialize  MainWindow widgets and actions"""
        super().__init__(parent)
        self.plotMenu = None
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
        fastPlot = QAction('Fast plot', self)
        fastPlot.setShortcut('Ctrl+L')
        fastPlot.setStatusTip('Plot given x and y arrays')
        fastPlot.triggered.connect(self.plot)
        plot = QAction('Plot', self)
        plot.triggered.connect(self.showPlotMenu)
        self.view.addActions((
            copy, cut, paste, merge, unmerge, saveArrayAs,
            fastPlot
            ))
        if not PLOT:
            plot.setDisabled(True)
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
        plotMenu = mainMenu.addMenu('Plot')
        plotMenu.addAction(plot)
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
        self.commandLineEdit.returnCommand.connect(
            lambda t, row, col, c: self.calculate(t, row, col, com=c))
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
                (v.row, v.col) for v in f[k].subsequent}
            f[k].precedence = {
                (v.row, v.col) for v in f[k].precedence}
        return f

    def saveFileAs(self, name=None):
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
                pickle.dump(MAGIC_NUMBER, myFile)
                pickle.dump(FILE_VERSION, myFile)
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
        for action in self.view.actions():
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

    def showPlotMenu(self):
        self.plotMenu = PlotMenu(self)
        self.plotMenu.model = self.view.model()
        self.plotMenu.show()

    def plot(self):
        model = self.view.model()
        selected = self.view.selectedIndexes()
        rows = []
        columns = []
        if not selected:
            return
        if len(selected) == 2:
            x_r, x_c = selected[0].row(), selected[0].column()
            y_r, y_c = selected[1].row(), selected[1].column()
            x_ = model.dataContainer[x_r, x_c]
            y_ = model.dataContainer[y_r, y_c]
            if type(x_) is np.ndarray and type(y_) is np.ndarray:
                x = x_
                y = y_
            else:
                return
        else:
            for index in selected:
                rows.append(index.row())
                columns.append(index.column())
            topRow = min(rows)
            leftColumn = min(columns)
            bottomRow = max(rows)
            rightColumn = max(columns)
            height = bottomRow - topRow + 1
            width = rightColumn - leftColumn + 1
            if width > 2:
                return
            if height * width == len(selected):
                x = np.zeros((height), dtype=np.float)
                y = np.zeros((height), dtype=np.float)
                for ry, r in enumerate(range(topRow, bottomRow+1)):
                    for cx, c in enumerate(range(leftColumn, rightColumn+1)):
                        try:
                            number = complex(model.dataContainer[r, c])
                        except ValueError as e:
                            print(e)
                            return
                        else:
                            if cx == 0:
                                x[ry] = number.real
                            else:
                                y[ry] = number.real
        print('printing x, y')
        print(x, y)
        self.plotWidget = PlotWidget()
        self.plotWidget.show()
        self.plotWidget.ax.plot(x, y)

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

    @staticmethod
    def getCoord(index):
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

    def calculate(self, text, *ridx, com=False, flag=False):
        """Format string into python executable code and evaluate"""
        print(text)
        arrays = globals_.REGEXP1.findall(text)
        coords = []
        numpyArrayList = []
        model = self.view.model()
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
                    element = model.dataContainer.get(
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
            exec(i1 + '= i2')
        commandExecutable = commandExecutable.format(*idPool)
        single = globals_.REGEXP2.findall(commandExecutable)
        singleIndexes = []
        singleElements = []
        idPool2 = set()
        for s in single:
            r, c = self.getCoord(s)
            singleIndexes.append((r, c))
            element = model.dataContainer.get((r, c), '0')
            if isinstance(element, np.ndarray):
                gen = (
                    'w_'
                    + str(random.randint(0, 255))
                    + str(random.randint(0, 255))
                    )
                while gen in idPool2:
                    gen = (
                        'w_'
                        + str(random.randint(0, 255))
                        + str(random.randint(0, 255))
                        )
                idPool2.add(gen)
                exec(gen + '= element')
                singleElements.append(gen)
            else:
                try:
                    complex(element)
                except Exception as e:
                    print(e, element)
                    return
                singleElements.append(element)
        commandExecutable = globals_.REGEXP2.sub('{}', commandExecutable)
        commandExecutable = commandExecutable.format(*singleElements)
        try:
            result = eval(commandExecutable)
        except Exception as e:
            print(e)
            return
        if not flag:
            if globals_.historyIndex != -1:
                hIndex = \
                    globals_.historyIndex + len(model.history) + 1
                model.history = model.history[:hIndex]
        domain = {}
        pF = model.formulas.get(
            (ridx[0], ridx[1]), None)
        pText = pF.text if pF else ''
        invert = False
        if type(result) is np.ndarray:
            rowIdx = ridx[0]
            colIdx = ridx[1]
            if com and result.ndim > 0:
                invert = True
                prev = model.dataContainer.get((rowIdx, colIdx), 'null')
                if prev != 'null':
                    if isinstance(prev, np.ndarray):
                        com = False
                        self.clean(
                            rowIdx,
                            colIdx,
                            1,
                            1)
                    else:
                        if len(result.shape) > 1:
                            x_cols = result.shape[1]
                        else:
                            x_cols = 1
                        self.clean(
                            rowIdx,
                            colIdx,
                            result.shape[0],
                            x_cols
                            )
            if com and result.ndim <= 2:
                domain['rowIdx'] = rowIdx
                domain['colIdx'] = colIdx
                domain['nRows'] = 1
                domain['nCols'] = 1
                if not flag and text != pText or invert:
                    try:
                        self.view.createFormula(
                            text,
                            coords,
                            singleIndexes,
                            domain
                            )
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                startIndex = model.index(rowIdx, colIdx)
                endIndex = startIndex
                model.setData(
                    model.index(rowIdx, colIdx),
                    result,
                    mode='a')
            elif len(result.shape) == 2:
                nRows = result.shape[0]
                nCols = result.shape[1]
                domain['rowIdx'] = rowIdx
                domain['colIdx'] = colIdx
                domain['nRows'] = nRows
                domain['nCols'] = nCols
                if not flag and text != pText or invert:
                    try:
                        self.view.createFormula(
                            text, coords, singleIndexes, domain
                            )
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                for line, rY in zip(
                        result,
                        range(rowIdx, rowIdx+nRows)):
                    for dE, cX in zip(
                            line,
                            range(
                                colIdx,
                                colIdx
                                + nCols)):
                        ind = model.createIndex(rY, cX)
                        model.setData(
                            ind,
                            globals_.currentFont,
                            role=Qt.FontRole
                            )
                        model.setData(
                            ind,
                            dE,
                            mode='a'
                            )
                startIndex = model.index(
                    rowIdx,
                    colIdx
                    )
                endIndex = model.index(
                    rowIdx+nRows-1,
                    colIdx+nCols-1)
            elif len(result.shape) == 1:
                nRows = result.shape[0]
                nCols = 1
                domain['rowIdx'] = rowIdx
                domain['colIdx'] = colIdx
                domain['nRows'] = nRows
                domain['nCols'] = nCols
                if not flag and text != pText or invert:
                    try:
                        self.view.createFormula(
                            text,
                            coords,
                            singleIndexes,
                            domain)
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                for row, ry in zip(
                        result,
                        range(rowIdx, rowIdx+nRows)):
                    ind = model.createIndex(ry, colIdx)
                    model.setData(
                        ind,
                        globals_.currentFont,
                        role=Qt.FontRole
                        )
                    model.setData(ind, row, mode='a')
                startIndex = model.index(
                    rowIdx,
                    colIdx
                    )
                endIndex = model.index(
                    rowIdx+nRows-1,
                    colIdx)
            else:
                return
            model.dataChanged.emit(startIndex, endIndex)
            self.commandLineEdit.clearFocus()
        elif isinstance(result, numbers.Number):
            rowIdx = ridx[0]
            colIdx = ridx[1]
            domain['rowIdx'] = rowIdx
            domain['colIdx'] = colIdx
            domain['nRows'] = 1
            domain['nCols'] = 1
            if not flag and text != pText:
                try:
                    self.view.createFormula(
                        text,
                        coords,
                        singleIndexes,
                        domain
                        )
                except Exception as e:
                    traceback.print_tb(e.__traceback__)
                    print(e)
                    return
            startIndex = model.createIndex(
                rowIdx,
                colIdx
                )
            endIndex = startIndex
            model.setData(
                startIndex,
                globals_.currentFont,
                role=Qt.FontRole
                )
            model.setData(
                startIndex,
                result,
                mode='a'
                )
            model.dataChanged.emit(startIndex, endIndex)
            self.commandLineEdit.clearFocus()
        currentFormula = model.formulas[
            ridx[0],
            ridx[1]
            ]
        if not flag:
            ordered = self.topologicalSort(currentFormula.precedence)
            self.executeOrder(ordered)
            model.ftoapply.clear()
            model.formulaSnap.clear()
            self.view.saveToHistory()
        else:
            self.view.saveToHistory()
        self.view.setFocus()

    def clean(self, x, y, rows, cols):
        """Delete data and formulas to prepare for new formula"""
        model = self.view.model()
        model.formulaSnap.update(model.formulas.values())
        for r in range(x, x + rows):
            for c in range(y, y + cols):
                model.setData(
                    model.index(r, c),
                    '',
                    mode='m',
                    erase='y',
                    )
        if rows > 1 or cols > 1:
            if model.ftoapply:
                ordered = self.topologicalSort(model.ftoapply)
                self.executeOrder(ordered)
        startIndex = model.index(x, y)
        endIndex = model.index(x + rows, y + cols)
        model.dataChanged.emit(startIndex, endIndex)

    def executeOrder(self, formulas):
        """Execute formulas in order"""
        for f in formulas:
            self.calculate(
                f.text,
                f.row,
                f.col,
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
    returnCommand = Signal(str, int, int, bool)

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
                compact = False
                if event.modifiers() == Qt.ControlModifier:
                    compact = True
                if self.text().startswith("="):
                    globals_.formula_mode = False
                    eval_ = self.text().partition("=")[2]
                    row = self.currentIndex.row()
                    column = self.currentIndex.column()
                    self.parent().parent().view.model().highlight = None
                    self.returnCommand.emit(eval_, row, column, compact)
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


class PlotWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        layout = QVBoxLayout(self)
        self.static_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        self.nav = NavigationToolbar(self.static_canvas, self)
        layout.addWidget(self.nav)
        layout.addWidget(self.static_canvas)
        self.ax = self.static_canvas.figure.subplots()

    def saveData(self, **vargs):
        for k, v in vargs.items():
            exec(f'self.{k} = v')


class PlotMenu(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.scene = QGraphicsScene()
        self.pView = PlotView(self.scene)
        self.conf = PlotConf(view=self.pView)
        mainSplitter = QSplitter()
        mainSplitter.addWidget(self.pView)
        mainSplitter.addWidget(self.conf)
        self.setCentralWidget(mainSplitter)
        mainSplitter.setStretchFactor(0, 1)
        mainSplitter.setStretchFactor(1, 0)
        self.setWindowTitle('Plot Menu')

    def closeEvent(self, event):
        super().closeEvent(event)
        self.parent().plotMenu = None


class PlotConf(QWidget):
    def __init__(self, parent=None, *, view):
        super().__init__(parent)
        self.view = view
        self.activeLE = None
        self.namesRecord = {}

        pW = QWidget()
        plotLayout = QGridLayout()
        xLabel = QLabel('x')
        self.xEdit = LinkedLineEdit()
        yLabel = QLabel('y')
        self.yEdit = LinkedLineEdit()
        self.addToCurrent = QPushButton('Add to current')

        self.xEdit.active.connect(
            self.setActiveLE
            )
        self.yEdit.active.connect(
            self.setActiveLE
            )
        self.addToCurrent.clicked.connect(self.validateInput)

        plotLayout.addWidget(xLabel, 0, 0)
        plotLayout.addWidget(self.xEdit, 0, 1)
        plotLayout.addWidget(yLabel, 1, 0)
        plotLayout.addWidget(self.yEdit, 1, 1)
        plotLayout.addWidget(self.addToCurrent, 2, 1)
        pW.setLayout(plotLayout)

        sW = QWidget()
        scatterLayout = QGridLayout()
        x2Label = QLabel('x')
        self.x2Edit = LinkedLineEdit()
        y2Label = QLabel('y')
        self.y2Edit = LinkedLineEdit()

        self.x2Edit.active.connect(
            self.setActiveLE
            )
        self.y2Edit.active.connect(
            self.setActiveLE
            )

        scatterLayout.addWidget(x2Label, 0, 0)
        scatterLayout.addWidget(self.x2Edit, 0, 1)
        scatterLayout.addWidget(y2Label, 1, 0)
        scatterLayout.addWidget(self.y2Edit, 1, 1)
        sW.setLayout(scatterLayout)

        bW = QWidget()
        barLayout = QGridLayout()
        x3Label = QLabel('x coords or labels:')
        self.x3Edit = LinkedLineEdit()
        y3Label = QLabel('values (bars height):')
        self.y3Edit = LinkedLineEdit()

        self.x3Edit.active.connect(
            self.setActiveLE
            )
        self.y3Edit.active.connect(
            self.setActiveLE
            )

        barLayout.addWidget(x3Label, 0, 0)
        barLayout.addWidget(self.x3Edit, 0, 1)
        barLayout.addWidget(y3Label, 1, 0)
        barLayout.addWidget(self.y3Edit, 1, 1)
        bW.setLayout(barLayout)

        hW = QWidget()
        histLayout = QGridLayout()
        x4Label = QLabel('x:')
        self.x4Edit = LinkedLineEdit()
        binsLabel = QLabel('bins:')
        self.binsEdit = LinkedLineEdit()
        rangeLabel = QLabel('range:')
        self.rangeEdit = LinkedLineEdit()
        densityLabel = QLabel('density:')
        self.density = QCheckBox()

        self.x4Edit.active.connect(
            self.setActiveLE
            )
        self.binsEdit.active.connect(
            self.setActiveLE
            )
        self.rangeEdit.active.connect(
            self.setActiveLE
            )

        histLayout.addWidget(x4Label, 0, 0)
        histLayout.addWidget(self.x4Edit, 0, 1)
        histLayout.addWidget(binsLabel, 1, 0)
        histLayout.addWidget(self.binsEdit, 1, 1)
        histLayout.addWidget(rangeLabel, 2, 0)
        histLayout.addWidget(self.rangeEdit, 2, 1)
        histLayout.addWidget(densityLabel, 3, 0)
        histLayout.addWidget(self.density, 3, 1)
        hW.setLayout(histLayout)

        pieW = QWidget()
        pieLayout = QGridLayout()
        x5Label = QLabel('x:')
        self.x5Edit = LinkedLineEdit()
        pLabel = QLabel('labels:')
        self.pieLEdit = LinkedLineEdit()

        self.x5Edit.active.connect(
            self.setActiveLE
            )
        self.pieLEdit.active.connect(
            self.setActiveLE)

        pieLayout.addWidget(x5Label, 0, 0)
        pieLayout.addWidget(self.x5Edit, 0, 1)
        pieLayout.addWidget(pLabel, 1, 0)
        pieLayout.addWidget(self.pieLEdit, 1, 1)
        pieW.setLayout(pieLayout)

        commonLayout = QGridLayout()
        gLabel = QLabel('Select graph:')
        self.graphSelect = QComboBox()
        tLabel = QLabel('Select type:')
        self.typeSelect = QComboBox()
        nLabel = QLabel('Graph name:')
        self.nEdit = QLineEdit()
        self.addBttn = QPushButton('New graph')
        self.rmBttn = QPushButton('Remove graph')
        self.stackW = QStackedWidget()

        self.graphSelect.addItem('--')
        self.graphSelect.currentIndexChanged.connect(
            self.updateInfo
            )
        self.typeSelect.addItems(globals_.GRAPH_TYPES)
        self.typeSelect.currentIndexChanged.connect(
            self.stackW.setCurrentIndex
            )
        self.addBttn.clicked.connect(self.validateInput)
        self.rmBttn.clicked.connect(self.removeGraph)
        self.stackW.addWidget(pW)
        self.stackW.addWidget(sW)
        self.stackW.addWidget(bW)
        self.stackW.addWidget(hW)
        self.stackW.addWidget(pieW)

        commonLayout.addWidget(gLabel, 0, 0)
        commonLayout.addWidget(self.graphSelect, 0, 1)
        commonLayout.addWidget(tLabel, 1, 0)
        commonLayout.addWidget(self.typeSelect, 1, 1)
        commonLayout.addWidget(self.stackW, 2, 0, 2, 2)
        commonLayout.addWidget(nLabel, 3, 0)
        commonLayout.addWidget(self.nEdit, 3, 1)
        commonLayout.addWidget(self.addBttn, 4, 0)
        commonLayout.addWidget(self.rmBttn, 4, 1)
        self.setLayout(commonLayout)

    def removeGraph(self):
        name = self.graphSelect.currentText()
        idx = self.graphSelect.currentIndex()
        location = self.namesRecord[name]
        graph = self.view.graphs[location]
        self.rearrangeItems(graph)
        self.view.scene().removeItem(graph)
        self.graphSelect.removeItem(idx)
        del self.namesRecord[name]
        self.view.c -= 1
        self.xEdit.clear()
        self.yEdit.clear()
        self.x2Edit.clear()
        self.y2Edit.clear()
        self.nEdit.clear()

    def rearrangeItems(self, g):
        if self.view.c == (self.view.side-1)**2:
            self.view.square = (self.view.side-1)**2
            self.view.side -= 1
        c = self.view.c - 1
        side = self.view.side
        square = self.view.square
        if c == square:
            side += 1
            square = side ** 2
        insert = square - c
        rows = insert // side
        if rows > 0 and insert > side:
            rows += (insert % side)
            y = side - rows
            x = side - 1
        else:
            y = side - 1
            x = side - insert
        last = self.view.graphs[y, x]
        if last == g:
            return
        pos = g.scenePos()
        last.setPos(pos)
        gCoord = self.namesRecord[g.widget().title]
        self.view.graphs[gCoord] = last
        del self.view.graphs[y, x]
        self.namesRecord[last.widget().title] = gCoord
