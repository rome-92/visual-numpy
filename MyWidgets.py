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

from PySide6.QtCore import QTimer, QSize, QEvent, Qt, Signal
from PySide6.QtWidgets import (QMainWindow, QLineEdit, QToolBar, QLabel, QFileDialog,
                               QMessageBox, QFontComboBox, QComboBox, QColorDialog,
                               QPushButton)
from PySide6.QtGui import (QAction, QGuiApplication, QActionGroup, QFontDatabase, QFont,
                           QPixmap,QIcon,QBrush,QColor)
from PySide6 import __version__ as PYSIDE6_VERSION
from PySide6.QtCore import __version__ as QT_VERSION
from MyView import MyView
from MyModel import MyModel
from MyDelegate import MyDelegate
import rcIcons
import globals_
import platform
import numbers
import traceback,random,csv,pickle
import numpy as np

version = '1.0.0-a.2'

class MainWindow(QMainWindow):
    
    styleSheet = '''
    MyView {selection-background-color: rgba(173, 255, 115,10%); selection-color: black;}'''
    currentFile = None

    def __init__(self,parent=None):
        '''Initialization of MainWindow widgets and actions'''
        super().__init__(parent)
        self.commandLineEdit = CommandLineEdit()
        self.view = MyView(self)
        self.view.setModel(MyModel(self.view))
        self.view.setItemDelegate(MyDelegate(self.view))
        self.setCentralWidget(self.view) 
        self.fontColor = QPushButton(self)
        self.fontColor.setObjectName('FontColor')
        self.fontColor.color_ = (240,240,240,255)
        self.fontColor.setIcon(QPixmap(':text_color.png'))
        self.fontColor.setIconSize(QSize(25,25))
        self.fontColor.setToolTip('Text color')
        self.backgroundColor = QPushButton(self)
        self.backgroundColor.setIcon(QPixmap(':background_color.png'))
        self.backgroundColor.setIconSize(QSize(25,25))
        self.backgroundColor.setToolTip('Background color')
        self.backgroundColor.setObjectName('BackgroundColor')
        self.backgroundColor.color_ = (240,240,240,255)
        self.fontColor.clicked.connect(self.showColorDialog)
        self.backgroundColor.clicked.connect(self.showColorDialog)
# Instantiate QActions
        newFile = QAction('&New File', self)
        newFile.setShortcut('Ctrl+N')
        newFile.setStatusTip('Create new file')
        newFile.triggered.connect(self.createNew)
        exportFile = QAction('&Export',self)
        exportFile.setShortcut('Ctrl+E')
        exportFile.setStatusTip('Export File to csv')
        exportFile.triggered.connect(self.fileExport)
        saveFile = QAction('&Save', self)
        saveFile.setShortcut('Ctrl+S')
        saveFile.setStatusTip('Save File')
        saveFile.triggered.connect(self.saveFile)
        saveFileAs = QAction('Save &As',self)
        saveFileAs.setShortcut('Shift+Ctrl+S')
        saveFileAs.setStatusTip('Save File As')
        saveFileAs.triggered.connect(self.saveFileAs)
        saveArrayAs = QAction('Save A&rray',self)
        saveArrayAs.setShortcut('Ctrl+J')
        saveArrayAs.setStatusTip('Save Array in .npy format')
        saveArrayAs.triggered.connect(self.saveArrayAs)
        self.view.addAction(saveArrayAs)
        loadFile = QAction('&Open',self)
        loadFile.setShortcut('Ctrl+O')
        loadFile.setStatusTip('Load .vnp file')
        loadFile.triggered.connect(self.loadFile)
        importFile = QAction('&Import',self)
        importFile.setShortcut('Ctrl+I')
        importFile.setStatusTip('Import File to csv')
        importFile.triggered.connect(self.importFile)
        thsndsSep = QAction('Thousands separator',self)
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
        self.boldAction = QAction('Bold',self)
        self.boldAction.setStatusTip('Bold font')
        self.boldAction.setIcon(QPixmap(':bold.png'))
        self.boldAction.setCheckable(True)
        self.boldAction.triggered.connect(self.bold)
        self.italicAction = QAction('Italic',self)
        self.italicAction.setStatusTip('Italicize font')
        self.italicAction.setIcon(QPixmap(':italic.png'))
        self.italicAction.setCheckable(True)
        self.italicAction.triggered.connect(self.italic)
        self.underlineAction = QAction('Underline',self)
        self.underlineAction.setStatusTip('Underline font')
        self.underlineAction.setIcon(QPixmap(':underline.png'))
        self.underlineAction.setCheckable(True)
        self.underlineAction.triggered.connect(self.underline)
# Create action group
        self.alignmentGroup1 = QActionGroup(self)
        self.alignmentGroup1.addAction(self.alignL)
        self.alignmentGroup1.addAction(self.alignC)
        self.alignmentGroup1.addAction(self.alignR)
        self.alignmentGroup2 = QActionGroup(self)
        self.alignmentGroup2.addAction(self.alignU)
        self.alignmentGroup2.addAction(self.alignM)
        self.alignmentGroup2.addAction(self.alignD)
        about = QAction('&About',self)
        about.setStatusTip('Show about information')
        about.triggered.connect(self.helpAbout)
# Instantiate menuBar
        mainMenu = self.menuBar()
# Add actions to menus
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(newFile)
        fileMenu.addAction(loadFile)
        fileMenu.addSeparator()
        fileMenu.addAction(saveFile)
        fileMenu.addAction(saveFileAs)
        fileMenu.addSeparator()
        fileMenu.addAction(exportFile)
        fileMenu.addAction(importFile)
        fileMenu.addSeparator()
        fileMenu.addAction(saveArrayAs)
        formatMenu = mainMenu.addMenu('For&mat')
        formatMenu.addAction(thsndsSep)
        helpMenu = self.menuBar().addMenu('&Help')
        helpMenu.addAction(about)
# End adding actions
        toolBar = QToolBar('Command Toolbar')
        toolBar.setAllowedAreas(Qt.TopToolBarArea)
        toolBar.setMovable(False)
        commandLabel = QLabel('Evaluate :') 
        self.commandLineEdit.installEventFilter(self.view)
        toolBar.addWidget(commandLabel)
        toolBar.addWidget(self.commandLineEdit)
        self.addToolBar(Qt.TopToolBarArea,toolBar)
        self.addToolBarBreak(Qt.TopToolBarArea)
# Second toolbar
        formatToolbar = QToolBar('Format Toolbar')
        formatToolbar.setAllowedAreas(Qt.TopToolBarArea)
        formatToolbar.setMovable(False)
        self.fontsComboBox = QFontComboBox()
        self.fontsComboBox.currentFontChanged.connect(self.updateFont)
        self.pointSize = QComboBox()
        self.pointSize.addItems(globals_.POINT_SIZES)
        self.pointSize.setMaxVisibleItems(10)
        self.pointSize.setStyleSheet('''QComboBox {combobox-popup: 0;}''')
        self.pointSize.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pointSize.setCurrentIndex(6)
        self.pointSize.currentTextChanged.connect(self.updateFont)
        self.fontsComboBox.activated.connect(self.saveFont2History)
        self.pointSize.activated.connect(self.saveFont2History)
        formatToolbar.addWidget(self.fontsComboBox)
        formatToolbar.addWidget(self.pointSize)
        formatToolbar.addWidget(self.fontColor)
        formatToolbar.addWidget(self.backgroundColor)
        formatToolbar.addActions([self.alignL,self.alignC,self.alignR,self.alignU,self.alignM,self.alignD,self.boldAction,self.italicAction,self.underlineAction])
        self.addToolBar(Qt.TopToolBarArea,formatToolbar)
        self.commandLineEdit.returnCommand.connect(self.calculate)
        self.statusBar()
        self.setStyleSheet(self.styleSheet)
        globals_.currentFont = QFont(self.fontsComboBox.currentFont())
        globals_.defaultFont = QFont(globals_.currentFont)
        globals_.defaultForeground = QBrush(Qt.black)
        globals_.defaultBackground = QBrush(Qt.white)
        QTimer.singleShot(0,self.center)
 
    def center(self):
        '''Centers MainwWindow on the screen'''
        qr = self.frameGeometry()
        cp = QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def sizeHint(self):
        '''Default initialization size'''
        return QSize(1280,720)

    def createNew(self):
        '''Creates a new file and clears history'''
        dataType = np.dtype('U32,D')
        self.view.model().dataContainer = np.zeros((52,52),dtype=dataType)
        self.view.model().formulas.clear()
        self.view.model().alignmentDict.clear()
        self.view.model().fonts.clear()
        self.viewl.model().foreground.clear()
        self.view.model().background.clear()
        self.view.model().history.clear()
        self.view.model().history.append((self.view.model().dataContainer.copy(),
                self.view.model().formulas.copy(),self.view.model().alignmentDict().copy(),
                self.view.model().fonts.copy(),self.view.model().foreground.copy(),
                self.view.model().background.copy()))

    def importFile(self):
        '''Imports csv files'''
        name,notUsed = QFileDialog.getOpenFileName(self,'Import File','','csv files (*.csv)')
        if name:
            try:
                with open(name,newline='') as myFile:
                    reader = csv.reader(myFile,dialect='excel')
                    dataType = np.dtype('U32,D')
                    self.view.model().dataContainer = np.zeros((52,52),dtype=dataType)
                    self.view.model().formulas.clear()
                    self.view.model().alignmentDict.clear()
                    self.view.model().fonts.clear()
                    self.view.model().foreground.clear()
                    self.view.model().background.clear()
                    self.view.model().history.clear()
                    self.view.model().history.append((self.view.model().dataContainer.copy(),
                            self.view.model().formulas.copy(),self.view.model().alignmentDict().copy(),
                            self.view.model().fonts.copy(),self.view.model().foreground.copy(),
                            self.view.model().background.copy()))
                    for rowNumber,row in enumerate(reader):
                        for columnNumber,column in enumerate(row):
                            index = self.view.model().createIndex(rowNumber,columnNumber)
                            self.view.model().setData(index,column)
                MainWindow.currentFile = name
                info = name+' was succesfully imported'
                self.statusBar().showMessage(info,5000)
            except Exception as e:
                print(e)
                info = 'There was an error importing '+name
                self.statusBar().showMessage(info,5000)

    def saveFile(self):
        '''Saves file into .vnp format'''
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
            nonzeroes = np.nonzero(model['f0'])
            rows = nonzeroes[0]
            columns = nonzeroes[1]
            maxrow = rows[rows.argmax()]
            maxcolumn = columns[columns.argmax()]
            finalModel = model[:maxrow+1, :maxcolumn+1]['f0']
            with open(name,'wb') as myFile:
                pickle.dump(finalModel, myFile)
                pickle.dump(self.view.model().formulas, myFile)
                pickle.dump(alignment, myFile)
                pickle.dump(fonts, myFile)
                pickle.dump(foreground, myFile)
                pickle.dump(background, myFile)
            info = name+' was succesfully saved'
            self.statusBar().showMessage(info,5000)
        else:
            self.saveFileAs()

    def saveFileAs(self):
        '''Saves file into .vnp format'''
        name, notUsed = QFileDialog.getSaveFileName(self,'Save File','','vnp files (*.vnp)')
        if name:
            model = self.view.model().dataContainer
            alignment = self.view.model().alignmentDict
            fonts = self.view.model().fonts.copy()
            foreground = self.view.model().foreground.copy()
            background = self.view.model().background.copy()
            self.encodeFonts(fonts)
            self.encodeColors(foreground)
            self.encodeColors(background)
            nonzeroes = np.nonzero(model['f0'])
            rows = nonzeroes[0]
            columns = nonzeroes[1]
            maxrow = rows[rows.argmax()]
            maxcolumn = columns[columns.argmax()]
            finalModel = model[:maxrow+1,:maxcolumn+1]['f0']
            with open(name+'.vnp','wb') as myFile:
                pickle.dump(finalModel,myFile)
                pickle.dump(self.view.model().formulas,myFile)
                pickle.dump(alignment, myFile)
                pickle.dump(fonts, myFile)
                pickle.dump(foreground, myFile)
                pickle.dump(background, myFile)
            info = name+ ' was succesfully saved'
            self.statusBar().showMessage(info,5000)
            MainWindow.currentFile = name

    def encodeFonts(self,fonts):
        for i,f in fonts.items():
            family = f.family()
            size = f.pointSizeF()
            if size.is_integer():
                size = int(size)
            bold = f.bold()
            italic = f.italic()
            underline = f.underline()
            fonts[i] = [family,size,bold,italic,underline]

    def encodeColors(self,brushes):
        for i,b in brushes.items():
            color = b.color().name()
            brushes[i] = color

    def decodeFonts(self,fonts):
        for i,f in fonts.items():
            font = QFont(f[0])
            font.setPointSizeF(f[1])
            font.setBold(f[2])
            font.setItalic(f[3])
            font.setUnderline(f[4])
            fonts[i] = font

    def decodeColors(self,colors):
        for i,c in colors.items():
            brush = QBrush(QColor(c))
            colors[i] = brush

    def saveArrayAs(self):
        '''Saves array into .npy array format'''
        name, notUsed = QFileDialog.getSaveFileName(self,'Save Array','','npy files(*.npy)')
        if name:
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
                model = self.view.model().dataContainer
                array = model[topRow:bottomRow+1,leftColumn:rightColumn+1]
                np.save(name,array)
                info = name+ ' was succesfully saved'
                self.statusBar().showMessage(info,5000)

    def loadFile(self):
        '''Loads .vnp format'''
        name,notUsed = QFileDialog.getOpenFileName(self,'Load File','','vnp files (*.vnp)')
        if name:
            self.view.model().formulas = []
            try:
                with open(name,'rb') as myFile:
                    loadedModel = pickle.load(myFile)
                    formulas = pickle.load(myFile)
                    alignment = pickle.load(myFile)
                    fonts = pickle.load(myFile)
                    foreground = pickle.load(myFile)
                    background = pickle.load(myFile)
                    self.decodeFonts(fonts)
                    self.decodeColors(foreground)
                    self.decodeColors(background)
                    dataType = np.dtype('U32,D')
                    self.view.model().dataContainer = np.zeros((52,52),dtype=dataType)
                    self.view.model().alignmentDict = alignment
                    self.view.model().fonts = fonts
                    self.view.model().foreground = foreground
                    self.view.model().background = background
                    self.view.model().formulas.clear()
                    self.view.model().history.clear()
                    for rowNumber,row in enumerate(loadedModel):
                        for columnNumber,column in enumerate(row):
                            index = self.view.model().createIndex(rowNumber,columnNumber)
                            self.view.model().setData(index,column)
                    self.view.model().formulas = formulas
                    self.view.model().history.append((self.view.model().dataContainer.copy(),
                            self.view.model().formulas.copy(),self.view.model().alignmentDict().copy(),
                            self.view.model().fonts.copy(),self.view.model().foreground.copy(),
                            self.view.model().background.copy()))
                MainWindow.currentFile = name
                info = name+ ' was succesfully loaded'
                self.statusBar().showMessage(info,5000)
            except Exception as e:
                print(e)
                info = 'There was an error loading '+name
                self.statusBar().showMessage(info,5000)
        
    def fileExport(self):
        '''Exports file into .csv format'''
        name, notUsed = QFileDialog.getSaveFileName(self,'Save File','','csv file (*.csv)')
        if name:
            model = self.view.model().dataContainer
            nonzeroes = np.nonzero(model['f0'])
            rows = nonzeroes[0]
            columns = nonzeroes[1]
            maxrow = rows[rows.argmax()]
            maxcolumn = columns[columns.argmax()]
            finalModel = model[:maxrow+1,:maxcolumn+1]['f0']
            np.savetxt(name,finalModel,delimiter=',',fmt='%-1s')
            info = name+ ' was succesfully exported'
            self.statusBar().showMessage(info,5000)

    def setThousandsSep(self):
        '''Simple enable/disable thousands separator'''
        if self.sender().isChecked():
            self.view.model().enableThousandsSep()
        else:
            self.view.model().disableThousandsSep()

    def alignLeft(self):
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
            model.alignmentDict[i.row(),i.column()] = int(Qt.AlignLeft|vertical)
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()


    def alignCenter(self):
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
            model.alignmentDict[i.row(),i.column()] = int(Qt.AlignHCenter|vertical) 
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def alignRight(self):
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
            model.alignmentDict[i.row(),i.column()] = int(Qt.AlignRight|vertical) 
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def alignUp(self):
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
            model.alignmentDict[i.row(),i.column()] = int(horizontal|Qt.AlignTop) 
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()
        

    def alignMiddle(self):
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
            model.alignmentDict[i.row(),i.column()] = int(horizontal|Qt.AlignVCenter) 
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()
        

    def alignDown(self):
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
            model.alignmentDict[i.row(),i.column()] = int(horizontal|Qt.AlignBottom) 
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def saveFont2History(self):
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        self.view.saveToHistory()

    def updateFont(self,varg=None):
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        if type(varg) == QFont:
            globals_.currentFont = QFont(varg)
            pointSize = float(self.pointSize.currentText())
            globals_.currentFont.setPointSizeF(pointSize)
        elif type(varg) == str:
            font = self.fontsComboBox.currentFont()
            globals_.currentFont = QFont(font)
            pointSize = float(varg)
            globals_.currentFont.setPointSizeF(pointSize)
        for index in selected:
            self.view.model().setData(index,globals_.currentFont,role=Qt.FontRole)
        self.view.model().dataChanged.emit(selected[0],selected[-1])

    def bold(self):
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(model.fonts.get((index.row(),index.column()),globals_.defaultFont))
            if self.sender().isChecked():
                font.setBold(True)
            else:
                font.setBold(False)
            model.setData(index,font,role=Qt.FontRole)
        model.dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()
    
    def italic(self):
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(model.fonts.get((index.row(),index.column()),globals_.defaultFont))
            if self.sender().isChecked():
                font.setItalic(True)
            else:
                font.setItalic(False)
            model.setData(index,font,role=Qt.FontRole)
        model.dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def underline(self):
        selectionModel = self.view.selectionModel()
        selected = selectionModel.selectedIndexes()
        if not selected:
            return
        model = self.view.model()
        for index in selected:
            font = QFont(model.fonts.get((index.row(),index.column()),globals_.defaultFont))
            if self.sender().isChecked():
                font.setUnderline(True)
            else:
                font.setUnderline(False)
            model.setData(index,font,role=Qt.FontRole)
        model.dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def showColorDialog(self):
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
        styleSheet = ''' 
        MyView {selection-background-color: rgba(173, 255, 115,25); selection-color: black;}
        QPushButton#FontColor {background-color: rgba(%s, %s, %s, %s);}
        QPushButton#BackgroundColor {background-color: rgba(%s, %s, %s, %s);}
        '''%(fontR,fontG,fontB,fontA,backgroundR,backgroundG,backgroundB,backgroundA)
        self.setStyleSheet(styleSheet)
        model = self.view.model()
        if not selected:
            return
        for index in selected:
            if self.sender().objectName() == 'FontColor':
                model.setData(index,QBrush(color),role=Qt.ForegroundRole)
            else:
                model.setData(index,QBrush(color),role=Qt.BackgroundRole)
        self.view.model().dataChanged.emit(selected[0],selected[-1])
        self.view.saveToHistory()

    def helpAbout(self):
        QMessageBox.about(self, "About Visual Numpy",
                """<b>Visual Numpy</b> version {0}
                <p>Copyright &copy; 2021 Román U. Martínez 
                <p>Simple spreadsheet application "numpy enabled".
                <p>License: GNU General Public License v3. 
                <p>Python {1} - Qt {2} - PySide {3} on {4}""".format(
                version, platform.python_version(),
                QT_VERSION, PYSIDE6_VERSION,
                platform.system()))

    def getCoord(self,index):
        '''Gets row,column (y,x) coordinates from alphanumeric coord'''
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
        return row,column

    def calculate(self,text,*resultIndex,flag=False):
        '''Formats corresponding string into python executable code and evaluates resulting expression'''
        #locates array selection via reg exp
        arrays = globals_.REGEXP1.findall(text)       
        #stores the 2 element tuples cont. the alphanumeric coords
        indexes = []
        #stores the 2 element tuples cont. the numeric coords
        coords = []         
        for array in arrays:
            topLeft,bottomRight = globals_.REGEXP2.findall(array)
            indexes.append((topLeft,bottomRight))
        for index in indexes:
            limit1 = self.getCoord(index[0])
            limit2 = self.getCoord(index[1])
            coords.append((limit1,limit2))
        #stores the actual numpy arrays
        numpyArrayList = []                                                 
        for limits in coords:
            #gets the rows for start and end coords of array
            rows = [limits[0][0],limits[1][0]+1]                            
            columns = [limits[0][1],limits[1][1]+1]
            data = self.view.model().dataContainer[rows[0]:rows[1],columns[0]:columns[1]]['f0']
            for cell in data.flat:
                if cell != '':
                    try:
                        complex(cell)
                    except Exception as e:
                        print(e)
                        return
            #gets the complex value of the data model array and appends it to numpyArraylList
            data = self.view.model().dataContainer[rows[0]:rows[1],columns[0]:columns[1]]['f1']
            numpyArrayList.append(data)                                                             
        #list to store names for each of the arrays
        idPool = set()
        #condition to match the quantity of names to the same as arrays
        while len(idPool) != len(numpyArrayList):
            idPool.add('x_'+str(random.randint(0,255))+str(random.randint(0,255)))
        #prepares the command in order to format it                                                    
        commandExecutable = globals_.REGEXP1.sub('{}',text)
        #matches a name with an array and iterates through it
        for i1,i2 in zip(idPool,numpyArrayList):
            #assigns the actual name to the array
            exec(i1+'= i2')
        #substitutes the alphanumeric reference of the array to the actual name of the array
        commandExecutable = commandExecutable.format(*idPool)
        #gets the scalar references i.e. the single cell references
        scalars = globals_.REGEXP2.findall(commandExecutable)
        #stores the numeric coordinate of the scalar
        scalarsIndexes = []
        for scalar in scalars:
            #adds the coordinate (a tuple containing the row and column coords)
            scalarsIndexes.append(self.getCoord(scalar))
        #stores the string representation of the complex numbers                                            
        scalarNumbers = []                                  
        for i in scalarsIndexes:
            row = i[0]
            column = i[1]
            #gets the string representation of the assumed complex number
            number = self.view.model().dataContainer[row,column]['f0'] 
            if number != '':
                #tries to convert to a complex number
                try:
                    complex(number)
                #if there's an error the value is assumed to be a NaN
                except Exception as e:
                    #the appropiate error message shows up and returns                                   
                    print(e,number)                                                                     
                    return
            else:
                number = '0'
            #if there's no error the number is appended to scalarNumbers list
            scalarNumbers.append(number)
        #the command string is prepared by subsitituing first                                                           
        commandExecutable = globals_.REGEXP2.sub('{}',commandExecutable)
        #the actual scalars are substitued for the {}                                    
        commandExecutable = commandExecutable.format(*scalarNumbers)
        try:
            result = eval(commandExecutable)
        except Exception as e:
            #the string is evaluated and pointed by result
            print(e)    
            return
        #update the history model length
        if not flag:
            if globals_.historyIndex != -1:
                hIndex = globals_.historyIndex + len(self.view.model().history) + 1
                self.view.model().history = self.view.model().history[:hIndex]
        domainDict = {}
        if type(result) is np.ndarray:
            resultIndexRow = resultIndex[0]
            resultIndexColumn = resultIndex[1]
            self.view.model().formulaSnap = self.view.model().formulas.copy()
            if len(result.shape) == 2:
                #gets the shape of the result
                resultRows = result.shape[0]
                resultColumns = result.shape[1] 
                domainDict['resultIndexRow'] = resultIndexRow
                domainDict['resultIndexColumn'] = resultIndexColumn
                domainDict['resultRows'] = resultRows
                domainDict['resultColumns'] = resultColumns
                if not flag:
                    try:
                        self.view.createFormula(text,coords,scalarsIndexes,domainDict)
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                for line,rY in zip(result,range(resultIndexRow,resultIndexRow+resultRows)):
                    for dE,cX in zip(line,range(resultIndexColumn,resultIndexColumn+resultColumns)):
                        ind =self.view.model().createIndex(ry,cx)
                        self.view.model().setData(ind,globals_.currentFont,role=Qt.FontRole)
                        self.view.model().setData(ind,dE,formulaTriggered=True)
                startIndex = self.view.model().index(resultIndexRow,resultIndexColumn)
                endIndex = self.view.model().index(resultIndexRow+resultRows-1,
                        resultIndexColumn+resultColumns-1)
            elif len(result.shape) == 1:
                #resultRows = 1
                #resultColumns = result.shape[0]
                resultRows = result.shape[0]
                resultColumns = 1
                domainDict['resultIndexRow'] = resultIndexRow
                domainDict['resultIndexColumn'] = resultIndexColumn
                domainDict['resultRows'] = resultRows
                domainDict['resultColumns'] = resultColumns
                if not flag:
                    try:
                        self.view.createFormula(text,coords,scalarsIndexes,domainDict)
                    except Exception as e:
                        traceback.print_tb(e.__traceback__)
                        print(e)
                        return
                #for line,cX in zip(result,range(resultIndexColumn,resultIndexColumn+resultColumns)):
                    #self.view.model().setData(self.view.model().createIndex(resultIndexRow,cX),line,formulaTriggered=True)
                for row,ry in zip(result,range(resultIndexRow,resultIndexRow+resultRows)):
                    ind = self.view.model().createIndex(ry,resultIndexColumn)
                    self.view.model().setData(ind,globals_.currentFont,role=Qt.FontRole)
                    self.view.model().setData(ind,row,formulaTriggered=True)
                startIndex = self.view.model().index(resultIndexRow,resultIndexColumn)
                #endIndex = self.view.model().index(resultIndexRow,
                        #resultIndexColumn+resultColumns-1)
                endIndex = self.view.model().index(resultIndexRow+resultRows-1,resultIndexColumn)
            else:
                return
            #emits the necessary signal in order to update the view
            self.view.model().dataChanged.emit(startIndex,endIndex)
            #resets the command line edit and clears focus
            self.commandLineEdit.clearFocus()
        #if the result is not an array
        elif isinstance(result, numbers.Number):
            resultIndexRow = resultIndex[0]
            resultIndexColumn = resultIndex[1]
            domainDict['resultIndexRow'] = resultIndexRow
            domainDict['resultIndexColumn'] = resultIndexColumn
            domainDict['resultRows'] = 1
            domainDict['resultColumns'] = 1
            if not flag:
                try:
                    self.view.createFormula(text,coords,scalarsIndexes,domainDict)
                except Exception as e:
                    traceback.print_tb(e.__traceback__)
                    print(e)
                    return
            startIndex = self.view.model().createIndex(resultIndexRow,resultIndexColumn)
            endIndex = startIndex
            self.view.model().formulaSnap = self.view.model().formulas.copy()
            self.view.model().setData(startIndex,globals_.currentFont,role=Qt.FontRole)
            self.view.model().setData(startIndex,result,formulaTriggered=True)
            self.view.model().dataChanged.emit(startIndex,endIndex)
            self.commandLineEdit.clearFocus()
        if not flag:
            if self.view.model().ftoapply:
                self.view.model().updateModel_()
        self.view.setFocus()
        self.view.saveToHistory()



class CommandLineEdit(QLineEdit):
    '''Handles expression and emits corresponding signals'''
    returnCommand = Signal(str,int,int)
    def event(self,event):
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                if globals_.formula_mode:
                    self.clearFocus()
                    globals_.formula_mode = False
                    self.parent().parent().view.setFocus(Qt.ShortcutFocusReason)
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
                    self.returnCommand.emit(eval_,row,column)
                    return True
                else:
                    return super().event(event)
            else:
                return super().event(event)
        else:
            return super().event(event)

    def focusInEvent(self,event):
        super().focusInEvent(event)
        globals_.formula_mode = True
        
    def focusOutEvent(self,event):
        super().focusOutEvent(event)
        globals_formula_mode = False
