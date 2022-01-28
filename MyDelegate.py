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

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QStyledItemDelegate, QLineEdit
from PySide6.QtGui import QColor, QBrush

import globals_


class MyDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        view = self.parent()
        lineEdit = MyLineEdit(view, parent)
        return lineEdit

    def setEditorData(self, editor, index):
        editor.setText(index.model().data(index, role=Qt.DisplayRole))

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

    def setModelData(self, editor, model, index):
        global historyIndex
        if globals_.historyIndex != - 1:
            hIndex = \
                globals_.historyIndex + len(self.parent().model().history) + 1
            self.parent().model().history = \
                self.parent().model().history[:hIndex]
        if editor.text() == model.data(index):
            return
        font = model.fonts.get(
            (index.row(), index.column()),
            globals_.defaultFont
            )
        textColor = model.foreground.get(
            (index.row(), index.column()),
            globals_.defaultForeground
            )
        backColor = model.background.get(
            (index.row(), index.column()),
            globals_.defaultBackground
            )
        model.setData(index, font, role=Qt.FontRole)
        model.setData(index, editor.text(), mode='s')
        model.setData(index, textColor, role=Qt.ForegroundRole)
        model.setData(index, backColor, role=Qt.BackgroundRole)
        for f in model.formulas.values():
            if (index.row(), index.column()) in f.domain:
                model.setData(
                    index,
                    QBrush(QColor(255, 69, 69)),
                    role=Qt.BackgroundRole
                    )
                break
        self.parent().saveToHistory()


class MyLineEdit(QLineEdit):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.installEventFilter(view)
