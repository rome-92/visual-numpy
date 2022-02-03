import sys
import os
import csv

import pytest
from PySide6.QtCore import Qt, QEvent, QItemSelectionModel
from PySide6.QtWidgets import QStyleOptionViewItem
from PySide6.QtGui import QKeyEvent

sys.path.append(os.path.dirname(__file__)+'/..')
from MyWidgets import MainWindow


dirname = os.path.dirname(__file__)


@pytest.fixture(scope='session')
def app(qapp):
    widget = MainWindow()
    return widget


@pytest.fixture(scope='function')
def loadF(app):
    yield app.loadFile(dirname + '/dependencies5.vnp')
    app.createNew()


@pytest.mark.parametrize(
    'index, expected', [
        ('A2', (1, 0)),
        ('B7', (6, 1)),
        ('AAC32', (31, 704)),
        ('ZD43', (42, 679))
        ]
    )
def test_getCoord(app, index, expected):
    assert app.getCoord(index) == expected


def test_loadFile(app, name='/dependencies5.vnp'):
    app.loadFile(dirname + name)
    model = app.view.model()
    assert model.formulas[5, 4].text == '[C6:C8]*2'
    assert model.dataContainer[5, 5] == 4 + 0j
    assert model.alignmentDict[5, 6] == int(
        Qt.AlignHCenter | Qt.AlignVCenter)
    assert model.fonts[4, 4].toString() == \
        'FreeMono,11,-1,5,400,0,0,0,0,0,0,0,0,0,0,1'
    assert model.foreground[5, 5].color().name() == '#204a87'
    assert model.background[4, 4].color().name() == '#e9b96e'


def test_importFile(app):
    app.importFile(dirname + '/csv_sample.csv')
    model = app.view.model()
    model.thousandsSep = False
    path = dirname + '/csv_sample.csv'
    with open(path, encoding='latin', newline='') as myFile:
        reader = csv.reader(myFile, dialect='excel')
        for y, row in enumerate(reader):
            for x, row in enumerate(row):
                try:
                    number = complex(row)
                    stored = complex(model.data(model.index(y, x)))
                    assert number == stored

                except ValueError:
                    assert row == model.data(model.index(y, x))


def test_createNew(app):
    app.loadFile(dirname + '/dependencies5.vnp')
    app.createNew()
    model = app.view.model()
    assert len(model.dataContainer) == 0
    assert len(model.formulas) == 0
    assert len(model.alignmentDict) == 0
    assert len(model.fonts) == 0
    assert len(model.foreground) == 0
    assert len(model.background) == 0


def test_fileExport(app, loadF):
    app.fileExport('export_test')
    model = app.view.model()
    try:
        with open('export_test.csv', encoding='latin', newline='') as myFile:
            reader = csv.reader(myFile, dialect='excel')
            for y, row in enumerate(reader):
                for x, e in enumerate(row):
                    index = model.index(y, x)
                    assert e == model.data(index)
    finally:
        os.remove('export_test.csv')


def test_input(app):
    index = app.view.model().index(0, 3)
    option = QStyleOptionViewItem()
    delegate = app.view.itemDelegateForIndex(index)
    editor = delegate.createEditor(app.view, option, index)
    editor.setText('test_input')
    delegate.setModelData(editor, app.view.model(), index)
    assert app.view.model().dataContainer[0, 3] == 'test_input'


def test_DeleteF(app, loadF):
    view = app.view
    model = view.model()
    sModel = view.selectionModel()
    sModel.select(
        view.model().index(5, 8),
        QItemSelectionModel.Select
        )
    dEvent = QKeyEvent(
        QEvent.KeyPress,
        Qt.Key_Delete,
        Qt.NoModifier
        )
    view.keyPressEvent(dEvent)
    formulas = model.formulas
    assert (5, 8) not in formulas
    assert model.data(model.index(5, 8)) == ''


def test_order(app, loadF):
    index = app.view.model().index(7, 2)
    option = QStyleOptionViewItem()
    delegate = app.view.itemDelegateForIndex(index)
    editor = delegate.createEditor(app, option, index)
    editor.setText('-2')
    delegate.setModelData(editor, app.view.model(), index)
    dataContainer = app.view.model().dataContainer
    assert dataContainer[7, 3] == -2
    assert dataContainer[7, 4] == -4
    assert dataContainer[7, 5] == 1
    assert dataContainer[5, 8] == -1
    assert dataContainer[5, 9] == 13
    assert dataContainer[5, 10] == pytest.approx(3.44444444)
    assert dataContainer[11, 4] == 6
    assert dataContainer[11, 5] == 20
    assert dataContainer[11, 6] == 29


def test_saveFileAs(app, loadF):
    app.saveFileAs('testSaveFile')
    try:
        test_loadFile(app, 'testSaveFile.vnp')
    finally:
        os.remove('testSaveFile.vnp')
