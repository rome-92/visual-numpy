import pytest
from MyWidgets import MainWindow
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QStyleOptionViewItem


@pytest.fixture(scope='session')
def app(qapp):
    widget = MainWindow()
    return widget


@pytest.fixture(scope='function')
def loadF(app):
    yield app.loadFile('dependencies5.vnp')
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


def test_loadFile(app):
    app.loadFile('dependencies5.vnp')
    model = app.view.model()
    assert model.formulas[5, 4].text == '[C6:C8]*2'
    assert model.dataContainer[5, 5] == 4 + 0j
    assert model.alignmentDict[5, 6] == int(
        Qt.AlignHCenter | Qt.AlignVCenter)
    assert model.fonts[4, 4].toString() == \
        'FreeMono,11,-1,5,400,0,0,0,0,0,0,0,0,0,0,1'
    assert model.foreground[5, 5].color().name() == '#204a87'
    assert model.background[4, 4].color().name() == '#e9b96e'


def test_createNew(app):
    app.loadFile('dependencies5.vnp')
    app.createNew()
    model = app.view.model()
    assert len(model.dataContainer) == 0
    assert len(model.formulas) == 0
    assert len(model.alignmentDict) == 0
    assert len(model.fonts) == 0
    assert len(model.foreground) == 0
    assert len(model.background) == 0


def test_input(app):
    index = app.view.model().index(0, 3)
    option = QStyleOptionViewItem()
    delegate = app.view.itemDelegateForIndex(index)
    editor = delegate.createEditor(app.view, option, index)
    editor.setText('test_input')
    delegate.setModelData(editor, app.view.model(), index)
    assert app.view.model().dataContainer[0, 3] == 'test_input'


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
