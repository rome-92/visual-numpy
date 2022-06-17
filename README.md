# visual-numpy
Visual Numpy is a spreadsheet application under constant development intended to make use of the relatively simple python syntax and the numpy library as well by evaluating a python expression and returning it into the appropriate cell or cells. The name "Visual Numpy" comes from its capability of evaluating an expression that returns a numpy array and present it either as a column (for one dimensional arrays) or a series of contiguous columns (for two dimensional arrays) making it easy and intuitive to use arrays into an "excel like" interface.

Visual Numpy is a 100% python based application using the PySide6 library which is the official supported binding for the Qt library.

### Prerequisites

Get the prerequisites:
- Python (>=3.8)
- PySide6 (>=6.0)
- numpy (>=1.15)

Optional:

- matplotlib (>=3.5.1) Provides plot functionality (early development)

Note: If using Ubuntu other than LTS 20.04 installing pyside6 via pip could resort to some needed libraries missing. In this case be sure to install all missing libraries separately.

## Installation
Currently no wheels or binaries are supported so you will need to clone the repository.

## Starting Visual Numpy
In order to start Visual Numpy directly from the cloned repository run

```
$ python visual_numpy.py
```

## Usage

To start typing a formula from any cell press "=". You can navigate through the cells with the keyboard or mouse and the corresponding cell reference will show up on the calculate bar just like any other spreadsheet, you can also select a rectangular array of cells, this cells will behave like numpy arrays (actually any contiguous selection) when evaluating the expression.![Visual Numpy (develop)](https://i.ibb.co/9V5NdSk/vnpy01.png)

The evaluation of the expression works as follows, first any string that matches the regex of a cell array (eg. [A2:F10]) will be represented internally as a numpy array object, then any remaining cell references (eg. B2) will just be represented internally by the value the cell "holds". So for example a formula to calculate the sum of an array would consist of the simple expression: "=[B2:B9].sum()". As for now any valid python expression that returns either an ndarray up to two dimensions or another numeric object will be displayed in order begining on the cell that started the formula. Whenever a value changes that is part of a formula the affected formula(s) are recalculated just as expected.

If you need to make an explicit reference to the numpy module just use the prefix "np". So for creating an array with the range of numbers from 1 to 10 just typing "=np.arange(1,11)" will work as expected.

As you will see usage of formulas and implementation differs from excel or any other "excel like" spreadsheet app on various ways, for instance to calculate a formula that references an array of cells only one formula is needed in contrast with excel in which a single formula is needed for every cell.

*Update*: If you are working with very large arrays or would like to store the array in a cell and not having it "expanded" you can invoque its "compact" version by holding **CTRL** and pressing **ENTER**  effectively storing the whole array in a cell. This array will show up in the view with the words array and its shape in parenthesis (e.g array (10,1)). This can be done at formula creation time inside the evaluate bar or after it is created by selecting the cell that holds the formula and pressing **CTRL-ENTER**. Inversely if the array is already in its compact form pressing the same key combination will result in the array in its "expanded" form. This compact form has the added benefit of increasing performance when working with very large arrays.

### Plotting

Matplotlib is requiered for plotting capabilities. There are currently two ways of plotting. The first one being called "fast plot" is accessible via context menu (right click on view), before selecting "Fast plot" a pair of arrays should be selected (that is two array objects or two contiguous columns which will be the x and y arrays respectively), this will create a simple x-y plot showing the result in a separate window. The second option is accesible via the Plot menu. Selecting this will open the Plot Window. In this window several types of plots are accesible. Currently **plot, scatter, bar, histogram and pie** plot types are available, from here you need to select the corresponding arrays to plot with the mouse, the references will appear on the selected edit lines so you must first click inside the edit line and then make the selection with the mouse over the data for the corresponding element. After all the corresponding elements have been selected press "New graph" this will generate the graph of your choice with the data selected and appear on the left side. You can zoom in and pan with the mouse wheel by rolling and clicking respectively. Also if you select a graph with the combo box the selected graph will be centered on the view. Continious creation of new graphs will be arranged in a square shape on the view. Currently graphs do not update automatically so any change on the data will not be reflected on the graphs. Automatic update of graphs is being worked on and should be soon available on the develop branch first. To delete a graph from the view select in with the combo box and then press "Remove graph" this will effectively remove the graph from the view and rearrange the other graphs if necessary. ![Plotting in Visual Numpy](https://i.ibb.co/chrqCkF/plot.png)

## Contributing
This is a WIP so any contribution is welcome. Currently documentation and bug fixes are number one priority.
If you plan on actively contributing to the project forking the repo and pushing your patches on their own branches is the best option after which you can send me a pull request either through here or via email with the help of ``` $ git request-pull ```

In the case you are not planning on actively contributing but nevertheless want to help every now and then please send me an email with your patch with the use of ```$ git format-patch```
## License
[GPL v3](https://www.gnu.org/licenses/)
