# visual-numpy
Visual Numpy is a basic WIP spreadsheet application intended to make use of the relatively simple python syntax and the numpy library as well by evaluating a python expression and returning it into the appropriate cell or cells. The name "Visual Numpy" comes from its capability of evaluating an expression that returns a numpy array and present it either as a column (for one dimensional arrays) or a series of contiguous columns (for two dimensional arrays) making it easy and intuitive to use arrays into an "excel like" interface.

Visual Numpy is a 100% python based application using the PySide6 library which is the official supported binding for the Qt library.

### Prerequisites

Get the prerequisites:
- Python (>=3.6)
- PySide6 (>=6.0)
- numpy (>=1.15)

## Installation
Currently no wheels or binaries are supported so you will need to clone the repository.

## Starting Visual Numpy
In order to start Visual Numpy directly from the cloned repository run

```
$ python visual_numpy.py
```

## Usage

To start typing a formula from any cell press "=". You can navigate through the cells and the corresponding cell reference will show up on the calculate bar, you can also select a rectangular array of cells, this cells will behave like numpy arrays (actually any contiguous selection) when evaluating the expression.![Visual Numpy (develop)](https://i.ibb.co/9V5NdSk/vnpy01.png)

As you will see usage of formulas and implementation differs from excel or any other "excel like" spreadsheet app on various ways, for instance to calculate a formula that references an array of cells only one formula is needed in contrast with excel in which a single formula is needed for every cell.

## Contributing
This is a WIP so any contribution is welcome. Currently documentation and bug fixes are number one priority.
For major changes, please open an issue first to discuss what you would like to change.

## License
[GPL v3](https://www.gnu.org/licenses/)
