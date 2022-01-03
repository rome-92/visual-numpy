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

import re

ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
POINT_SIZES = [
    '6', '7', '8', '9', '10', '10.5',
    '11', '12', '13', '14', '15', '16',
    '18', '20', '22', '24', '28', '32',
    '36', '40', '44', '48', '54', '60',
    '66', '72', '80', '88', '96'
    ]
currentFont = None
defaultFont = None
defaultForeground = None
defaultBackground = None
formula_mode = False
historyIndex = -1
drag = False
domainHighlight = False
REGEXP1 = re.compile(r'\[[A-Z]{1,3}[0-9]+:[A-Z]{1,3}[0-9]+]')
REGEXP2 = re.compile(r'[A-Z]{1,3}[0-9]+')
REGEXP3 = re.compile(r'[A-Z]{1,3}[0-9]+$')
REGEXP4 = re.compile(r'\[[A-Z]{1,3}[0-9]+:[A-Z]{1,3}[0-9]+]$')
LETTERS_REG_EXP = re.compile(r'[A-Z]+')
NUMBERS_REG_EXP = re.compile(r'[0-9]+')
