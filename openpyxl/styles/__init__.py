from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

from copy import deepcopy

from .alignment import Alignment
from .borders import Borders, Border
from .colors import Color
from .fills import Fill
from .fonts import Font
from .hashable import HashableObject
from .numbers import NumberFormat, is_date_format, is_builtin
from .protection import Protection


class Style(HashableObject):
    """Style object containing all formatting details."""
    __fields__ = ('font',
                  'fill',
                  'borders',
                  'alignment',
                  'number_format',
                  'protection')
    __slots__ = __fields__
    __check__ = {'font': Font,
                 'fill': Fill,
                 'borders': Borders,
                 'alignment': Alignment,
                 'number_format': NumberFormat,
                 'protection': Protection}

    def __init__(self, font=Font(), fill=Fill(), borders=Borders(),
                 alignment=Alignment(), number_format=NumberFormat(),
                 protection=Protection()):
        self.font = font
        self.fill = fill
        self.borders = borders
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection

DEFAULTS = Style()

