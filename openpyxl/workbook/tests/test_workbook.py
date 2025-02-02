# Copyright (c) 2010-2024 openpyxl

import datetime

# package imports
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils.exceptions import ReadOnlyWorkbookException
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import TableList

from openpyxl.xml.constants import (
    XLSM,
    XLSX,
    XLTM,
    XLTX
)

import pytest

@pytest.fixture
def Workbook():
    """Workbook Class"""
    from openpyxl import Workbook
    return Workbook


@pytest.fixture
def Table():
    """Table Class"""
    from openpyxl.worksheet.table import Table
    return Table


class TestWorkbook:

    @pytest.mark.parametrize("has_vba, as_template, content_type",
                             [
                                 (None, False, XLSX),
                                 (None, True, XLTX),
                                 (True, False, XLSM),
                                 (True, True, XLTM)
                             ]
                             )
    def test_template(self, has_vba, as_template, content_type, Workbook):
        wb = Workbook()
        wb.vba_archive = has_vba
        wb.template = as_template
        assert wb.mime_type == content_type


    def test_named_styles(self, Workbook):
        wb = Workbook()
        assert wb.named_styles == ['Normal']


    def test_immutable_builtins(self, Workbook):
        wb1 = Workbook()
        wb2 = Workbook()
        normal = wb1._named_styles['Normal']
        normal.font.color = "FF0000"
        assert wb2._named_styles['Normal'].font.color.index == 1


    def test_duplicate_table_name(self, Workbook, Table):
        wb = Workbook()
        ws = wb.create_sheet()
        ws.add_table(Table(displayName="Table1", ref="A1:D10"))
        assert True == wb._duplicate_name("Table1")
        assert True == wb._duplicate_name("TABLE1")


    def test_duplicate_defined_name(self, Workbook):
        wb1 = Workbook()
        wb1.defined_names["dfn1"] = DefinedName("dfn1")
        assert True == wb1._duplicate_name("dfn1")
        assert True == wb1._duplicate_name("DFN1")

def test_get_active_sheet(Workbook):
    wb = Workbook()
    assert wb.active == wb.worksheets[0]


def test_set_active_by_sheet(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2',]
    for n in names:
        wb.create_sheet(n)

    for n in names:
        sheet = wb[n]
        wb.active = sheet
        assert wb.active == wb[n]


def test_set_active_by_index(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2',]
    for n in names:
        wb.create_sheet(n)

    for idx, name in enumerate(names):
        wb.active = idx
        assert wb.active == wb.worksheets[idx]


@pytest.mark.xfail
def test_set_invalid_active_index(Workbook):
    wb = Workbook()
    with pytest.raises(ValueError):
        wb.active = 1


def test_set_invalid_sheet_by_name(Workbook):
    wb = Workbook()
    with pytest.raises(TypeError):
        wb.active = "Sheet"


def test_set_invalid_child_as_active(Workbook):
    wb1 = Workbook()
    wb2 = Workbook()
    ws2 = wb2['Sheet']
    with pytest.raises(ValueError):
        wb1.active = ws2


def test_set_hidden_sheet_as_active(Workbook):
    wb = Workbook()
    ws = wb.create_sheet()
    ws.sheet_state = 'hidden'
    with pytest.raises(ValueError):
        wb.active = ws


def test_no_active(Workbook):
    wb = Workbook(write_only=True)
    assert wb.active is None


def test_create_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    assert new_sheet == wb.worksheets[-1]

def test_create_sheet_with_name(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet(title='LikeThisName')
    assert new_sheet == wb.worksheets[-1]

def test_add_correct_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    wb._add_sheet(new_sheet)
    assert new_sheet == wb.worksheets[2]

def test_add_sheetname(Workbook):
    wb = Workbook()
    with pytest.raises(TypeError):
        wb._add_sheet("Test")


def test_add_sheet_from_other_workbook(Workbook):
    wb1 = Workbook()
    wb2 = Workbook()
    ws = wb1.active
    with pytest.raises(ValueError):
        wb2._add_sheet(ws)


def test_create_sheet_readonly(Workbook):
    wb = Workbook()
    wb._read_only = True
    with pytest.raises(ReadOnlyWorkbookException):
        wb.create_sheet()


def test_remove_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    wb.remove(new_sheet)
    assert new_sheet not in wb.worksheets


def test_move_sheet(Workbook):
    wb = Workbook()
    for i in range(9):
        wb.create_sheet()
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4',
                            'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8', 'Sheet9']
    ws = wb['Sheet9']
    wb.move_sheet(ws, -5)
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet9',
                            'Sheet4', 'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8']


def test_move_sheet(Workbook):
    wb = Workbook()
    for i in range(9):
        wb.create_sheet()
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4',
                            'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8', 'Sheet9']
    wb.move_sheet("Sheet9", -5)
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet9',
                            'Sheet4', 'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8']

def test_getitem(Workbook):
    wb = Workbook()
    ws = wb['Sheet']
    assert isinstance(ws, Worksheet)
    with pytest.raises(KeyError):
        wb['NotThere']


def test_get_chartsheet(Workbook):
    wb = Workbook()
    cs = wb.create_chartsheet()
    assert wb[cs.title] is cs


def test_del_worksheet(Workbook):
    wb = Workbook()
    del wb['Sheet']
    assert wb.worksheets == []


def test_del_chartsheet(Workbook):
    wb = Workbook()
    cs = wb.create_chartsheet()
    del wb[cs.title]
    assert wb.chartsheets == []


def test_contains(Workbook):
    wb = Workbook()
    assert "Sheet" in wb
    assert "NotThere" not in wb

def test_iter(Workbook):
    wb = Workbook()
    for ws in wb:
        pass
    assert ws.title == "Sheet"

def test_index(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    sheet_index = wb.index(new_sheet)
    assert sheet_index == 1


def test_get_sheet_names(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
    for count in range(5):
        wb.create_sheet(0)
    assert wb.sheetnames == names


def test_add_invalid_worksheet_class_instance(Workbook):

    class AlternativeWorksheet:
        def __init__(self, parent_workbook, title=None):
            self.parent_workbook = parent_workbook
            if not title:
                title = 'AlternativeSheet'
            self.title = title

    wb = Workbook
    ws = AlternativeWorksheet(parent_workbook=wb)
    with pytest.raises(TypeError):
        wb._add_sheet(worksheet=ws)


class TestCopy:


    def test_worksheet_copy(self, Workbook):
        wb = Workbook()
        ws1 = wb.active
        ws2 = wb.copy_worksheet(ws1)
        assert ws2 is not None


    @pytest.mark.parametrize("title, copy",
                             [
                                 ("TestSheet", "TestSheet Copy"),
                                 (u"D\xfcsseldorf", u"D\xfcsseldorf Copy")
                                 ]
                             )
    def test_worksheet_copy_name(self, title, copy, Workbook):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = title
        ws2 = wb.copy_worksheet(ws1)
        assert ws2.title == copy


    def test_cannot_copy_readonly(self, Workbook):
        wb = Workbook()
        ws = wb.active
        wb._read_only = True
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)


    def test_cannot_copy_writeonly(self, Workbook):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet()
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)


    def test_default_epoch(self, Workbook):
        wb = Workbook()
        assert wb.epoch == datetime.datetime(1899, 12, 30)


    def test_assign_epoch(self, Workbook):
        wb = Workbook()
        wb.epoch = datetime.datetime(1904, 1, 1)


    def test_invalid_epoch(self, Workbook):
        wb = Workbook()
        with pytest.raises(ValueError):
            wb.epoch = datetime.datetime(1970, 1, 1)

