# Copyright (c) 2010-2024 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import xmlfile

from openpyxl.tests.helper import compare_xml
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

from openpyxl import LXML

@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def etree_write_cell():
    from .._writer import etree_write_cell
    return etree_write_cell


@pytest.fixture
def lxml_write_cell():
    from .._writer import lxml_write_cell
    return lxml_write_cell


@pytest.fixture(params=['etree', 'lxml'])
def write_cell_implementation(request, etree_write_cell, lxml_write_cell):
    if request.param == "lxml" and LXML:
        return lxml_write_cell
    return etree_write_cell


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="inlineStr" r="A1"><is><t>Hello</t></is></c>"""),
                             ("", """<c r="A1" t="inlineStr"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                         ])
def test_write_cell(worksheet, write_cell_implementation, value, expected):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, iso_dates, expected,",
                         [
                             (datetime.date(2011, 12, 25), False, """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                             (datetime.date(2011, 12, 25), True, """<c r="A1" t="d" s="1"><v>2011-12-25</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), False, """<c r="A1" t="n" s="1"><v>40902.59994212963</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), True, """<c r="A1" t="d" s="1"><v>2011-12-25T14:23:55</v></c>"""),
                             (datetime.time(14, 15, 25), False, """<c r="A1" t="n" s="1"><v>0.5940393518518519</v></c>"""),
                             (datetime.time(14, 15, 25), True, """<c r="A1" t="d" s="1"><v>14:15:25</v></c>"""),
                             (datetime.timedelta(1, 3, 15), False, """<c r="A1" t="n" s="1"><v>1.000034722395833</v></c>"""),
                             (datetime.timedelta(1, 3, 15), True, """<c r="A1" t="n" s="1"><v>1.000034722395833</v></c>"""),
                         ]
                         )
def test_write_date(worksheet, write_cell_implementation, value, expected, iso_dates):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value
    cell.parent.parent.iso_dates = iso_dates

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, iso_dates",
                         [
                             (datetime.datetime(2021, 3, 19, 23, tzinfo=datetime.timezone.utc), True),
                             (datetime.datetime(2021, 3, 19, 23, tzinfo=datetime.timezone.utc), False),
                             (datetime.time(23, 58, tzinfo=datetime.timezone.utc), True),
                             (datetime.time(23, 58, tzinfo=datetime.timezone.utc), False),
                         ]
                         )
def test_write_invalid_date(worksheet, write_cell_implementation, value, iso_dates):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value
    cell.parent.parent.iso_dates = iso_dates

    out = BytesIO()
    with pytest.raises(TypeError):
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)


@pytest.mark.parametrize("value, expected, epoch",
                         [
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>""",
                              CALENDAR_WINDOWS_1900),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>39440</v></c>""",
                              CALENDAR_MAC_1904),
                         ]
                         )
def test_write_epoch(worksheet, write_cell_implementation, value, expected, epoch):
    write_cell = write_cell_implementation

    ws = worksheet
    ws.parent.epoch = epoch
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_hyperlink(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = "test"
    cell.hyperlink = "http://www.test.com"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    assert len(worksheet._hyperlinks) == 1


@pytest.mark.parametrize("value, result, attrs",
                         [
                             ("test", "test", {'r': 'A1', 't': 'inlineStr'}),
                             ("=SUM(A1:A2)", "=SUM(A1:A2)", {'r': 'A1'}),
                             (datetime.date(2018, 8, 25), 43337, {'r':'A1', 't':'n'}),
                         ]
                         )
def test_attributes(worksheet, value, result, attrs):
    from .._writer import _set_attributes

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    assert(_set_attributes(cell)) == (result, attrs)


def test_whitespace(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet
    cell = ws['A1']
    cell.value = "  whitespace   "

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c t="inlineStr" r="A1">
      <is>
        <t xml:space="preserve">  whitespace   </t>
      </is>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula

def test_table_formula(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet
    cell = ws["A1"]
    cell.value =  DataTableFormula(ref="A1:B10")
    cell.data_type = "f"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="A1">
      <f t="dataTable" ref="A1:B10" />
      <v/>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_array_formula(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet

    cell = ws["E2"]
    cell.value = ArrayFormula(ref="E2:E11", text="=C2:C11*D2:D11")

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="E2">
      <f t="array" ref="E2:E11">C2:C11*D2:D11</f>
      <v/>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_rich_text(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet

    from ..rich_text import CellRichText, TextBlock, InlineFont

    red = InlineFont(color='FF0000')
    rich_string = CellRichText(
        [TextBlock(red, 'red'),
         ' is used, you can expect ',
         TextBlock(red, 'danger')]
    )
    cell = ws["A2"]
    cell.value = rich_string


    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="A2" t="inlineStr">
      <is>
        <r>
        <rPr>
          <color rgb="00FF0000" />
        </rPr>
        <t>red</t>
        </r>
        <r>
          <t xml:space="preserve"> is used, you can expect </t>
        </r>
        <r>
          <rPr>
            <color rgb="00FF0000" />
          </rPr>
          <t>danger</t>
        </r>
      </is>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff
