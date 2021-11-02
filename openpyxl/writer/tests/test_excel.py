from __future__ import absolute_import
# Copyright (c) 2010-2021 openpyxl

from io import BytesIO
import os
from string import ascii_letters
from zipfile import ZipFile

import pytest

from openpyxl import load_workbook
from openpyxl.chart import BarChart
from openpyxl.comments import Comment
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils.exceptions import InvalidFileException


@pytest.fixture
def ExcelWriter():
    from ..excel import ExcelWriter
    return ExcelWriter


@pytest.fixture
def archive():
    out = BytesIO()
    return ZipFile(out, "w")


class TestExcelWriter:


    def test_worksheet(self, ExcelWriter, archive):
        wb = Workbook()
        ws = wb.active
        writer = ExcelWriter(wb, archive)
        writer._write_worksheets()

        assert ws.path[1:] in archive.namelist()
        assert ws.path in writer.manifest.filenames


    def test_tables(self, ExcelWriter, archive):
        wb = Workbook()
        ws = wb.active
        ws.append(list(ascii_letters))
        ws._rels = []
        t = Table(displayName="Table1", ref="A1:D10")
        ws.add_table(t)

        writer = ExcelWriter(wb, archive)
        writer._write_worksheets()

        assert t.path[1:] in archive.namelist()
        assert t.path in writer.manifest.filenames


    def test_drawing(self, ExcelWriter, archive):
        wb = Workbook()
        ws = wb.active

        drawing = SpreadsheetDrawing()

        writer = ExcelWriter(wb, archive)
        writer._write_drawing(drawing)
        assert drawing.path == '/xl/drawings/drawing1.xml'
        assert drawing.path[1:] in archive.namelist()
        assert drawing.path in writer.manifest.filenames


    def test_write_chart(self, ExcelWriter, archive):
        wb = Workbook()
        ws = wb.active

        chart = BarChart()
        ws.add_chart(chart)

        writer = ExcelWriter(wb, archive)
        writer._write_worksheets()
        assert 'xl/worksheets/sheet1.xml' in archive.namelist()
        assert ws.path in writer.manifest.filenames

        rel = ws._rels["rId1"]
        assert dict(rel) == {'Id': 'rId1', 'Target': '/xl/drawings/drawing1.xml',
                             'Type':
                             'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}


    @pytest.mark.pil_required
    def test_write_images(self, datadir, ExcelWriter, archive):
        from openpyxl.drawing.image import Image
        datadir.chdir()

        writer = ExcelWriter(None, archive)

        img = Image("plain.png")
        writer._images.append(img)

        writer._write_images()
        archive.close()

        #zipinfo = archive.infolist()
        assert 'xl/media/image1.png' in archive.namelist()


    def test_chartsheet(self, ExcelWriter, archive):
        wb = Workbook()
        cs = wb.create_chartsheet()

        writer = ExcelWriter(wb, archive)
        writer._write_chartsheets()

        assert cs.path in writer.manifest.filenames
        assert cs.path[1:] in writer._archive.namelist()


    def test_comment(self, ExcelWriter, archive):

        wb = Workbook()
        ws = wb.active
        ws['B5'].comment = Comment("A comment", "The Author")

        writer = ExcelWriter(None, archive)
        writer._write_comment(ws)

        assert archive.namelist() == ['xl/comments/comment1.xml', 'xl/drawings/commentsDrawing1.vml']
        assert '/xl/comments/comment1.xml' in writer.manifest.filenames
        assert ws.legacy_drawing == 'xl/drawings/commentsDrawing1.vml'


    def test_duplicate_comment(self, ExcelWriter, archive):

        wb = Workbook()
        ws = wb.active
        ws['B5'].comment = Comment("A comment", "The Author")

        writer = ExcelWriter(wb, archive)
        writer.write_worksheet(ws)
        writer.write_worksheet(ws)
        assert len(ws._comments) == 1


    @pytest.mark.xfail
    def test_merge_vba(self, ExcelWriter, archive, datadir):
        from openpyxl import load_workbook
        datadir.chdir()
        wb = load_workbook("vba+comments.xlsm", keep_vba=True)

        writer = ExcelWriter(wb, archive)
        writer._merge_vba()

        assert set(archive.namelist()) ==  set([
            'xl/vbaProject.bin',
            'xl/drawings/vmlDrawing1.vml',
            'xl/ctrlProps/ctrlProp3.xml',
            'xl/ctrlProps/ctrlProp1.xml',
            'xl/ctrlProps/ctrlProp10.xml',
            'xl/ctrlProps/ctrlProp9.xml',
            'xl/ctrlProps/ctrlProp4.xml',
            'xl/ctrlProps/ctrlProp5.xml',
            'xl/ctrlProps/ctrlProp6.xml',
            'xl/ctrlProps/ctrlProp7.xml',
            'xl/ctrlProps/ctrlProp8.xml',
            'xl/ctrlProps/ctrlProp2.xml',
        ])


    def test_duplicate_chart(self, ExcelWriter, archive):
        from openpyxl.chart import PieChart
        pc = PieChart()
        wb = Workbook()
        writer = ExcelWriter(wb, archive)

        writer._charts = [pc]*2
        with pytest.raises(InvalidFileException):
            writer._write_charts()


    def test_controls(self, ExcelWriter, archive):
        from openpyxl.worksheet.controls import ControlList, ActiveXControl
        from openpyxl.xml.functions import fromstring
        from openpyxl.packaging.relationship import Relationship

        src = """
        <controls xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice Requires="x14">
            <control shapeId="47129" r:id="rId8" name="MainSVCheckBox">
              <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                <anchor moveWithCells="1">
                  <from>
                    <xdr:col>12</xdr:col>
                    <xdr:colOff>219075</xdr:colOff>
                    <xdr:row>7</xdr:row>
                    <xdr:rowOff>95250</xdr:rowOff>
                  </from>
                  <to>
                    <xdr:col>12</xdr:col>
                    <xdr:colOff>400050</xdr:colOff>
                    <xdr:row>7</xdr:row>
                    <xdr:rowOff>276225</xdr:rowOff>
                  </to>
                </anchor>
              </controlPr>
            </control>
          </mc:Choice>
          <mc:Fallback>
            <control shapeId="47129" r:id="rId8" name="MainSVCheckBox"/>
          </mc:Fallback>
        </mc:AlternateContent>
        </controls>
        """
        tree = fromstring(src)
        controls = ControlList.from_tree(tree)
        ctrl = controls.control[0]
        ctrl.shape = ActiveXControl()
        prop = ctrl.controlPr
        prop.image = Relationship(type="image", Target="")
        prop.image.blob = b"\001"

        wb = Workbook()
        ws = wb.active
        ws.controls = controls
        writer = ExcelWriter(wb, archive)
        writer.write_worksheet(ws)

        assert archive.namelist() == [
            '/xl/activeX/activeX1.xml',
            'xl/worksheets/sheetNone.xml',
        ]


def test_write_empty_workbook(tmpdir):
    tmpdir.chdir()
    wb = Workbook()
    from ..excel import save_workbook
    dest_filename = 'empty_book.xlsx'
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_write_virtual_workbook():
    old_wb = Workbook()
    from ..excel import save_virtual_workbook
    saved_wb = save_virtual_workbook(old_wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    assert new_wb
