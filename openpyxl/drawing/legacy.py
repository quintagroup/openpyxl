# Copyright (c) 2010-2021 openpyxl

from openpyxl.xml.constants import VML_NS


class LegacyDrawing:

    mime_type = "application/vnd.openxmlformats-officedocument.vmlDrawing"
    rel_type = VML_NS
    _counter = 0
    _rel_id = None
    _path = "/xl/vmlDrawing{0}.xml"
    content = None
    children = [] # can have emf as children


    @property
    def path(self):
        return self._path.format(self._counter)


    def _write(self, archive, manifest):
        self._write_rels(archive, manifest)


    def _write_rels(self, archive, manifest):
        pass
