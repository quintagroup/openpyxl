# Copyright (c) 2010-2021 openpyxl

from openpyxl.packaging.relationship import (
    get_rels_path,
    RelationshipList,
    Relationship,
)
from openpyxl.xml.constants import (
    VML_NS,
    IMAGE_NS,
)
from openpyxl.xml.functions import tostring


class LegacyDrawing:

    mime_type = "application/vnd.openxmlformats-officedocument.vmlDrawing"
    rel_type = VML_NS
    _counter = 0
    _rel_id = None
    _path = "/xl/vmlDrawing{0}.xml"
    vml = None
    children = [] # rels from the worksheet

    def __init__(self, vml):
        self.vml = vml


    @property
    def path(self):
        return self._path.format(self._counter)


    def _write(self, archive):
        archive.writestr(self.path[1:], self.vml)

