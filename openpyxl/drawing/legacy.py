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
    children = {} # can have emf as children

    def __init__(self, vml):
        self.vml = vml


    @property
    def path(self):
        return self._path.format(self._counter)


    def _write(self, archive, manifest):
        if self.children:
            self._write_rels(archive, manifest=None)
        archive.writestr(self.path[1:], self.vml)


    def _write_rels(self, archive, manifest=None):
        rels = RelationshipList()
        path = get_rels_path(self.path.format(self._counter))
        for k, image in self.children:
            rel = Relationship(Type=IMAGE_NS, Target=image.path)
            rels.append(rel)
        tree = rels.to_tree()
        archive.writestr(path[1:], tostring(tree))

