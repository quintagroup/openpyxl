Connections
===========

Support for the Connections part is provided. The resulting object is accessible
as part of the workbook object. An instance of this part type describes all of the
connections currently established for a workbook. Workbooks typically contain this
part when there's some sort of connection to an external data source. So things like
OLAP, Databases or even real time data sources.

Typical use of the library does not require you to access or modify these elements but should
you require it they can be found as follows:

.. code::

    >>> from openpyxl import load_workbook
    >>>
    >>> wb = load_workbook("sample_with_metadata.xlsx") # This doc uses OLAP data functions
    >>> connections = wb._connections
    >>>
    >>> connections[1].name # indexed by connection id
    'ThisWorkbookDataModel'
    >>> connections[1].description
    'Data Model'
    >>> conn.type_descriptions[conn.type]
    'OLE DB-based source'

.. note::

    Some connections in Excel documents may have a 'type' attribute that is outside the OOXML spec.
    When this happens a warning is given and the connection will be removed::

        UserWarning: Connections type 100 is not supported, references to it will be dropped to keep the Workbook valid
