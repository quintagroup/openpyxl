Connections
===========

.. testsetup:: connections

   import os
   os.chdir(os.path.join("..", "data"))


Support for the Connections part is provided. The resulting object is accessible
as part of the workbook object. An instance of this part type describes all of the
connections currently established for a workbook. Workbooks typically contain this
part when there's some sort of connection to an external data source. So things like
OLAP, Databases or even real time data sources.

Typical use of the library does not require you to access or modify these elements but should
you require it they can be found as follows:

.. doctest:: connections

    >>> from openpyxl import load_workbook
    >>>
    >>> wb = load_workbook("sample_with_metadata.xlsx") # This doc uses OLAP data functions
    >>> connections = wb._connections
    >>>
    >>> connections.connection[0].name
    'Query - Data'
    >>> connections.connection[0].description
    "Connection to the 'Data' query in the workbook."


.. testcleanup:: connections

   import os
   os.chdir(os.path.join("..", "tmp"))
