data-diff
=========

Compare data sets which should be equivalent, but differ in structure, and create a report of the discrepancies.
The original use case is a study which compares two different instruments for questionnaire-based research.
Interviews were recorded both on paper forms and then double-entered into a computer, and directly entered into
electronic forms on a tablet PC. The optimal way of implementing questionnaires differs between the two instrument.
The createDiff.py script maps each encoding to a master variable, compares data from the two instruments, and creates a
report in the form of an Excel spreadsheet.

DemoDataPaper.csv and DemoDataTablet.csv are to sample data sets. mappingDemo.csv contains the mapping of
instrument-specific encodings to master variables. The mappings can involve renaming of variables an/or arbitrary
transformations of values which are specified as Python snippets.

