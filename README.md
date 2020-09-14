# Python Excel Import

## Converts excel sheet to JSON or Python dictionary

### Command Line

    python3 xlimport.py <filename> <sheetindex>
  
Parses the first sheet or the sheet number given in the supplied file and outputs a JSON string to the standard output.

### Python script

    import xlimport
  
    dictionary = xlimport.table(filename, sheet_index)

Parses the first sheet or the sheet number given in the supplied filename and returns a Python dictionary.
