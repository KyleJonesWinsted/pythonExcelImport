import xlrd
import json

def table(filename, output=False, sheet_index=0):
  
  book = xlrd.open_workbook(filename)
  sheet = book.sheet_by_index(sheet_index)

  headers = []
  results = []

  for row in range(sheet.nrows):
    result_row = {}
    for col in range(sheet.ncols):
      if row == 0:
        headers.append(sheet.cell_value(row, col))
      else:
        result_row[headers[col]] = sheet.cell_value(row, col)
    if row != 0:
      results.append(result_row)

  results_json = json.dumps(results)

  if output:
    print(results_json)

  return results

if __name__ == "__main__":
  import sys
  if len(sys.argv) == 2:
    table(
      filename=sys.argv[1],
      output=True,
    )
  elif len(sys.argv) == 3:
    table(
      filename=sys.argv[1],
      output=True,
      sheet_index=int(sys.argv[2]),
    )
  else:
    filename = input("Enter filename: ")
    table(filename=filename, output=True)
