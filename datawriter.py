import xlsxwriter

def exportTestData(filename, vect):
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some simple text.
    # worksheet.write('A1', 'Hello')

    # Text with formatting.
    # worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    # worksheet.write(2, 0, 123)
    # worksheet.write(3, 0, 123.456)

    # testdata = [[1,2,3], [4,5,6], [7,8,9]]

    for i in range(0, len(vect)):
        for j in range(0,len(vect[0])):
            worksheet.write(i, j, vect[i][j])

    workbook.close()

testdata = [[1,2,3], [4,5,6], [7,8,9]]

exportTestData('demo7.xls', testdata)