package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

func CopyingRows() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleCopyingRowsAndColumns := dirPath + "sampleCopyingRowsAndColumns.xlsx"

	// Path of output excel file
	outputCopyingRowsAndColumns := outPath + "outputCopyingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleCopyingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Copy the second row with data, formattings, images and drawing objects to the 16th row in the worksheet.
	cells, _ := ws.GetCells()
	cells.CopyRow(cells, 1, 15)

	// Save the Excel file.
	workbook.Save_String(outputCopyingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("CopyingRows executed successfully.\r\n\r\n")
}

func CopyingColumns() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleCopyingRowsAndColumns := dirPath + "sampleCopyingRowsAndColumns.xlsx"

	// Path of output excel file
	outputCopyingRowsAndColumns := outPath + "outputCopyingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleCopyingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Copy the third column to eighth column
	cells, _ := ws.GetCells()
	cells.CopyColumn(cells, 2, 7)

	// Save the Excel file.
	workbook.Save_String(outputCopyingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("CopyingColumns executed successfully.\r\n\r\n")
}
