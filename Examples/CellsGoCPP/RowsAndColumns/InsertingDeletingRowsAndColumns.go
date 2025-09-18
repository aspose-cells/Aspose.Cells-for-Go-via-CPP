package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Insert a Row
func InsertRow() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleInsertingDeletingRowsAndColumns := dirPath + "sampleInsertingDeletingRowsAndColumns.xlsx"

	// Path of output excel file
	outputInsertingDeletingRowsAndColumns := outPath + "outputInsertingDeletingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleInsertingDeletingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Inserting a row into the worksheet at 3rd position
	cells, _ := ws.GetCells()
	cells.InsertRow(2)

	// Save the Excel file.
	workbook.Save_String(outputInsertingDeletingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("InsertRow executed successfully.\r\n\r\n")
}

// Inserting Multiple Rows
func InsertingMultipleRows() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleInsertingDeletingRowsAndColumns := dirPath + "sampleInsertingDeletingRowsAndColumns.xlsx"

	// Path of output excel file
	outputInsertingDeletingRowsAndColumns := outPath + "outputInsertingDeletingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleInsertingDeletingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Inserting 10 rows into the worksheet starting from 3rd row
	cells, _ := ws.GetCells()
	cells.InsertRows_Int_Int(2, 10)

	// Save the Excel file.
	workbook.Save_String(outputInsertingDeletingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("InsertingMultipleRows executed successfully.\r\n\r\n")
}

// Deleting Multiple Rows
func DeletingMultipleRows() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleInsertingDeletingRowsAndColumns := dirPath + "sampleInsertingDeletingRowsAndColumns.xlsx"

	// Path of output excel file
	outputInsertingDeletingRowsAndColumns := outPath + "outputInsertingDeletingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleInsertingDeletingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Deleting 10 rows from the worksheet starting from 3rd row
	cells, _ := ws.GetCells()
	cells.DeleteRows_Int_Int(2, 10)

	// Save the Excel file.
	workbook.Save_String(outputInsertingDeletingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("DeletingMultipleRows executed successfully.\r\n\r\n")
}

// Insert a Column
func InsertColumn() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleInsertingDeletingRowsAndColumns := dirPath + "sampleInsertingDeletingRowsAndColumns.xlsx"

	// Path of output excel file
	outputInsertingDeletingRowsAndColumns := outPath + "outputInsertingDeletingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleInsertingDeletingRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Inserting a column into the worksheet at 2nd position
	cells, _ := ws.GetCells()
	cells.InsertColumn_Int(1)

	// Save the Excel file.
	workbook.Save_String(outputInsertingDeletingRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("InsertColumn executed successfully.\r\n\r\n")
}

// Delete a Column
func DeleteColumn_() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleDeleteColumn := dirPath + "sampleInsertingDeletingRowsAndColumns.xlsx"

	// Path of output excel file
	outputDeleteColumn := outPath + "outputInsertingDeletingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleDeleteColumn)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Deleting a column from the worksheet at 2nd position
	cells, _ := ws.GetCells()
	cells.DeleteColumn_Int(4)

	// Save the Excel file.
	workbook.Save_String(outputDeleteColumn)

	// Show successful execution message on console
	ShowMessageOnConsole("DeleteColumn executed successfully.\r\n\r\n")
}
