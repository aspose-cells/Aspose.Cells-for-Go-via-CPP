package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Setting the Height of a Row
func SettingHeightOfRow() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleRowsAndColumns := dirPath + "sampleRowsAndColumns.xlsx"

	// Path of output excel file
	outputRowsAndColumns := outPath + "outputRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Setting the height of the second row to 35
	cells, _ := worksheet.GetCells()
	cells.SetRowHeight(1, 35)

	// Save the Excel file.
	workbook.Save_String(outputRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("SettingHeightOfRow executed successfully.\r\n\r\n")
}

// Setting the Height of All Rows in a Worksheet
func SettingHeightOfAllRowsInWorksheet() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleRowsAndColumns := dirPath + "sampleRowsAndColumns.xlsx"

	// Path of output excel file
	outputRowsAndColumns := outPath + "outputRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Setting the height of all rows in the worksheet to 25
	cells, _ := worksheet.GetCells()
	cells.SetStandardHeight(25)

	// Save the Excel file.
	workbook.Save_String(outputRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("SettingHeightOfAllRowsInWorksheet executed successfully.\r\n\r\n")
}

// Setting the Width of a Column
func SettingWidthOfColumn() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleRowsAndColumns := dirPath + "sampleRowsAndColumns.xlsx"

	// Path of output excel file
	outputRowsAndColumns := outPath + "outputRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Setting the width of the second column to 56.5
	cells, _ := worksheet.GetCells()
	cells.SetColumnWidth(1, 56.5)

	// Save the Excel file.
	workbook.Save_String(outputRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("SettingWidthOfColumn executed successfully.\r\n\r\n")
}

// Setting the Width of All Columns in a Worksheet
func SettingWidthOfAllColumnsInWorksheet() {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleRowsAndColumns := dirPath + "sampleRowsAndColumns.xlsx"

	// Path of output excel file
	outputRowsAndColumns := outPath + "outputRowsAndColumns.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleRowsAndColumns)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Setting the width of all columns in the worksheet to 20.5
	cells, _ := worksheet.GetCells()
	cells.SetStandardWidth(20.5)

	// Save the Excel file.
	workbook.Save_String(outputRowsAndColumns)

	// Show successful execution message on console
	ShowMessageOnConsole("SettingWidthOfAllColumnsInWorksheet executed successfully.\r\n\r\n")
}
