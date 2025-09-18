package Data

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Accessing Cells Using Cell Name
func AccessingCellsUsingCellName() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	sampleData := dirPath + "sampleData.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleData)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Get cells from sheet
	cells, _ := worksheet.GetCells()

	// Accessing a cell using its name
	cell, _ := cells.Get_String("B3")

	// Write string value of the cell on console
	stringValue, _ := cell.GetStringValue()
	fmt.Println("Value of cell B3:", stringValue)

	// Show successful execution message on console
	ShowMessageOnConsole("AccessingCellsUsingCellName executed successfully.\r\n\r\n")
}

// Accessing Cells Using Row & Column Index of the Cell
func AccessingCellsUsingRowAndColumnIndexOfTheCell() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	sampleData := dirPath + "sampleData.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleData)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Get cells from sheet
	cells, _ := worksheet.GetCells()

	// Accessing a cell using its row and column index
	cell, _ := cells.Get_Int_Int(2, 1)

	// Write string value of the cell on console
	stringValue, _ := cell.GetStringValue()
	fmt.Println("Value of cell B3:", stringValue)

	// Show successful execution message on console
	ShowMessageOnConsole("AccessingCellsUsingRowAndColumnIndexOfTheCell executed successfully.\r\n\r\n")
}

// Accessing Maximum Display Range of Worksheet
func AccessingMaximumDisplayRangeOfWorksheet() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	sampleData := dirPath + "sampleData.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleData)

	// Accessing the first worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Get cells from sheet
	cells, _ := worksheet.GetCells()

	// Access the Maximum Display Range
	_range, _ := cells.GetMaxDisplayRange()

	// Print string value of the cell on console
	refersTo, _ := _range.GetRefersTo()
	fmt.Println("Maximum Display Range of Worksheet:", refersTo)

	// Show successful execution message on console
	ShowMessageOnConsole("AccessingMaximumDisplayRangeOfWorksheet executed successfully.\r\n\r\n")
}
