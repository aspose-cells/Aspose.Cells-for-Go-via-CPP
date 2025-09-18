package Data

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// AddingDataToCells adds data to cells in an Excel worksheet
func AddingDataToCells() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleData := dirPath + "sampleData.xlsx"

	// Path of output excel file
	outputData := outPath + "outputData.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleData)

	// Accessing the second worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(1)

	// Adding a string value to the cell
	cells, _ := worksheet.GetCells()
	cell, _ := cells.Get_String("A1")
	cell.PutValue_String("Hello World")

	// Adding a double value to the cell
	cell, _ = cells.Get_String("A2")
	cell.PutValue_Double(20.5)

	// Adding an integer value to the cell
	cell, _ = cells.Get_String("A3")
	cell.PutValue_Int(15)

	// Adding a boolean value to the cell
	cell, _ = cells.Get_String("A4")
	cell.PutValue_Bool(true)

	// Setting the display format of the date
	cell, _ = cells.Get_String("A5")
	style, _ := cell.GetStyle()
	style.SetNumber(15)
	cell.SetStyle_Style(style)

	// Save the workbook
	workbook.Save_String(outputData)

	// Show successful execution message on console
	fmt.Println("AddingDataToCells executed successfully.\r\n\r\n")
}

// RetrievingDataFromCells retrieves data from cells in an Excel worksheet
func RetrievingDataFromCells() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	sampleData := dirPath + "sampleData.xlsx"

	// Read input excel file
	workbook, _ := NewWorkbook_String(sampleData)

	// Accessing the third worksheet in the Excel file
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(2)

	// Get cells from sheet
	cells, _ := worksheet.GetCells()

	enCell, _ := cells.GetEnumerator()

	for {
		loop, _ := enCell.MoveNext()
		if !loop {
			break
		}
		cell, _ := enCell.GetCurrent()
		cellType, _ := cell.GetType()
		switch cellType {
		// Evaluating the data type of the cell data for string value
		case CellValueType_IsString:
			fmt.Println("Cell Value Type Is String.")
			_, _ = cell.GetStringValue()
			break
		// Evaluating the data type of the cell data for double value
		case CellValueType_IsNumeric:
			fmt.Println("Cell Value Type Is Numeric.")
			_, _ = cell.GetDoubleValue()
			break
		// Evaluating the data type of the cell data for boolean value
		case CellValueType_IsBool:
			fmt.Println("Cell Value Type Is Bool.")
			_, _ = cell.GetBoolValue()
			break
		// Evaluating the data type of the cell data for date/time value
		case CellValueType_IsDateTime:
			fmt.Println("Cell Value Type Is DateTime.")
			_, _ = cell.GetDateTimeValue()
			break
		// Evaluating the unknown data type of the cell data
		case CellValueType_IsUnknown:
			_, _ = cell.GetStringValue()
			break
		default:
			break
		}
	}

	// Show successful execution message on console
	fmt.Println("RetrievingDataFromCells executed successfully.\r\n\r\n")
}
