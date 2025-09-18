package WorksheetSet

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Adding Worksheets to a New Excel File
func AddingWorksheetsToNewExcelFile() {
	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output excel file
	outputManageWorksheets := outDir + "outputManageWorksheets.xlsx"

	// Create workbook
	workbook, _ := NewWorkbook()

	// Adding a new worksheet to the Workbook object
	worksheets, _ := workbook.GetWorksheets()
	i, _ := worksheets.Add()

	// Obtaining the reference of the newly added worksheet by passing its sheet index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(i)

	// Setting the name of the newly added worksheet
	worksheet.SetName("My Worksheet")

	// Save the Excel file.
	workbook.Save_String(outputManageWorksheets)

	fmt.Println("New worksheet added successfully with in a workbook!")

	// Show successful execution message on console
	ShowMessageOnConsole("AddingWorksheetsToNewExcelFile executed successfully.\r\n\r\n")
}

// Accessing Worksheets using Sheet Index
func AccessingWorksheetsUsingSheetIndex() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Path of input excel file
	sampleManageWorksheets := srcDir + "sampleManageWorksheets.xlsx"

	// Load the sample Excel file
	workbook, _ := NewWorkbook_String(sampleManageWorksheets)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Access the cell by its name
	cells, _ := worksheet.GetCells()
	cell, _ := cells.Get_String("F7")

	// Print the value of cell F7
	val, _ := cell.GetStringValue()

	// Print the value on console
	fmt.Println("Value of cell F7:", val)

	// Show successful execution message on console
	ShowMessageOnConsole("AccessingWorksheetsUsingSheetIndex executed successfully.\r\n\r\n")
}

// Removing Worksheets using Sheet Index
func RemovingWorksheetsUsingSheetIndex() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleManageWorksheets := srcDir + "sampleManageWorksheets.xlsx"

	// Path of output excel file
	outputManageWorksheets := outDir + "outputManageWorksheets.xlsx"

	// Load the sample Excel file
	workbook, _ := NewWorkbook_String(sampleManageWorksheets)

	// Removing a worksheet using its sheet index
	worksheets, _ := workbook.GetWorksheets()
	worksheets.RemoveAt_Int(0)

	// Save the Excel file
	workbook.Save_String(outputManageWorksheets)

	// Show successful execution message on console
	ShowMessageOnConsole("RemovingWorksheetsUsingSheetIndex executed successfully.\r\n\r\n")
}
