package Data

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Tracing Precedents
func TracingPrecedents() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	samplePrecedentsAndDependents := dirPath + "samplePrecedentsAndDependents.xlsx"

	// Load source Excel file
	workbook, _ := NewWorkbook_String(samplePrecedentsAndDependents)

	// Calculate workbook formula
	workbook.CalculateFormula()

	// Access first worksheet
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Access cell F6
	cells, _ := worksheet.GetCells()
	cell, _ := cells.Get_String("F6")

	// Get precedents of the cells and print them on console
	fmt.Println("Printing Precedents of Cell: ")
	fmt.Println(cell.GetName())
	fmt.Println("")
	fmt.Println("-------------------------------")

	refac, _ := cell.GetPrecedents()
	count, _ := refac.GetCount()
	var i int32
	for i = 0; i < count; i++ {
		refa, _ := refac.Get(i)

		row, _ := refa.GetStartRow()
		col, _ := refa.GetStartColumn()

		cell, _ = cells.Get_Int_Int(row, col)
		fmt.Println(cell.GetName())
	}

	fmt.Println("")

	// Show successful execution message on console
	ShowMessageOnConsole("TracingPrecedents executed successfully.\r\n\r\n")
}

// Tracing Dependents
func TracingDependents() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	samplePrecedentsAndDependents := dirPath + "samplePrecedentsAndDependents.xlsx"

	// Load source Excel file
	workbook, _ := NewWorkbook_String(samplePrecedentsAndDependents)

	// Calculate workbook formula
	workbook.CalculateFormula()

	// Access first worksheet
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Access cell F6
	cells, _ := worksheet.GetCells()
	cell, _ := cells.Get_String("F6")

	// Get dependents of the cells and print them on console
	fmt.Println("Printing Dependents of Cell: ")
	fmt.Println(cell.GetName())
	fmt.Println("")
	fmt.Println("-------------------------------")

	// Parameter false means we do not want to search other sheets
	depCells, _ := cell.GetDependents(false)

	// Get the length of the array
	len := len(depCells)

	// Print the names of all the cells inside the array
	for i := 0; i < len; i++ {
		dCell := depCells[i]
		fmt.Println(dCell.GetName())
	}

	fmt.Println("")

	// Show successful execution message on console
	ShowMessageOnConsole("TracingDependents executed successfully.\r\n\r\n")
}
