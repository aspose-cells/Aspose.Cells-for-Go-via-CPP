package Formulas

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// CalculateWorkbookFormulas calculates workbook formulas
func CalculateWorkbookFormulas() {
	// Create a new workbook
	wb, _ := NewWorkbook()

	// Get first worksheet which is created by default
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Adding a value to "A1" cell
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("A1")
	cell.PutValue_Int(5)

	// Adding a value to "A2" cell
	cell, _ = cells.Get_String("A2")
	cell.PutValue_Int(15)

	// Adding a value to "A3" cell
	cell, _ = cells.Get_String("A3")
	cell.PutValue_Int(25)

	// Adding SUM formula to "A4" cell
	cell, _ = cells.Get_String("A4")
	cell.SetFormula_String("=SUM(A1:A3)")

	// Calculating the results of formulas
	wb.CalculateFormula()

	// Get the calculated value of the cell "A4" and print it on console
	cell, _ = cells.Get_String("A4")
	sCalcuInfo := "Calculated Value of Cell A4: "
	calculatedValue, _ := cell.GetStringValue()
	fmt.Println(sCalcuInfo + calculatedValue)

	// Show successful execution message on console
	ShowMessageOnConsole("CalculateWorkbookFormulas executed successfully.\r\n\r\n")
}
