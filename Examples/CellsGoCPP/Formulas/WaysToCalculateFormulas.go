package Formulas

import (
	"fmt"
	. "main/Common"
	"time"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// AddingFormulasAndCalculatingResults adds formulas and calculates results.
func AddingFormulasAndCalculatingResults() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputAddingFormulasAndCalculatingResults := outPath + "outputAddingFormulasAndCalculatingResults.xlsx"

	// Create workbook
	wb, _ := NewWorkbook()

	// Access first worksheet in the workbook
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Adding integer values to cells A1, A2 and A3
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("A1")
	cell.PutValue_Int(10)
	cell, _ = cells.Get_String("A2")
	cell.PutValue_Int(20)
	cell, _ = cells.Get_String("A3")
	cell.PutValue_Int(70)

	// Adding a SUM formula to "A4" cell
	cell, _ = cells.Get_String("A4")
	cell.SetFormula_String("=SUM(A1:A3)")

	// Calculating the results of formulas
	wb.CalculateFormula()

	// Get the calculated value of the cell
	cell, _ = cells.Get_String("A4")
	strVal, _ := cell.GetStringValue()

	// Print the calculated value on console
	fmt.Printf("Calculated Result: %s\n", strVal)

	// Saving the workbook
	wb.Save_String(outputAddingFormulasAndCalculatingResults)

	// Show successful execution message on console
	ShowMessageOnConsole("AddingFormulasAndCalculatingResults executed successfully.\r\n\r\n")
}

// CalculatingFormulasOnceOnly calculates formulas once only.
func CalculatingFormulasOnceOnly() {
	// Source directory path
	dirPath := "..\\Data\\Formulas\\"

	// Path of input excel file
	sampleCalculatingFormulasOnceOnly := dirPath + "sampleCalculatingFormulasOnceOnly.xlsx"

	// Create workbook
	wb, _ := NewWorkbook_String(sampleCalculatingFormulasOnceOnly)

	// Set the CreateCalcChain as false
	settings, _ := wb.GetSettings()
	formulaSettings, _ := settings.GetFormulaSettings()
	formulaSettings.SetEnableCalculationChain(false)

	// Get the time before formula calculation
	startTime := time.Now()

	// Calculate the workbook formulas
	wb.CalculateFormula()

	// Get the time after formula calculation
	interval := time.Since(startTime)
	fmt.Printf("Workbook Formula Calculation Elapsed Time in Milliseconds: %d\n", interval.Milliseconds())

	// Show successful execution message on console
	ShowMessageOnConsole("CalculatingFormulasOnceOnly executed successfully.\r\n\r\n")
}
