package Data

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Creating Subtotals
func CreatingSubtotals() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleCreatingSubtotals := dirPath + "sampleCreatingSubtotals.xlsx"

	// Path of output excel file
	outputCreatingSubtotals := outPath + "outputCreatingSubtotals.xlsx"

	// Load sample excel file into a workbook object
	wb, _ := NewWorkbook_String(sampleCreatingSubtotals)

	// Get first worksheet of the workbook
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Get the cells collection of the worksheet
	cells, _ := ws.GetCells()

	// Create cell area covering the cell range B3:C19
	ca, _ := CellArea_CreateCellArea_String_String("B3", "C19")

	// Create integer array of size 1 and set its first value to 1
	totalList := []int32{1}

	// Apply subtotal, the consolidation function is Sum and it will be applied to the second column
	cells.Subtotal_CellArea_Int_ConsolidationFunction_int32Array(ca, 0, ConsolidationFunction_Sum, totalList)

	// Save the workbook in xlsx format
	wb.Save_String(outputCreatingSubtotals)

	// Show successful execution message on console
	fmt.Println("CreatingSubtotals executed successfully.\r\n\r\n")
}
