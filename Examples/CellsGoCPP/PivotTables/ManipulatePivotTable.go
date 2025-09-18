package LoadingSavingAndConverting

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ManipulatePivotTable manipulates a pivot table in an Excel file.
func ManipulatePivotTable() {
	// Source directory path
	dirPath := "..\\Data\\PivotTables\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleManipulatePivotTable := dirPath + "sampleManipulatePivotTable.xlsx"

	// Path of output excel file
	outputManipulatePivotTable := outPath + "outputManipulatePivotTable.xlsx"

	// Load the sample excel file
	wb, _ := NewWorkbook_String(sampleManipulatePivotTable)

	// Access first worksheet
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Change value of cell B3 which is inside the source data of pivot table
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("B3")
	cell.PutValue_String("Cup")

	// Get the value of cell H8 before refreshing pivot table
	valCell, _ := cells.Get_String("H8")
	val, _ := valCell.GetStringValue()
	fmt.Println("Before refreshing Pivot Table value of cell H8:", val)

	// Access pivot table, refresh and calculate it
	pivotTables, _ := ws.GetPivotTables()
	pt, _ := pivotTables.Get_Int(0)
	pt.RefreshData()
	pt.CalculateData()

	// Get the value of cell H8 after refreshing pivot table
	valCell, _ = cells.Get_String("H8")
	val, _ = valCell.GetStringValue()
	fmt.Println("After refreshing Pivot Table value of cell H8:", val)

	// Save the output excel file
	wb.Save_String(outputManipulatePivotTable)

	// Show successful execution message on console
	ShowMessageOnConsole("ManipulatePivotTable executed successfully.\r\n\r\n")
}
