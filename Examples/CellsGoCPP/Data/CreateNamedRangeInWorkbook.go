package Data

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// CreateNamedRangeInWorkbook creates a named range in a workbook
func CreateNamedRangeInWorkbook() {
	// Source directory path
	dirPath := "../Data/Data/"

	// Path of output excel file
	outputCreateNamedRange := dirPath + "outputCreateNamedRange.xlsx"

	// Create a workbook
	wb, _ := NewWorkbook()

	// Access first worksheet
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Create a range
	cells, _ := ws.GetCells()
	rng, _ := cells.CreateRange_String("A5:C10")

	// Set its name to make it named range
	rng.SetName("MyNamedRange")

	// Read the named range created above from names collection
	names, _ := wss.GetNames()
	nm, _ := names.Get_Int(0)

	// Print its FullText and RefersTo members
	fullText := "Full Text : "
	value, _ := nm.GetFullText()
	fullText += value
	fmt.Println(fullText)

	referTo := "Refers To: "
	value, _ = nm.GetRefersTo()
	referTo += value
	fmt.Println(referTo)

	// Save the workbook in xlsx format
	wb.Save_String(outputCreateNamedRange)

	// Show successful execution message on console
	ShowMessageOnConsole("CreateNamedRangeInWorkbook executed successfully.\r\n\r\n")
}
