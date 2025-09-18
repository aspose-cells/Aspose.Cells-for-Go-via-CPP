package LoadingSavingAndConverting

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ReadAndWriteTabDelimitedFileFormat reads and writes a tab delimited file format
func ReadAndWriteTabDelimitedFileFormat() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input tab delimited file
	srcReadWriteTabDelimited := dirPath + "srcReadWriteTabDelimited.txt"

	// Path of output tab delimited file
	outReadWriteTabDelimited := outPath + "outReadWriteTabDelimited.txt"

	// Read source tab delimited file
	wb, _ := NewWorkbook_String(srcReadWriteTabDelimited)

	// Access first worksheet
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Access cell A1
	cells, _ := ws.GetCells()
	cellA1, _ := cells.Get_String("A1")

	// Get the string value of cell A1
	strVal, _ := cellA1.GetStringValue()

	// Print the string value of cell A1
	cellValue := "Cell Value: "
	fmt.Println(cellValue + strVal)

	// Access cell C4
	cellC4, _ := cells.Get_String("C4")

	// Put the string value of cell A1 into C4
	strValPtr := strVal
	cellC4.PutValue_String(strValPtr)

	// Save the workbook in tab delimited format
	wb.Save_String_SaveFormat(outReadWriteTabDelimited, SaveFormat_Tsv)

	// Show successful execution message on console
	ShowMessageOnConsole("ReadAndWriteTabDelimitedFileFormat executed successfully.\r\n\r\n")
}
