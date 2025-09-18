package LoadingSavingAndConverting

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ReadAndWriteXLSBFileFormat reads and writes XLSB File Format
func ReadAndWriteXLSBFileFormat() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	srcReadWriteXLSB := dirPath + "srcReadWriteXLSB.xlsb"

	// Path of output excel file
	outReadWriteXLSB := outPath + "outReadWriteXLSB.xlsb"

	// Read source xlsb file
	wb, _ := NewWorkbook_String(srcReadWriteXLSB)

	// Access first worksheet
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Access cell A1
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("A1")

	// Get the string value of cell A1
	strVal, _ := cell.GetStringValue()

	// Print the string value of cell A1
	cellValue := "Cell Value: "
	fmt.Println(cellValue + strVal)

	// Access cell C4
	cell, _ = cells.Get_String("C4")

	// Put the string value of cell A1 into C4
	strValPtr := strVal
	cell.PutValue_String(strValPtr)

	// Save the workbook in XLSB format
	wb.Save_String_SaveFormat(outReadWriteXLSB, SaveFormat_Xlsb)

	// Show successful execution message on console
	ShowMessageOnConsole("ReadAndWriteXLSBFileFormat executed successfully.\r\n\r\n")
}
