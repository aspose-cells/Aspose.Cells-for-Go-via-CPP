package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ReadAndWriteXLSMFileFormat demonstrates reading and writing an XLSM file format
func ReadAndWriteXLSMFileFormat() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	srcReadWriteXLSM := dirPath + "srcReadWriteXLSM.xlsm"

	// Path of output excel file
	outReadWriteXLSM := outPath + "outReadWriteXLSM.xlsm"

	// Read source xlsm file
	wb, _ := NewWorkbook_String(srcReadWriteXLSM)

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
	cellC4.PutValue_String(strVal)

	// Save the workbook in XLSM format
	wb.Save_String_SaveFormat(outReadWriteXLSM, SaveFormat_Xlsm)

	// Show successful execution message on console
	fmt.Println("ReadAndWriteXLSMFileFormat executed successfully.\r\n\r\n")
}
