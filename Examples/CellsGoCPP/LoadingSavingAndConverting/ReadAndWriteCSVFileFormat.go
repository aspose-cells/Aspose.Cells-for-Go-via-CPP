package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ReadAndWriteCSVFileFormat reads and writes CSV file format
func ReadAndWriteCSVFileFormat() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input csv file
	srcReadWriteCSV := dirPath + "srcReadWriteCSV.csv"

	// Path of output csv file
	outReadWriteCSV := outPath + "outReadWriteCSV.csv"

	// Read source csv file
	wb, _ := NewWorkbook_String(srcReadWriteCSV)

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
	cell.PutValue_String(strVal)

	// Save the workbook in csv format
	wb.Save_String_SaveFormat(outReadWriteCSV, SaveFormat_Csv)

	// Show successful execution message on console
	fmt.Println("ReadAndWriteCSVFileFormat executed successfully.\r\n\r\n")
}
