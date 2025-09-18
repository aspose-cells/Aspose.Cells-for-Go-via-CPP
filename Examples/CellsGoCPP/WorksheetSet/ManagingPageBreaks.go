package WorksheetSet

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Adding Page Breaks
func AddingPageBreaks() error {
	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output excel file
	outputPageBreaks := outDir + "outputManagingPageBreaks.xlsx"

	// Instantiating a Workbook object
	workbook, _ := NewWorkbook()

	// Add a page break at cell J20
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)
	_ = ws.AddPageBreaks("J20")

	// Save the Excel file.
	err := workbook.Save_String(outputPageBreaks)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("AddingPageBreaks executed successfully.\r\n\r\n")
	return nil
}
