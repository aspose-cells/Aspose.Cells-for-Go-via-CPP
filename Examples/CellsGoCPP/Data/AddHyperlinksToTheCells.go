package Data

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// AddHyperlinksToTheCells adds Hyperlinks to the Cells
func AddHyperlinksToTheCells() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputAddHyperlinksToTheCells := outPath + "outputAddHyperlinksToTheCells.xlsx"

	// Create a new workbook
	workbook, _ := NewWorkbook()

	// Get the first worksheet
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Add hyperlink in cell C7 and make use of its various methods
	hypLnks, _ := ws.GetHyperlinks()
	idx, _ := hypLnks.Add_String_Int_Int_String("C7", 1, 1, "http://www.aspose.com/")
	lnk, _ := hypLnks.Get(idx)
	lnk.SetTextToDisplay("Aspose")
	lnk.SetScreenTip("Link to Aspose Website")

	// Save the workbook
	workbook.Save_String(outputAddHyperlinksToTheCells)

	// Show successful execution message on console
	fmt.Println("AddHyperlinksToTheCells executed successfully.\r\n\r\n")
}
