package Data

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ManipulateNamedRangeInWorkbook manipulates named ranges in a workbook.
func ManipulateNamedRangeInWorkbook() {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleManipulateNamedRangeInWorkbook := dirPath + "sampleManipulateNamedRangeInWorkbook.xlsx"

	// Path of output excel file
	outputManipulateNamedRangeInWorkbook := outPath + "outputManipulateNamedRangeInWorkbook.xlsx"

	// Create a workbook
	wb, _ := NewWorkbook_String(sampleManipulateNamedRangeInWorkbook)

	// Read the named range created above from names collection
	worksheets, _ := wb.GetWorksheets()
	names, _ := worksheets.GetNames()
	nm, _ := names.Get_Int(0)

	// Print its FullText and RefersTo members
	fullText := "Full Text : "
	text, _ := nm.GetFullText()
	fmt.Println(fullText + text)

	referTo := "Refers To: "
	text, _ = nm.GetRefersTo()
	fmt.Println(referTo + text)

	// Manipulate the RefersTo property of NamedRange
	nm.SetRefersTo_String("=Sheet1!$D$5:$J$10")

	// Save the workbook in xlsx format
	wb.Save_String(outputManipulateNamedRangeInWorkbook)

	// Show successful execution message on console
	ShowMessageOnConsole("ManipulateNamedRangeInWorkbook executed successfully.\r\n\r\n")
}
