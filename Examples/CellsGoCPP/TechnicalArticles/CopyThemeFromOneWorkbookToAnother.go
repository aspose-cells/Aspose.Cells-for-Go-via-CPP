package TechnicalArticles

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// CopyThemeFromOneWorkbookToAnother copies theme from one workbook to another
func CopyThemeFromOneWorkbookToAnother() {
	// Source directory path
	dirPath := "..\\Data\\TechnicalArticles\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Paths of source and output excel files
	damaskPath := dirPath + "DamaskTheme.xlsx"
	sampleCopyThemeFromOneWorkbookToAnother := dirPath + "sampleCopyThemeFromOneWorkbookToAnother.xlsx"
	outputCopyThemeFromOneWorkbookToAnother := outPath + "outputCopyThemeFromOneWorkbookToAnother.xlsx"

	// Read excel file that has Damask theme applied on it
	damask, _ := NewWorkbook_String(damaskPath)

	// Read your sample excel file
	wb, _ := NewWorkbook_String(sampleCopyThemeFromOneWorkbookToAnother)

	// Copy theme from source file
	wb.CopyTheme(damask)

	// Save the workbook in xlsx format
	wb.Save_String_SaveFormat(outputCopyThemeFromOneWorkbookToAnother, SaveFormat_Xlsx)

	// Show successful execution message on console
	fmt.Println("CopyThemeFromOneWorkbookToAnother executed successfully.\r\n\r\n")
}
