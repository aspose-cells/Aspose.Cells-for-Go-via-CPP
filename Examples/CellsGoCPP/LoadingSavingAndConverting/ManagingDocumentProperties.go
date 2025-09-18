package LoadingSavingAndConverting

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ManagingDocumentProperties manages document properties in an Excel file
func ManagingDocumentProperties() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Paths of source and output excel files
	samplePath := dirPath + "sampleManagingDocumentProperties.xlsx"
	outputPath := outPath + "outputManagingDocumentProperties.xlsx"

	// Load the sample excel file
	wb, _ := NewWorkbook_String(samplePath)

	// Read built-in title and subject properties
	builtInProps, _ := wb.GetBuiltInDocumentProperties()
	strTitle, _ := builtInProps.GetTitle()
	strSubject, _ := builtInProps.GetSubject()

	title := "Title: "
	fmt.Println(title, strTitle)

	subject := "Subject: "
	fmt.Println(subject, strSubject)

	// Modify built-in title and subject properties
	strTitle = "Aspose.Cells New Title"
	strSubject = "Aspose.Cells New Subject"
	builtInProps.SetTitle(strTitle)
	builtInProps.SetSubject(strSubject)

	// Read the custom property
	strCustomPropName := "MyCustom1"
	customProps, _ := wb.GetCustomDocumentProperties()
	strCustomPropValue, _ := customProps.Get_String(strCustomPropName)
	myCustom1 := "\r\nMyCustom1: "
	fmt.Println(myCustom1, strCustomPropValue)

	// Add a new custom property
	strCustomPropName = "MyCustom5"
	strCustomPropValue1 := "This is my custom five."
	customProps.Add_String_String(strCustomPropName, strCustomPropValue1)

	// Save the output excel file
	wb.Save_String(outputPath)

	// Show successful execution message on console
	ShowMessageOnConsole("ManagingDocumentProperties executed successfully.\r\n\r\n")
}
