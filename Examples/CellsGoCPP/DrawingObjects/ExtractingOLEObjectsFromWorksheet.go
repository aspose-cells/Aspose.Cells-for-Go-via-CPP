package DrawingObjects

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Extracting OLE Objects From Worksheet
func ExtractingOLEObjectsFromWorksheet() {
	// Source directory path.
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleExtractingOLEObjectsFromWorksheet := srcDir + "sampleExtractingOLEObjectsFromWorksheet.xlsx"

	// Load sample Excel file containing OLE objects.
	workbook, _ := NewWorkbook_String(sampleExtractingOLEObjectsFromWorksheet)

	// Get the first worksheet.
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Access the count of Ole objects.
	oleObjects, _ := worksheet.GetOleObjects()
	oleCount, _ := oleObjects.GetCount()

	// Iterate all the Ole objects and save to disk with correct file format extension.
	var i int32
	for i = 0; i < oleCount; i++ {
		// Access Ole object.
		oleObj, _ := oleObjects.Get(i)

		// Access the Ole ProgID.
		strProgId, _ := oleObj.GetProgID()

		// Find the correct file extension.
		fileExt := ""
		if strProgId == "Document" {
			fileExt = ".docx"
		} else if strProgId == "Presentation" {
			fileExt = ".pptx"
		} else if strProgId == "Acrobat Document" {
			fileExt = ".pdf"
		}

		// Find the correct file name with file extension.
		fileName := outDir + "outputExtractOleObject" + fileExt

		// Save the object data to file.
		objectData, _ := oleObj.GetObjectData()
		SaveDataToFile(objectData, fileName)
	} // for

	// Show successful execution message on console
	ShowMessageOnConsole("ExtractingOLEObjectsFromWorksheet executed successfully.")
}
