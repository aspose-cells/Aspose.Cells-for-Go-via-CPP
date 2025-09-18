package DrawingObjects

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

func InsertingOLEObjectsIntoWorksheet() error {

	// Source directory path.
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output Excel file.
	outputInsertingOLEObjectsIntoWorksheet := outDir + "outputInsertingOLEObjectsIntoWorksheet.xlsx"

	// Instantiate a new workbook.
	workbook, err := NewWorkbook()
	if err != nil {
		return err
	}

	// Get the first worksheet.
	wss, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Read Image for Ole Object into array of bytes.
	imagePath := srcDir + "AsposeLogo.png"
	imageData, err := GetDataFromFile(imagePath)
	if err != nil {
		return err
	}

	// Read Ole Object into array of bytes.
	oleObjectPath := srcDir + "inputInsertOleObject.xlsx"
	oleObjectData, err := GetDataFromFile(oleObjectPath)
	if err != nil {
		return err
	}

	// Add an Ole object into the worksheet with the image.
	oleObjects, err := ws.GetOleObjects()
	idx, err := oleObjects.Add_Int_Int_Int_Int_Stream(2, 2, 200, 220, imageData)
	if err != nil {
		return err
	}

	// Set the Ole object data.
	oleObjects, err = ws.GetOleObjects()
	if err != nil {
		return err
	}
	oleObj, err := oleObjects.Get(idx)
	if err != nil {
		return err
	}
	err = oleObj.SetObjectData(oleObjectData)
	if err != nil {
		return err
	}

	// Save the workbook.
	err = workbook.Save_String(outputInsertingOLEObjectsIntoWorksheet)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	ShowMessageOnConsole("InsertingOLEObjectsIntoWorksheet executed successfully.")

	return nil
}
