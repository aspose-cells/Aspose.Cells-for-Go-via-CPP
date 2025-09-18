package LoadingSavingAndConverting

import (
	"fmt"
	"time"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ConvertExcelWorkbookToPDF_SetPDFCreationTime converts Excel Workbook to PDF and sets the PDF creation time.
func ConvertExcelWorkbookToPDF_SetPDFCreationTime() error {
	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output Pdf file.
	outputConvertExcelWorkbookToPDF := outDir + "outputConvertExcelWorkbookToPDF_PDFCreationTime.pdf"

	// Create an empty workbook.
	workbook, err := NewWorkbook()
	if err != nil {
		return err
	}

	// Access first worksheet.
	wss, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Access cell A1.
	cells, err := ws.GetCells()
	if err != nil {
		return err
	}
	cell, err := cells.Get_String("A1")
	if err != nil {
		return err
	}

	// Add some text in cell.
	cell.PutValue_String("PDF Creation Time is 25-May-2017.")

	// Create pdf save options object.
	pdfSaveOptions, err := NewPdfSaveOptions()
	if err != nil {
		return err
	}

	// Set the created time for the PDF i.e. 25-May-2017
	t1 := time.Now()
	// date := Time{Year: 2017, Month: 5, Day: 25}
	pdfSaveOptions.SetCreatedTime(t1)

	// Save the Excel Document in PDF format
	err = workbook.Save_String_SaveOptions(outputConvertExcelWorkbookToPDF, pdfSaveOptions.ToSaveOptions())
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("ConvertExcelWorkbookToPDF_SetPDFCreationTime executed successfully.")
	return nil
}
