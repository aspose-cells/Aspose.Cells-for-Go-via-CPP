package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ConvertExcelWorkbookToPDF_A_CompliedFiles converts an Excel Workbook to PDF/A complied files.
func ConvertExcelWorkbookToPDF_A_CompliedFiles() error {

	// Output directory path.
	outDir := "../Data/02_OutputDirectory/"

	// Path of output Pdf file.
	outputConvertExcelWorkbookToPDF := outDir + "outputConvertExcelWorkbookToPDF_PdfCompliance_PdfA1b.pdf"

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
	worksheet, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Access cell A1.
	cells, err := worksheet.GetCells()
	if err != nil {
		return err
	}
	cell, err := cells.Get_String("A1")
	if err != nil {
		return err
	}

	// Add some text in cell.
	cell.PutValue_String("Testing PDF/A")

	// Create pdf save options object.
	pdfSaveOptions, err := NewPdfSaveOptions()
	if err != nil {
		return err
	}

	// Set the compliance to PDF/A-1b.
	pdfSaveOptions.SetCompliance(PdfCompliance_PdfA1b)

	// Save the Excel Document in PDF format
	err = workbook.Save_String_SaveOptions(outputConvertExcelWorkbookToPDF, pdfSaveOptions.ToSaveOptions())
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("ConvertExcelWorkbookToPDF_A_CompliedFiles executed successfully.")

	return nil
}
