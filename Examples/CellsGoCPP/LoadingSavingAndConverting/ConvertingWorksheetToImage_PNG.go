package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ConvertingWorksheetToImage_PNG converts a worksheet to PNG images
func ConvertingWorksheetToImage_PNG() error {
	// For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-C

	// Source directory path.
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input Excel file.
	sampleConvertingWorksheetToDifferentImageFormats := srcDir + "sampleConvertingWorksheetToDifferentImageFormats.xlsx"

	// Create an empty workbook.
	workbook, err := NewWorkbook_String(sampleConvertingWorksheetToDifferentImageFormats)
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

	// Create image or print options object.
	imgOptions, _ := NewImageOrPrintOptions()
	// Specify the image format.
	imgOptions.SetImageType(ImageType_Png)
	// Specify horizontal and vertical resolution
	imgOptions.SetHorizontalResolution(200)
	imgOptions.SetVerticalResolution(200)
	// Render the sheet with respect to specified image or print options.
	sr, _ := NewSheetRender(worksheet, imgOptions)
	// Get page count.
	pageCount, err := sr.GetPageCount()
	if err != nil {
		return err
	}

	// Render each page to png image one by one.
	var i int32
	for i = 0; i < pageCount; i++ {
		// Clear string builder and create output image path with string concatenations.
		outputPNG := outDir + "outputConvertingWorksheetToImagePNG_" + fmt.Sprintf("%d", i) + ".png"

		// Convert worksheet to png image.
		err := sr.ToImage_Int_String(i, outputPNG)
		if err != nil {
			return err
		}
	}

	// Show successful execution message on console
	fmt.Println("ConvertingWorksheetToImage_PNG executed successfully.")

	return nil
}
