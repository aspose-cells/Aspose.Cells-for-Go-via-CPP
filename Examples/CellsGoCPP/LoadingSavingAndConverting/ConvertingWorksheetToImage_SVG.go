package LoadingSavingAndConverting

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ConvertingWorksheetToImage_SVG converts a worksheet to an image in SVG format.
func ConvertingWorksheetToImage_SVG() {
	// Source directory path.
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input Excel file.
	sampleConvertingWorksheetToDifferentImageFormats := srcDir + "sampleConvertingWorksheetToDifferentImageFormats.xlsx"

	// Create an empty workbook.
	workbook, _ := NewWorkbook_String(sampleConvertingWorksheetToDifferentImageFormats)

	// Access first worksheet.
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Create image or print options object.
	imgOptions, _ := NewImageOrPrintOptions()

	// Specify the image format.
	imgOptions.SetImageType(ImageType_Svg)

	// Specify horizontal and vertical resolution
	imgOptions.SetHorizontalResolution(200)
	imgOptions.SetVerticalResolution(200)

	// Render the sheet with respect to specified image or print options.
	sr, _ := NewSheetRender(worksheet, imgOptions)

	// Get page count.
	pageCount, _ := sr.GetPageCount()

	// Render each page to svg image one by one.
	var i int32
	for i = 0; i < pageCount; i++ {
		// Create output image path with string concatenations.
		outputSvg := outDir + "outputConvertingWorksheetToImageSVG_" + fmt.Sprintf("%d", i) + ".svg"

		// Convert worksheet to svg image.
		sr.ToImage_Int_String(i, outputSvg)
	}

	// Show successful execution message on console
	ShowMessageOnConsole("ConvertingWorksheetToImage_SVG executed successfully.")
}
