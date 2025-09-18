package main

import (
	"fmt"
	. "main/Charts"
	. "main/Common"
	. "main/Data"
	. "main/DrawingObjects"
	. "main/Formulas"
	. "main/LoadingSavingAndConverting"
	. "main/PivotTables"
	. "main/RowsAndColumns"
	. "main/TechnicalArticles"
	. "main/WorksheetSet"
)

func main() {
	//Show Aspose.Cells for Go Version
	ShowCellsVersion()

	// Uncomment the one you want to try out

	// Charts
	print("Charts\n")
	CreatingAndCustomizingCharts_BubbleChart()
	CreatingAndCustomizingCharts_CustomChart()
	CreatingAndCustomizingCharts_LineChart()
	CreatingAndCustomizingCharts_PyramidChart()
	ReadAndManipulateExcel2016Charts()

	// Data
	print("Data\n")
	AccessingCellsUsingCellName()
	AccessingCellsUsingRowAndColumnIndexOfTheCell()
	AccessingMaximumDisplayRangeOfWorksheet()
	AddingDataToCells()
	RetrievingDataFromCells()
	AddHyperlinksToTheCells()
	ApplyConditionalFormattingInWorksheet()
	FormatCellOrRangeOfCells()
	CreateNamedRangeInWorkbook()
	CreatingSubtotals()
	FindOrSearchData()
	ManipulateNamedRangeInWorkbook()
	TracingPrecedents()
	TracingDependents()

	// DrawingObjects
	ExtractingOLEObjectsFromWorksheet()
	InsertingOLEObjectsIntoWorksheet()

	// Formulas
	CalculateWorkbookFormulas()
	AddingFormulasAndCalculatingResults()
	CalculatingFormulasOnceOnly()

	// LoadingSavingAndConverting
	ConvertExcelWorkbookToPDF_A_CompliedFiles()
	ConvertExcelWorkbookToPDF_DirectConversion()
	ConvertExcelWorkbookToPDF_SetPDFCreationTime()
	ConvertingWorksheetToImage_PNG()
	ConvertingWorksheetToImage_SVG()
	ManagingDocumentProperties()
	OpeningExcelFileUsingPath()
	OpeningExcelFileUsingStream()
	ReadAndWriteCSVFileFormat()
	ReadAndWriteTabDelimitedFileFormat()
	ReadAndWriteXLSBFileFormat()
	ReadAndWriteXLSMFileFormat()
	SavingFiletoSomeLocation()
	SavingFiletoStream()

	// PivotTables
	CreatePivotTable()
	ManipulatePivotTable()

	// RowsAndColumns
	SettingHeightOfRow()
	SettingHeightOfAllRowsInWorksheet()
	SettingWidthOfColumn()
	SettingWidthOfAllColumnsInWorksheet()
	CopyingRows()
	CopyingColumns()
	GroupingRowsColumns()
	UnGroupingRowsColumns()
	InsertRow()
	InsertingMultipleRows()
	DeletingMultipleRows()
	InsertColumn()
	DeleteColumn_()

	// TechnicalArticles
	ApplyCustomThemeColorsOfWorkbookUsingArrayOfColors()
	CopyThemeFromOneWorkbookToAnother()
	CreateAndManipulateExcelTable()
	GroupRowsAndColumnsOfWorksheet()

	// Worksheets
	CopyWorksheetsWithinWorkbook()
	MoveWorksheetsWithinWorkbook()
	AddingWorksheetsToNewExcelFile()
	AccessingWorksheetsUsingSheetIndex()
	RemovingWorksheetsUsingSheetIndex()
	AddingPageBreaks()
	EnablingPageBreakPreview()
	ZoomFactor()
	FreezePanes()
	SplitPanes()
	RemovingPanes()

	//Stop before exiting
	fmt.Println("\nProgram Finished. Press Enter to Exit....")
}
