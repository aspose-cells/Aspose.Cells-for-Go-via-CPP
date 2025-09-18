package Data

import (
	"fmt"
	. "main/Common"
	"strings"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// FindOrSearchData function to find or search data in the Excel file
func FindOrSearchData() error {
	// Source directory path
	dirPath := "..\\Data\\Data\\"

	// Path of input excel file
	sampleFindOrSearchData := dirPath + "sampleFindOrSearchData.xlsx"

	// Load sample excel file into a workbook object
	wb, err := NewWorkbook_String(sampleFindOrSearchData)
	if err != nil {
		return err
	}

	// Get first worksheet of the workbook
	wss, err := wb.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	cells, err := ws.GetCells()
	enCell, err := cells.GetEnumerator()
	if err != nil {
		return err
	}

	for {
		loop, _ := enCell.MoveNext()
		if !loop {
			break
		}
		cell, err := enCell.GetCurrent()
		if err != nil {
			return err
		}
		fmlVal, err := cell.GetFormula()
		if err != nil {
			return err
		}

		// Finding the cell containing the specified formula

		if fmlVal == "=SUM(A5:A10)" {
			cellName, err := cell.GetName()
			if err != nil {
				return err
			}
			fmt.Printf("Name of the cell containing formula =SUM(A5:A10): %s\n", cellName)
		} else if strings.Index(fmlVal, "CHA") > -1 {
			cellName, err := cell.GetName()
			if err != nil {
				return err
			}
			fmt.Printf("Name of the cell containing the formula that contains CHA: %s\n", cellName)
		} else {
			cellType, err := cell.GetType()
			if err != nil {
				return err
			}

			switch cellType {
			case CellValueType_IsString:
				strVal, err := cell.GetStringValue()
				if err != nil {
					return err
				}
				if strings.Index(strVal, "SampleData") > -1 {
					cellName, err := cell.GetName()
					if err != nil {
						return err
					}
					fmt.Printf("Name of the cell containing specified string: %s\n", cellName)
				} else if strings.Index(strVal, "Two") > -1 {
					cellName, err := cell.GetName()
					if err != nil {
						return err
					}
					fmt.Printf("Name of the cell containing the string that contains Two: %s\n", cellName)
				} else if strings.Index(strVal, "AAA") == 0 {
					cellName, err := cell.GetName()
					if err != nil {
						return err
					}
					fmt.Printf("Name of the cell containing the string that starts with AAA: %s\n", cellName)
				} else if strings.Index(strVal, "BBB") == len(strVal)-3 {
					cellName, err := cell.GetName()
					if err != nil {
						return err
					}
					fmt.Printf("Name of the cell containing the string that ends with BBB: %s\n", cellName)
				}
			case CellValueType_IsNumeric:
				intValue, err := cell.GetIntValue()
				if err != nil {
					return err
				}
				if intValue == 80 {
					cellName, err := cell.GetName()
					if err != nil {
						return err
					}
					fmt.Printf("Name of the cell containing the number 80: %s\n", cellName)
				}
			}
		}
	}

	// Show successful execution message on console
	ShowMessageOnConsole("FindOrSearchData executed successfully.\r\n\r\n")
	return err
}
