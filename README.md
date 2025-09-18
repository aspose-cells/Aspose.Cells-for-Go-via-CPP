![](https://img.shields.io/badge/aspose.cells%20for%20Go%20via%20C++-v25.9.0-green?style=for-the-badge&logo=go) [![Product Page](https://img.shields.io/badge/Product-0288d1?style=for-the-badge&logo=Google-Chrome&logoColor=white)](https://products.aspose.com/cells/go-cpp/) [![Documentation](https://img.shields.io/badge/Documentation-388e3c?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.aspose.com/cells/go-cpp/) [![API Ref](https://img.shields.io/badge/Reference-f39c12?style=for-the-badge&logo=html5&logoColor=white)](https://reference.aspose.com/cells/go-cpp/)  [![Blog](https://img.shields.io/badge/Blog-d32f2f?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.aspose.com/categories/aspose.cells-product-family/) [![Support](https://img.shields.io/badge/Support-7b1fa2?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.aspose.com/c/cells/9) ![GitHub commits since latest release (by date)](https://img.shields.io/github/commits-since/aspose-cells/aspose-cells-go-cpp/v25.9.0?style=for-the-badge)

[Aspose.Cells for Go via C++](https://products.aspose.com/cells/go-cpp) is a cross-platform native assembly that can be deployed simply by copying it. You can use it to develop 64-bit applications in any development environment that supports Go, such as, Microsoft Visual Code.  You don't have to worry about other services or modules. It supports Excel 97-2003 (XLS), Excel 2007-2013/2016/365 (XLSX, XLSM, XLSB), OpenOffice XML,  LibreOffice (ODS), and other formats such as CSV, TSV, and more.

Aspose.Cells for Go via C++ Library allows you to work with the built-in as well as the custom document properties in Microsoft Excel. Supports high-quality conversion of Excel Workbooks to PDF/A compliant files. Work with formulas, pivot tables, worksheets, tables, ranges, charts, OLE objects, and much more.

Aspose.Cells for Go via C++ Library  allows the developers to work with spreadsheet rows, columns, data, formulas, pivot tables, worksheets, tables, charts, and drawing objects from their own Go applications.

<p align="center">
  <a title="Download ZIP" href="https://github.com/aspose-cells/Aspose.Cells-for-Go-via-CPP/archive/refs/heads/main.zip">
    <img src="http://i.imgur.com/hwNhrGZ.png" />
  </a>
</p>

Directory | Description
--------- | -----------
[Examples](Examples)  | A collection of GoLang examples that help you learn and explore the API features.

## Excel File Processing Features

- Load existing spreadsheets or create one from scratch.
- Convert spreadsheets to any [supportted file format](https://docs.aspose.com/cells/go-cpp/supported-file-formats/).
- [Convert worksheets to different image formats](https://docs.aspose.com/cells/go-cpp/converting-worksheet-to-different-image-formats/).
- [Apply conditional formatting](https://docs.aspose.com/cells/go-cpp/apply-conditional-formatting-in-worksheet/) as per your choice.
-- [Manipulate Pivot Tables](https://docs.aspose.com/cells/go-cpp/manipulate-pivot-table/) in your spreadsheets.
- [Convert table to range](https://docs.aspose.com/cells/go-cpp/tables-and-ranges/) without losing formatting.
- Fetch a cell's name by providing the row and column index, similarly, fetch row and column index of cell from its name.
-- [Create & customize Excel charts](https://docs.aspose.com/cells/go-cpp/creating-and-customizing-charts/).
- [Render charts as images & PDF](https://docs.aspose.com/cells/go-cpp/chart-rendering/).

## Read & Write Spreadsheets

**Microsoft Excel:** XLS, XLSX, XLSB\
**Text:** CSV, TSV\
**OpenDocument:** ODS\
**Others:** HTML, MHTML

## Save Spreadsheets As

**Microsoft Excel:** XLSM, XLTX, XLTM, XLAM\
**Fixed Layout:** PDF, XPS\
**Images:** JPEG, PNG, BMP, GIF, EMF, SVG

## Quick Start Guide with Aspose.Cells for Go via C++

<a id="installationinyourproject"></a>

### Installation Aspose.Cells for Go via C++ package and running your code in your project

1. Create a directory for your project and a main.go file within. Add the following code to your main.go.

```Go

package main

import (
 . "github.com/aspose-cells/aspose-cells-go-cpp/v25"
 "fmt"
)

func main() {
 lic, _ := NewLicense()
 lic.SetLicense_String("YOUR_LICENSE_File_PATH")
 workbook, _ := NewWorkbook()
 worksheets, _ := workbook.GetWorksheets()
 worksheet, _ := worksheets.Get_Int(0)
 cells, _ := worksheet.GetCells()
 cell, _ := cells.Get_String("A1")
 cell.PutValue_String_Bool("Hello World!", true)
 style, _ := cell.GetStyle()
 style.SetPattern(BackgroundType_Solid)
 color, _ := NewColor()
 color.Set_Color_R(uint8(255))
 color.Set_Color_G(uint8(128))
 style.SetForegroundColor(color)
 cell.SetStyle_Style(style)
 workbook.Save_String("HELLO.pdf")

}

```

1. Initialize project go.mod

```bash

go mod init main

```

1. Fetch the dependencies for your project.

```bash

go mod tidy

```

If Aspose.Cells for Go via C++ is not installed in the development environment, Go will automatically install the package in the `$GOPATH` subdirectory.

1. Set your PATH to point to the shared libraries in Aspose.Cells for Go via C++ in your current command shell. Replace your_version with the version of Aspose.Cells for Go via C++ you are running.

```cmd

set PATH=%PATH%;%GOPATH%\github.com\aspose-cells\aspose-cells-go-cpp\v25@v25.9.0\lib\win_x86_64\

```

Or in your powershell

```powershell

$env:Path = $env:Path+ ";${env:GOPATH}\github.com\aspose-cells\aspose-cells-go-cpp\v25@v25.9.0\lib\win_x86_64\"

```

Or in your linux bash

```bash
export PATH=$PATH:$GOPATH/github.com/aspose-cells/aspose-cells-go-cpp/v25@v25.9.0/lib/linux_x86_64/

```

You may also copy the DLL files from the above path to the same place as your project executable.

1. Run your created application.

```bash

go run main.go

```
