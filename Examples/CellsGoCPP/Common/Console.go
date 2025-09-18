package Common

import (
	"fmt"
	"os"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// / <summary>
// / Get file content.
// / </summary>
// / <param name="file">source file</param>
// / <returns>Returns a Vector<byte> object.</returns>
func GetDataFromFile(file string) ([]byte, error) {
	// open a file
	fileStream, err := os.Open(file)
	if err != nil {
		fmt.Println("Failed to open the file.")
		return nil, err
	}
	defer fileStream.Close()

	// Get file size
	fileInfo, err := fileStream.Stat()
	if err != nil {
		return nil, err
	}

	fileSize := fileInfo.Size()
	buffer := make([]byte, fileSize)

	// Read file contents into byte array
	_, err = fileStream.Read(buffer)
	if err != nil {
		return nil, err
	}

	data := make([]byte, fileSize)
	copy(data, buffer)

	return data, nil
}

func SaveDataToFile(data []byte, file string) error {
	// open a file
	fileStream, err := os.Create(file)
	if err != nil {
		fmt.Println("Failed to open the file.")
		return err
	}
	defer fileStream.Close()

	_, err = fileStream.Write(data)
	if err != nil {
		return err
	}

	return nil
}

func ShowMessageOnConsole(msg string) {
	fmt.Println(msg)
}

func ShowCellsVersion() {
	version, _ := CellsHelper_GetVersion()
	fmt.Println("Aspose.Cells for Go Version: " + version)
	fmt.Println("\n\n")
}
