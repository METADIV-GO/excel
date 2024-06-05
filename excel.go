package excel

import (
	"bytes"
	"log"
	"strconv"

	excel "github.com/xuri/excelize/v2"
)

/*
NewExcelFile creates a new excel file
*/
func NewExcelFile() *excel.File {
	return excel.NewFile()
}

/*
ReadExcelFile reads the file from the given bytes
*/
func ReadExcelFile(b []byte) *excel.File {
	file, err := excel.OpenReader(bytes.NewReader(b))
	if err != nil {
		log.Println(err)
		return nil
	}
	return file
}

/*
AddExcelSheet adds a new sheet to the file with the given data
*/
func AddExcelSheet[T any](file *excel.File, sheetName string, data []T) *excel.File {
	if file.GetSheetName(0) == "Sheet1" {
		file.SetSheetName("Sheet1", sheetName)
	} else {
		file.NewSheet(sheetName)
	}

	rows := ToCSV[T](data)

	var row = 1
	for i := range rows {
		for j := range rows[i] {
			file.SetCellValue(sheetName, getColumnName(j+1)+strconv.Itoa(row), rows[i][j])
		}
		row++
	}

	return file
}

/*
ReadExcelSheet reads the data from the sheet
*/
func ReadExcelSheet[T any](file *excel.File, sheetName string) []T {
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Println(err)
		return nil
	}
	return FromCSV[T](rows)
}

/*
ExcelToBytes converts the file to bytes
*/
func ExcelToBytes(file *excel.File) []byte {
	b, _ := file.WriteToBuffer()
	return b.Bytes()
}

func getColumnName(col int) string {
	name := make([]byte, 0, 3) // max 16,384 columns (2022)
	const aLen = 'Z' - 'A' + 1 // alphabet length
	for ; col > 0; col /= aLen + 1 {
		name = append(name, byte('A'+(col-1)%aLen))
	}
	for i, j := 0, len(name)-1; i < j; i, j = i+1, j-1 {
		name[i], name[j] = name[j], name[i]
	}
	return string(name)
}
