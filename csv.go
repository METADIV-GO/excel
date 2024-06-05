package excel

import (
	"bufio"
	"bytes"
	"strings"

	"github.com/gocarina/gocsv"
)

/*
ToCSV converts the data to csv
*/
func ToCSV[T any](data []T) [][]string {
	var b bytes.Buffer
	writer := bufio.NewWriter(&b)
	gocsv.Marshal(data, writer)
	writer.Flush()
	var result = make([][]string, 0)

	lines := strings.Split(b.String(), "\n")
	for i := range lines {
		values := strings.Split(lines[i], ",")
		result = append(result, values)
	}
	return result
}

/*
FromCSV converts the csv to data
*/
func FromCSV[T any](csv [][]string) []T {
	var data []T
	var b bytes.Buffer
	for i := range csv {
		b.WriteString(strings.Join(csv[i], ","))
		b.WriteString("\n")
	}
	reader := bufio.NewReader(&b)
	gocsv.Unmarshal(reader, &data)
	return data
}

/*
CSVToBytes converts the csv to bytes
*/
func CSVToBytes(csv [][]string) []byte {
	var b bytes.Buffer
	for i := range csv {
		b.WriteString(strings.Join(csv[i], ","))
		b.WriteString("\n")
	}
	return b.Bytes()
}

/*
BytesToCSV converts the bytes to csv
*/
func BytesToCSV(data []byte) [][]string {
	var result = make([][]string, 0)
	lines := strings.Split(string(data), "\n")
	for i := range lines {
		values := strings.Split(lines[i], ",")
		result = append(result, values)
	}
	return result
}
