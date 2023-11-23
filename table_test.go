package shyexcel

import (
	"fmt"
	"testing"
)

func Test_NewTable(t *testing.T) {

	table, err := File("./example.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable1(t *testing.T) {
	table, err := File("./example1.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable2(t *testing.T) {

	table, err := File("./example2.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example2.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable3(t *testing.T) {

	table, err := File("./example3.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example3.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable4(t *testing.T) {

	table, err := File("./example4.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d \n", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example4.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func TestNewHTTP(t *testing.T) {
	f, err := NewHTTP("https://file.xw-jd.com/static/shy-excel/example.json", "GET", JSON, nil, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("example_http.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
