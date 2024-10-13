package shyexcel

import (
	"fmt"
	"testing"
)

func Test_NewTable(t *testing.T) {

	table, err := File("./examples/example.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable1(t *testing.T) {
	table, err := File("./examples/example1.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable2(t *testing.T) {

	table, err := File("./examples/example2.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example2.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable3(t *testing.T) {

	table, err := File("./examples/example3.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d ", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example3.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable4(t *testing.T) {

	table, err := File("./examples/example4.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d \n", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example4.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable5(t *testing.T) {

	table, err := File("./examples/example5.json")
	f := NewTable(table, func(sheetIndex, rowIndex int) {
		fmt.Printf("当前sheet: %d,当前行数: %d \n", sheetIndex, rowIndex)
	})
	err = f.SaveAs("./examples/example5.xlsx")
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
