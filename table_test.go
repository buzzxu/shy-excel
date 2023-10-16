package shyexcel

import (
	"fmt"
	"testing"
)

func Test_NewTable(t *testing.T) {

	table, err := File("./example.json")
	f := NewTable(table)
	err = f.SaveAs("example.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable1(t *testing.T) {
	table, err := File("./example1.json")
	f := NewTable(table)
	err = f.SaveAs("example1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable2(t *testing.T) {

	table, err := File("./example2.json")
	f := NewTable(table)
	err = f.SaveAs("example2.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func TestNewHTTP(t *testing.T) {
	f, err := NewHTTP("https://file.xw-jd.com/static/shy-excel/example.json", "GET", JSON, nil)
	err = f.SaveAs("example_http.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
