package excel

import (
	"fmt"
	"testing"
)

func Test_NewTable(t *testing.T) {

	table := fromFile("./example.json")
	f := NewTable(table)
	err := f.SaveAs("example.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable1(t *testing.T) {
	table := fromFile("./example1.json")
	f := NewTable(table)
	err := f.SaveAs("example1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func Test_NewTable2(t *testing.T) {

	table := fromFile("./example2.json")
	f := NewTable(table)
	err := f.SaveAs("example2.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
