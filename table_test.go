package shy_excel

import (
	"fmt"
	"testing"
)

func Test_NewTable(t *testing.T) {

	table := fromFile("./example.json")
	f := NewTable(table)
	err := f.SaveAs("xx.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
