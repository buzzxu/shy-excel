package shyexcel

import (
	"fmt"
	"testing"
)

func Test_fromFile(t *testing.T) {

	table := fromFile("./example.json")
	if table != nil && len(*table) > 0 {
		for _, sheet := range *table {
			fmt.Printf("%d--->%s", sheet.Header.Count(), sheet.Header.Title)

		}
	}

}
