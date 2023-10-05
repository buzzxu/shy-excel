package shy_excel

import (
	"github.com/xuri/excelize/v2"
	"strconv"
)

func newCell(f *excelize.File, sheetName string, sheet *Sheet) error {

	return nil
}

func data(f *excelize.File, sheetName string, sheet *Sheet) {

}

func axis(rowN, colN int) string {
	colName, err := excelize.ColumnNumberToName(colN)
	if err != nil {
		panic(err)
	}
	return colName + strconv.Itoa(rowN)
}
