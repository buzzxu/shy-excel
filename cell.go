package shy_excel

import (
	"github.com/xuri/excelize/v2"
	"strconv"
)

func newCell(f *excelize.File, sheet *Sheet, startRow, colCount int) error {
	//empty data return
	if sheet.Data == nil || len(*sheet.Data) == 0 {
		return nil
	}
	columns := sheet.Header.Columns
	for i, data := range *sheet.Data {
		colStart := 0
		col := columns[colStart]
		if col.Type == TypeIndex {
			//如果第一列是索引
			f.SetCellInt(sheet.Name, axis(startRow, colStart+1), i+1)
			i++
			colStart++
		}
		for ; colStart < colCount; colStart++ {
			col = columns[colStart]
			cell := axis(startRow, colStart+1)
			var err error
			switch col.Type {
			case TypeIndex:
				err = f.SetCellInt(sheet.Name, cell, i)
			case TypeHyperLink:
				value := data[col.Name]
				v := value.(string)
				err = f.SetCellHyperLink(sheet.Name, cell, v, "External")
				err = f.SetCellStyle(sheet.Name, cell, cell, defStyle(DefStyleKeys_Link, f))
				err = f.SetCellStr(sheet.Name, cell, v)
			default:
				err = f.SetCellValue(sheet.Name, cell, data[col.Name])
			}
			if err != nil {
				return err
			}

		}
		startRow++
	}
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
