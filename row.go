package shyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

func newRows(f *excelize.File, sheetIndex int, sheet *Sheet, startRow, keysCount int, consumer func(int, int)) error {
	//empty data return
	if sheet.Data == nil || len(*sheet.Data) == 0 {
		return nil
	}

	colCount := len(sheet.Header.Columns)
	columns := sheet.Header.Columns
	merge := sheet.Header.dept != nil || *sheet.Header.dept > 0

	var err error
	for rowIndex, rowData := range *sheet.Data {
		maxRow := calcMaxRow(columns, rowData)
		var startCol = 0
		consumer(sheetIndex, rowIndex)
		startRow, err = newRow(&rowOption{
			f:         f,
			sheetName: sheet.Name,
			columns:   columns,
			rowData:   rowData,
			merge:     merge,
			index:     rowIndex,
			startRow:  startRow,
			startCol:  startCol,
			colCount:  colCount,
			rows:      maxRow,
		})
		startRow++
	}
	return err
}
func newRow(option *rowOption) (int, error) {
	startRow := option.startRow
	var err error
	var startCol = 0
	for ; startCol < option.colCount; startCol++ {
		col := option._col(startCol)
		//判断当前列是否有子集合
		if col.Collection {
			if arrays, ok := option.rowData[col.Name].([]interface{}); ok {
				_rowCount := len(arrays)
				for _index, item := range arrays {
					itemMap, ok := item.(map[string]interface{})
					if !ok {
						return startRow, fmt.Errorf("row %d,item %s is not of type map[string]interface{}", startRow, col.Name)
					}
					startRow, err = newRow(&rowOption{
						f:         option.f,
						sheetName: option.sheetName,
						columns:   col.Columns,
						rowData:   itemMap,
						index:     _index,
						startRow:  startRow,
						startCol:  option.startCol,
						colCount:  len(col.Columns),
						rows:      itemMap["__rows__"].(int),
					})
					if err != nil {
						return startRow, err
					}
					if _index != _rowCount-1 {
						startRow++
					}
				}
			}
		} else {
			// 处理普通列
			startRow, err = handleColumn(option, col, startRow, option.startCol, startCol)
			if err != nil {
				return startRow, err
			}
		}
	}

	return startRow, nil
}

func handleColumn(option *rowOption, column *Column, startRow, startCol int, colIndex int) (int, error) {
	col := option._col(colIndex)
	cell := axis(startRow, startCol+1)

	if col.Type == TypeIndex {
		err := option.f.SetCellInt(option.sheetName, cell, option.index+1)
		if err != nil {
			return startRow, err
		}
	} else {
		value := option.rowData[col.Name]
		if value == nil {
			err := option.f.SetCellValue(option.sheetName, cell, "")
			if err != nil {
				return startRow, err
			}
		} else {
			err := setCellValue(option.f, option.sheetName, cell, value, col.Type)
			if err != nil {
				return startRow, err
			}
		}
		//设置单元格样式
		if col.cellStyleId > 0 {
			err := setCellStyle(option.f, option.sheetName, cell, col.cellStyleId)
			if err != nil {
				return startRow, err
			}
		}
	}

	if option.rows > 0 {
		vCell := axis(startRow+(option.rows-1), option.startCol+1)
		option.f.MergeCell(option.sheetName, cell, vCell)
		err := option.f.SetCellStyle(option.sheetName, cell, vCell, defStyle(DefStyleKeys_Merge_ROW, option.f))
		if err != nil {
			return startRow, err
		}
	}
	option.startCol = option.startCol + 1

	return startRow, nil
}

func setCellValue(f *excelize.File, sheetName, cell string, value interface{}, colType Type) error {
	switch colType {
	case TypeIndex:
		return f.SetCellInt(sheetName, cell, value.(int))
	case TypeHyperLink:
		v := fmt.Sprintf("%v", value)
		if err := f.SetCellHyperLink(sheetName, cell, v, "External"); err != nil {
			return err
		}
		err := f.SetCellStyle(sheetName, cell, cell, defStyle(DefStyleKeys_Link, f))
		if err != nil {
			return err
		}
		err = f.SetCellStr(sheetName, cell, v)
		if err != nil {
			return err
		}
		return f.SetCellValue(sheetName, cell, v)
	default:
		return f.SetCellValue(sheetName, cell, value)
	}
}

func setCellStyle(f *excelize.File, sheetName, cell string, style int) error {
	return f.SetCellStyle(sheetName, cell, cell, style)
}

func axis(rowN, colN int) string {
	colName, err := excelize.ColumnNumberToName(colN)
	if err != nil {
		panic(err)
	}
	return colName + strconv.Itoa(rowN)
}

type rowOption struct {
	f                                   *excelize.File
	sheetName                           string
	columns                             []*Column
	rowData                             map[string]interface{}
	merge                               bool
	index, colCount, startRow, startCol int
	dept                                int //是否包含集合以及深度
	rows                                int //当前行数
}
type dataOption struct {
	f                                   *excelize.File
	sheetName                           string
	columns                             []*Column
	col                                 *Column
	rowData                             []map[string]interface{}
	index, colCount, startRow, startCol int
}

func (option *rowOption) _col(index int) *Column {
	return option.columns[index]
}
func (option *rowOption) _dept() int {
	if option.dept > 0 {
		return option.dept
	}
	for _, data := range option.rowData {
		switch v := data.(type) {
		case []interface{}:
			option.dept = len(v)
			break
		default:
			continue
		}
	}
	if option.dept == 0 {
		option.dept = -1
	}
	return option.dept
}

func (option *dataOption) _col(index int) *Column {
	return option.columns[index]
}

func _toMaps(key string, rowData map[string]interface{}) ([]map[string]interface{}, error) {
	if arrays, ok := rowData[key].([]interface{}); ok {
		result := make([]map[string]interface{}, len(arrays))
		for i, item := range arrays {
			itemMap, ok := item.(map[string]interface{})
			if !ok {
				return nil, fmt.Errorf("item %s is not of type map[string]interface{}", key)
			}
			result[i] = itemMap
		}
		return result, nil
	}
	return nil, fmt.Errorf("item %s is not of type map[string]interface{}", key)
}
