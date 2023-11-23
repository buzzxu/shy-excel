package shyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

func newRows(f *excelize.File, sheet *Sheet, startRow, keysCount int) error {
	//empty data return
	if sheet.Data == nil || len(*sheet.Data) == 0 {
		return nil
	}

	colCount := len(sheet.Header.Columns)
	columns := sheet.Header.Columns
	merge := sheet.Header.dept != nil || *sheet.Header.dept > 0
	var err error
	for index, rowData := range *sheet.Data {
		var startCol = 0
		startRow, err = newRow(&rowOption{
			f:         f,
			sheetName: sheet.Name,
			columns:   columns,
			rowData:   rowData,
			merge:     merge,
			index:     index,
			startRow:  startRow,
			startCol:  startCol,
			colCount:  colCount,
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
		if col.dept != nil {
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
						startCol:  startCol,
						colCount:  len(col.Columns),
					})
					if _index != _rowCount-1 {
						startRow++
					}
				}
			}
		} else {
			cell := axis(startRow, option.startCol+1)
			if col.Type == TypeIndex {
				err = option.f.SetCellInt(option.sheetName, cell, option.index+1)
			} else {
				value := option.rowData[col.Name]
				if value == nil {
					err = option.f.SetCellValue(option.sheetName, cell, "")
				} else {
					switch col.Type {
					case TypeHyperLink:
						v := value.(string)
						err = option.f.SetCellHyperLink(option.sheetName, cell, v, "External")
						err = option.f.SetCellStyle(option.sheetName, cell, cell, defStyle(DefStyleKeys_Link, option.f))
						err = option.f.SetCellStr(option.sheetName, cell, v)
					case TypeImage:
						//todo
					default:
						err = option.f.SetCellValue(option.sheetName, cell, value)
					}
				}
			}

			if option.merge {
				//如果需要合并 获取当前行数据中 集合的最大数量
				//todo 未处理多级
				if option._dept() > 0 {
					vCell := axis(startRow+(option.dept-1), option.startCol+1)
					option.f.MergeCell(option.sheetName, cell, vCell)
					err = option.f.SetCellStyle(option.sheetName, cell, vCell, defStyle(DefStyleKeys_Merge_ROW, option.f))
				}
			}
			option.startCol++
		}
	}
	if err != nil {
		return startRow, err
	}
	option.startCol = 0
	return startRow, nil
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
