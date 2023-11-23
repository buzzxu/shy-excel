package shyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"net/http"
)

// JsonToTable 通过JSON数据生成表格
func JsonToTable(json string, consumer func(int, map[string]interface{})) (*excelize.File, error) {
	table, err := Json([]byte(json))
	if err != nil {
		return nil, err
	}
	return NewTable(table, consumer), nil
}

// NewHTTP 通过Http请求生成表格
func NewHTTP(url, method string, responseType ResponseType, funcHeader func(header http.Header), consumer func(int, map[string]interface{})) (*excelize.File, error) {
	table, err := Http(url, method, responseType, funcHeader)
	if err != nil {
		return nil, err
	}
	return NewTable(table, consumer), nil
}

func NewTable(table *Table, consumer func(int, map[string]interface{})) *excelize.File {
	f := excelize.NewFile()
	active := false
	sheets := table.Sheets
	for _, sheet := range sheets {
		err := newSheet(f, sheet, consumer)
		if err != nil {
			fmt.Println(err)
			return nil
		}
		if sheet.Active {
			index, err := f.GetSheetIndex(sheet.Name)
			if err != nil {
				fmt.Println(err)
				return nil
			}
			f.SetActiveSheet(index)
			active = true
		}
	}
	if (len(sheets) == 1 && (sheets)[0].Name != "Sheet1") || (len(sheets) > 1 && !active) {
		//默认取第一个sheet
		setActiveSheet(f, (sheets)[0].Name)
		f.DeleteSheet("Sheet1")
	}
	return f
}

func setActiveSheet(f *excelize.File, sheet string) error {
	index, err := f.GetSheetIndex(sheet)
	if err != nil {
		fmt.Println(err)
		return err
	}
	f.SetActiveSheet(index)
	return nil
}
