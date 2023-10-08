package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func NewTable(table *Table) *excelize.File {
	f := excelize.NewFile()
	active := false
	for _, sheet := range *table {
		err := newSheet(f, sheet)
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
	if (len(*table) == 1 && (*table)[0].Name != "Sheet1") || (len(*table) > 1 && !active) {
		//默认取第一个sheet
		setActiveSheet(f, (*table)[0].Name)
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
