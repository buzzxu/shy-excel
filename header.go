package shyexcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
)

// 表头
func newHeader(f *excelize.File, sheet *Sheet) (int, int, error) {
	var columns = sheet.Header.Columns
	var columnLength = sheet.Header.Count()
	if columnLength == 0 {
		return 0, 0, errors.New("not found column headers")
	}
	//起始位置
	start := 1

	var fontFamily = "Calibri"
	if sheet.Header.FontFamily != nil {
		fontFamily = *sheet.Header.FontFamily
	}
	//表头样式
	headerStyle := newStyleHeader(f, sheet, fontFamily)

	//标题
	if &sheet.Header.Title != nil {
		headerStyleLevel1 := newStyleHeaderLever1(f, sheet, fontFamily)
		hCell := "A1"
		vCell := axis(start, columnLength)
		if err := f.SetCellStyle(sheet.Name, hCell, vCell, headerStyleLevel1); err != nil {
			return 0, 0, err
		}
		if err := f.SetCellStr(sheet.Name, hCell, sheet.Header.Title); err != nil {
			return 0, 0, err
		}
		if err := f.MergeCell(sheet.Name, hCell, vCell); err != nil {
			return 0, 0, err
		}
		start++
	}

	//列头
	startCol := 1
	dept := sheet.Header.Depth()
	setColumnTitle(f, sheet.Name, headerStyle, dept, columns, start, startCol)

	if _autoFilter := sheet.Header.AutoFilter != nil; _autoFilter {
		x, _ := excelize.CoordinatesToCellName(1, start)
		y, _ := excelize.CoordinatesToCellName(columnLength, start)
		if err := f.AutoFilter(sheet.Name, x+":"+y, []excelize.AutoFilterOptions{}); err != nil {
			return 0, 0, err
		}
	}
	if sheet.Header.Height > 0 {
		err := f.SetRowHeight(sheet.Name, start, sheet.Header.Height)
		if err != nil {
			return 0, 0, err
		}
	}

	return (start + 1) + dept, columnLength, nil
}

func setColumnTitle(f *excelize.File, sheetName string, headerStyle int, dept int, columns []*Column, startRow int, startCol int) int {
	colIndex := startCol
	for _, column := range columns {
		if column.Width > 0 {
			colN, _ := excelize.ColumnNumberToName(colIndex)
			if err := f.SetColWidth(sheetName, colN, colN, column.Width); err != nil {
				fmt.Println(err)
			}
		}
		cell := axis(startRow, colIndex)
		//设置单元格样式
		column.cellStyleId = newStyleWithColumn(f, column)
		if err := f.SetCellStr(sheetName, cell, column.Title); err != nil {
			fmt.Println(err)
		}
		if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
			fmt.Println(err)
		}
		if len(column.Columns) > 0 {
			colIndex += setColumnTitle(f, sheetName, headerStyle, column.Depth(), column.Columns, startRow+1, colIndex)
			vCell := axis(startRow, colIndex-1)
			f.MergeCell(sheetName, cell, vCell)
			if err := f.SetCellStyle(sheetName, cell, vCell, headerStyle); err != nil {
				fmt.Println(err)
			}
		} else {
			if column.Merge || dept > 0 {
				vCell := axis(startRow+dept, colIndex)
				f.MergeCell(sheetName, cell, vCell)
				if err := f.SetCellStyle(sheetName, cell, vCell, headerStyle); err != nil {
					fmt.Println(err)
				}
			}
			colIndex++
		}
	}
	return colIndex - startCol
}

// 标题样式
func newStyleHeaderLever1(f *excelize.File, sheet *Sheet, fontFamily string) int {
	var headerStyleLevel1 int
	if sheet.Header.Style != nil && sheet.Header.Style.Border != nil && sheet.Header.Style.Fill != nil && sheet.Header.Style.Font != nil {
		headerStyleLevel1, _ = newStyle(f, sheet.Header.Style)
	} else {
		var fontTitleSize float64
		if sheet.Header.FontTitleSize != nil {
			fontTitleSize = *sheet.Header.FontTitleSize
		} else {
			fontTitleSize = 14
		}
		headerStyleLevel1 = defStyleHeader(f, &excelize.Font{
			Bold:   true,
			Size:   fontTitleSize,
			Family: fontFamily,
		})
	}
	return headerStyleLevel1
}

// 表头样式
func newStyleHeader(f *excelize.File, sheet *Sheet, fontFamily string) int {
	var fontHeaderSize float64
	if sheet.Header.FontHeaderSize != nil {
		fontHeaderSize = *sheet.Header.FontHeaderSize
	} else {
		//default font size
		fontHeaderSize = 11
	}
	return defStyleHeader(f, &excelize.Font{
		Bold:   true,
		Size:   fontHeaderSize,
		Family: fontFamily,
	})
}
