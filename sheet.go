package shyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func newSheet(f *excelize.File, sheetIndex int, sheet *Sheet, consumer func(int, int)) error {
	if sheet.Name == "" {
		sheet.Name = "Sheet1"
	}
	_, err := f.NewSheet(sheet.Name)
	if err != nil {
		fmt.Println(err)
		return err
	}
	//是否隐藏
	if !sheet.Visible {
		err := f.SetSheetVisible(sheet.Name, false)
		if err != nil {
			fmt.Println(err)
			return err
		}
	}

	//冻结
	if sheet.Header.FreezeCol != 0 {
		//err := f.SetPanes(sheet.Name, &excelize.Panes{XSplit: sheet.Header.FreezeCol, YSplit: 0, TopLeftCell: "A1"})
		//if err != nil {
		//	fmt.Println(err)
		//	return err
		//}
	}
	layout(f, sheet)
	margins(f, sheet)
	startRow, colCount, err := newHeader(f, sheet)
	if err != nil {
		return err
	}
	err = newRows(f, sheetIndex, sheet, startRow, colCount, consumer)
	if err != nil {
		return err
	}
	return nil
}

func layout(f *excelize.File, sheet *Sheet) {
	if sheet.Layout != nil {
		var pageLayout = &excelize.PageLayoutOptions{}
		if sheet.Layout.Size != nil {
			pageLayout.Size = sheet.Layout.Size
		}
		if sheet.Layout.Orientation != nil {
			pageLayout.Orientation = sheet.Layout.Orientation
		}
		if sheet.Layout.FirstPageNumber != nil {
			pageLayout.FirstPageNumber = sheet.Layout.FirstPageNumber
		}
		if sheet.Layout.AdjustTo != nil {
			pageLayout.AdjustTo = sheet.Layout.AdjustTo
		}
		if sheet.Layout.FitToHeight != nil {
			pageLayout.FitToHeight = sheet.Layout.FitToHeight
		}
		if sheet.Layout.FitToWidth != nil {
			pageLayout.FitToWidth = sheet.Layout.FitToWidth
		}
		if sheet.Layout.BlackAndWhite != nil {
			pageLayout.BlackAndWhite = sheet.Layout.BlackAndWhite
		}
		if err := f.SetPageLayout(sheet.Name, pageLayout); err != nil {
			fmt.Println(err)
			return
		}
	}
}

func margins(f *excelize.File, sheet *Sheet) {
	if sheet.Margins != nil {
		var pageMargins = &excelize.PageLayoutMarginsOptions{}
		if sheet.Margins.Bottom != nil {
			pageMargins.Bottom = sheet.Margins.Bottom
		}
		if sheet.Margins.Footer != nil {
			pageMargins.Footer = sheet.Margins.Footer
		}
		if sheet.Margins.Header != nil {
			pageMargins.Header = sheet.Margins.Header
		}
		if sheet.Margins.Left != nil {
			pageMargins.Left = sheet.Margins.Left
		}
		if sheet.Margins.Right != nil {
			pageMargins.Right = sheet.Margins.Right
		}
		if sheet.Margins.Top != nil {
			pageMargins.Top = sheet.Margins.Top
		}
		if sheet.Margins.Horizontally != nil {
			pageMargins.Horizontally = sheet.Margins.Horizontally
		}
		if sheet.Margins.Vertically != nil {
			pageMargins.Vertically = sheet.Margins.Vertically
		}
		if err := f.SetPageMargins(sheet.Name, pageMargins); err != nil {
			fmt.Println(err)
			return
		}
	}
}
