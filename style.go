package shyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type DefStyleKeys int16

const (
	DefStyleKeys_Link      DefStyleKeys = 3
	DefStyleKeys_Merge_ROW DefStyleKeys = 4
	DefStyleKeys_WRAP_TEXT DefStyleKeys = 5
)

var DEF_STYLE map[DefStyleKeys]int

var DEF_STYLE_BORDER = []excelize.Border{
	{Type: "left", Color: "000000", Style: 1},
	{Type: "top", Color: "000000", Style: 1},
	{Type: "bottom", Color: "000000", Style: 1},
	{Type: "right", Color: "000000", Style: 1},
}
var DEF_STYLE_ALIGN = &excelize.Alignment{
	Horizontal: "center",
	Vertical:   "center",
}

var DEF_STYLE_HYPERLINK = &excelize.Style{
	Font: &excelize.Font{Color: "1265BE", Underline: "single"},
}

var DEF_STYLE_MERGE_ROW = &excelize.Style{
	Alignment: &excelize.Alignment{
		Vertical: "center",
	},
}

func init() {
	DEF_STYLE = make(map[DefStyleKeys]int)
}

// 获取默认样式
func defStyle(key DefStyleKeys, f *excelize.File) int {
	if style, ok := DEF_STYLE[key]; ok {
		return style
	} else {
		var style int
		switch key {
		case DefStyleKeys_Link:
			//默认链接样式
			style, _ = f.NewStyle(DEF_STYLE_HYPERLINK)
			DEF_STYLE[DefStyleKeys_Link] = style
		case DefStyleKeys_Merge_ROW:
			//默认合并单元格样式
			style, _ = f.NewStyle(DEF_STYLE_MERGE_ROW)
			DEF_STYLE[DefStyleKeys_Merge_ROW] = style
		case DefStyleKeys_WRAP_TEXT:
			//默认自动换行样式
			style, _ = f.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{
					WrapText: true,
				},
			})
			DEF_STYLE[DefStyleKeys_WRAP_TEXT] = style
		}
		return style
	}
}

func defStyleHeader(f *excelize.File, font *excelize.Font) int {
	style, err := f.NewStyle(&excelize.Style{
		Border:    DEF_STYLE_BORDER,
		Font:      font,
		Alignment: DEF_STYLE_ALIGN,
		Fill: excelize.Fill{
			Type:    "pattern",
			Shading: 0,
			Pattern: 13,
			Color:   []string{"BFBFBF"},
		},
	})
	if err != nil {
		fmt.Println(err)
		return 0
	}
	return style
}

func newStyle(f *excelize.File, style *Style) (int, error) {
	return f.NewStyle(toStyle(style))
}

func newStyleWithColumn(f *excelize.File, column *Column) int {
	if column.Style != nil {
		styleId, err := newStyle(f, column.Style)
		if err != nil {
			return 0
		}
		return styleId
	}
	return 0
}

func toStyle(style *Style) *excelize.Style {
	var _style = &excelize.Style{}
	if style != nil {
		if style.Font != nil {
			var font = &excelize.Font{}
			if style.Font.Bold != nil {
				font.Bold = *style.Font.Bold
			}
			if style.Font.Italic != nil {
				font.Italic = *style.Font.Italic
			}
			if style.Font.Underline != nil {
				font.Underline = *style.Font.Underline
			}
			if style.Font.Family != nil {
				font.Family = *style.Font.Family
			}
			if style.Font.Size != nil {
				font.Size = *style.Font.Size
			}
			if style.Font.Strike != nil {
				font.Strike = *style.Font.Strike
			}
			if style.Font.Color != nil {
				font.Color = *style.Font.Color
			}
			if style.Font.ColorIndexed != nil {
				font.ColorIndexed = *style.Font.ColorIndexed
			}
			if style.Font.ColorTheme != nil {
				font.ColorTheme = style.Font.ColorTheme
			}
			if style.Font.ColorTint != nil {
				font.ColorTint = *style.Font.ColorTint
			}
			if style.Font.VertAlign != nil {
				font.VertAlign = *style.Font.VertAlign
			}
			_style.Font = font
		}
		if style.Border != nil && len(style.Border) > 0 {
			var borders = []excelize.Border{}
			for i, border := range style.Border {
				var _border = &excelize.Border{}
				if border.Type != nil {
					_border.Type = *border.Type
				}
				if border.Color != nil {
					_border.Color = *border.Color
				}
				if border.Style != nil {
					_border.Style = *border.Style
				}
				borders[i] = *_border
			}
			_style.Border = borders
		}
		if style.Fill != nil {
			var _fill = &excelize.Fill{}
			if style.Fill.Type != nil {
				_fill.Type = *style.Fill.Type
			}
			if style.Fill.Pattern != nil {
				_fill.Pattern = *style.Fill.Pattern
			}
			if style.Fill.Color != nil {
				_fill.Color = *style.Fill.Color
			}
			if style.Fill.Shading != nil {
				_fill.Shading = *style.Fill.Shading
			}
			_style.Fill = *_fill
		}
		if style.Alignment != nil {
			var _alignment = &excelize.Alignment{}
			_alignment.WrapText = style.Alignment.WrapText
			if style.Alignment.Vertical != nil && *style.Alignment.Vertical != "" {
				_alignment.Vertical = *style.Alignment.Vertical
			}
			if style.Alignment.Horizontal != nil && *style.Alignment.Horizontal != "" {
				_alignment.Horizontal = *style.Alignment.Horizontal
			}
			if style.Alignment.Indent > 0 {
				_alignment.Indent = style.Alignment.Indent
			}
			if style.Alignment.JustifyLastLine {
				_alignment.JustifyLastLine = style.Alignment.JustifyLastLine
			}
			if style.Alignment.ShrinkToFit {
				_alignment.ShrinkToFit = style.Alignment.ShrinkToFit
			}
			if style.Alignment.TextRotation > 0 {
				_alignment.TextRotation = style.Alignment.TextRotation
			}
			_style.Alignment = _alignment
		}
	}
	return _style
}
