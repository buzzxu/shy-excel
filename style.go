package shy_excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

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
	}
	return _style
}
