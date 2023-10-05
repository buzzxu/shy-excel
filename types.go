package shy_excel

type Type string

const (
	TypeIndex     Type = "index"
	TypeString    Type = "string"
	TypeNumeric   Type = "numeric"
	TypeTime      Type = "datetime"
	TypeBool      Type = "bool"
	TypeHyperLink Type = "hyperLink"
)

type Header struct {
	Title          string    `json:"title,omitempty"`
	FreezeCol      int       `json:"freezeCol,omitempty"`
	Columns        []*Column `json:"columns"`
	Height         float64   `json:"height,omitempty"`
	Style          *Style    `json:"style,omitempty"`
	FontTitleSize  *float64  `json:"font_title_size,omitempty"`
	FontColumnSize *float64  `json:"font_column_size,omitempty"`
	FontHeaderSize *float64  `json:"font_header_size,omitempty"`
	FontFamily     *string   `json:"font_family"`
	AutoFilter     *bool     `json:"auto_filter,omitempty"`
}

type Column struct {
	Name    string    `json:"name"`
	Title   string    `json:"title"`
	Type    Type      `json:"type"`
	Merge   bool      `json:"merge,omitempty"`
	Font    *Font     `json:"font,omitempty"`
	Width   float64   `json:"width,omitempty"`
	Columns []*Column `json:"columns,omitempty"`
}

type Sheet struct {
	Name    string                    `json:"name,omitempty"`
	Header  *Header                   `json:"header"`
	Active  bool                      `json:"active,omitempty"`
	Visible bool                      `json:"visible,omitempty"`
	Layout  *PageLayout               `json:"layout,omitempty"`
	Margins *PageMargins              `json:"margins,omitempty"`
	Data    *[]map[string]interface{} `json:"data"`
}
type PageLayout struct {
	Size            *int    `json:"size,omitempty"`
	Orientation     *string `json:"orientation,omitempty"`
	FirstPageNumber *uint   `json:"firstPageNumber,omitempty"`
	AdjustTo        *uint   `json:"adjustTo,omitempty"`
	FitToHeight     *int    `json:"fitToHeight,omitempty"`
	FitToWidth      *int    `json:"fitToWidth,omitempty"`
	BlackAndWhite   *bool   `json:"blackAndWhite,omitempty"`
}
type PageMargins struct {
	Bottom       *float64 `json:"bottom,omitempty"`
	Footer       *float64 `json:"footer,omitempty"`
	Header       *float64 `json:"header,omitempty"`
	Left         *float64 `json:"left,omitempty"`
	Right        *float64 `json:"right,omitempty"`
	Top          *float64 `json:"top,omitempty"`
	Horizontally *bool    `json:"horizontally,omitempty"`
	Vertically   *bool    `json:"vertically,omitempty"`
}
type Font struct {
	Bold         *bool    `json:"bold,omitempty"`
	Italic       *bool    `json:"italic,omitempty"`
	Underline    *string  `json:"underline,omitempty"`
	Family       *string  `json:"family,omitempty"`
	Size         *float64 `json:"size,omitempty"`
	Strike       *bool    `json:"strike,omitempty"`
	Color        *string  `json:"color,omitempty"`
	ColorIndexed *int     `json:"colorIndexed,omitempty"`
	ColorTheme   *int     `json:"colorTheme,omitempty"`
	ColorTint    *float64 `json:"colorTint,omitempty"`
	VertAlign    *string  `json:"vertAlign,omitempty"`
}
type Style struct {
	Border []*Border `json:"border,omitempty"`
	Font   *Font     `json:"font,omitempty"`
	Fill   *Fill     `json:"fill,omitempty"`
}
type Border struct {
	Type  *string `json:"type"`
	Color *string `json:"color"`
	Style *int    `json:"style"`
}
type Fill struct {
	Type    *string   `json:"type"`
	Pattern *int      `json:"pattern,omitempty"`
	Color   *[]string `json:"color,omitempty"`
	Shading *int      `json:"shading,omitempty"`
}
type Table []*Sheet

func (header *Header) Count() int {
	count := 0
	for _, column := range header.Columns {
		if len(column.Columns) > 0 {
			count += countColumns(column.Columns)
		} else {
			count++
		}
	}
	return count
}

func (header *Header) Keys() []string {
	return nil
}

func (column *Column) IsColl() bool {
	return column.Columns != nil && len(column.Columns) > 0
}

// 递归结算数量
func countColumns(columns []*Column) int {
	count := 0
	for _, column := range columns {
		if len(column.Columns) > 0 {
			count += countColumns(column.Columns)
		} else {
			count++
		}
	}
	return count
}
