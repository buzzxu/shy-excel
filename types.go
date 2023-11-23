package shyexcel

import "reflect"

type Type string
type ResponseType string

const (
	TypeIndex     Type = "index"
	TypeString    Type = "string"
	TypeNumeric   Type = "numeric"
	TypeTime      Type = "datetime"
	TypeBool      Type = "bool"
	TypeHyperLink Type = "hyperLink"
	TypeImage     Type = "image"

	JSON     ResponseType = "json"
	MsgPack  ResponseType = "msgpack"
	Protobuf ResponseType = "protobuf"
)

type Table struct {
	Sheets []*Sheet `json:"sheets" msgpack:"sheets,as_array"`
}
type Sheet struct {
	Name    string                    `json:"name,omitempty" msgpack:"name"`
	Header  *Header                   `json:"header" msgpack:"header"`
	Active  bool                      `json:"active,omitempty" msgpack:"active,omitempty"`
	Visible bool                      `json:"visible,omitempty" msgpack:"visible,omitempty"`
	Layout  *PageLayout               `json:"layout,omitempty" msgpack:"layout,omitempty"`
	Margins *PageMargins              `json:"margins,omitempty" msgpack:"margins,omitempty"`
	Data    *[]map[string]interface{} `json:"data" msgpack:"data,as_array"`
}

type Header struct {
	Title          string    `json:"title,omitempty" msgpack:"title,omitempty"`
	FreezeCol      int       `json:"freezeCol,omitempty" msgpack:"freezeCol,omitempty"`
	Columns        []*Column `json:"columns" msgpack:"columns,omitempty,as_array"`
	Height         float64   `json:"height,omitempty" msgpack:"height,omitempty"`
	Style          *Style    `json:"style,omitempty" msgpack:"style,omitempty"`
	FontTitleSize  *float64  `json:"font_title_size,omitempty" msgpack:"fontTitleSize,omitempty"`
	FontColumnSize *float64  `json:"font_column_size,omitempty" msgpack:"fontColumnSize,omitempty"`
	FontHeaderSize *float64  `json:"font_header_size,omitempty" msgpack:"fontHeaderSize,omitempty"`
	FontFamily     *string   `json:"font_family,omitempty" msgpack:"fontFamily,omitempty"`
	AutoFilter     *bool     `json:"auto_filter,omitempty" msgpack:"autoFilter,omitempty"`
	dept           *int      `msgpack:"-"`
}

type Column struct {
	Name    string    `json:"name" msgpack:"name"`
	Title   string    `json:"title" msgpack:"title"`
	Type    Type      `json:"type" msgpack:"type"`
	Merge   bool      `json:"merge,omitempty" msgpack:"merge"`
	Font    *Font     `json:"font,omitempty" msgpack:"font"`
	Width   float64   `json:"width,omitempty" msgpack:"width"`
	Columns []*Column `json:"columns,omitempty" msgpack:"columns"`
	dept    *int      `msgpack:"-"`
}

type Data interface {
	map[string]interface{}
}

type PageLayout struct {
	Size            *int    `json:"size,omitempty"  msgpack:"size,omitempty"`
	Orientation     *string `json:"orientation,omitempty" msgpack:"orientation,omitempty"`
	FirstPageNumber *uint   `json:"firstPageNumber,omitempty" msgpack:"firstPageNumber,omitempty"`
	AdjustTo        *uint   `json:"adjustTo,omitempty" msgpack:"adjustTo,omitempty"`
	FitToHeight     *int    `json:"fitToHeight,omitempty" msgpack:"fitToHeight,omitempty"`
	FitToWidth      *int    `json:"fitToWidth,omitempty" msgpack:"fitToWidth,omitempty"`
	BlackAndWhite   *bool   `json:"blackAndWhite,omitempty" msgpack:"blackAndWhite,omitempty"`
}
type PageMargins struct {
	Bottom       *float64 `json:"bottom,omitempty" msgpack:"bottom,omitempty"`
	Footer       *float64 `json:"footer,omitempty" msgpack:"footer,omitempty"`
	Header       *float64 `json:"header,omitempty" msgpack:"header,omitempty"`
	Left         *float64 `json:"left,omitempty" msgpack:"left,omitempty"`
	Right        *float64 `json:"right,omitempty" msgpack:"right,omitempty"`
	Top          *float64 `json:"top,omitempty" msgpack:"top,omitempty"`
	Horizontally *bool    `json:"horizontally,omitempty" msgpack:"horizontally,omitempty"`
	Vertically   *bool    `json:"vertically,omitempty" msgpack:"vertically,omitempty"`
}
type Font struct {
	Bold         *bool    `json:"bold,omitempty" msgpack:"bold,omitempty"`
	Italic       *bool    `json:"italic,omitempty" msgpack:"italic,omitempty"`
	Underline    *string  `json:"underline,omitempty" msgpack:"underline,omitempty"`
	Family       *string  `json:"family,omitempty" msgpack:"family,omitempty"`
	Size         *float64 `json:"size,omitempty" msgpack:"size,omitempty"`
	Strike       *bool    `json:"strike,omitempty" msgpack:"strike,omitempty"`
	Color        *string  `json:"color,omitempty" msgpack:"color,omitempty"`
	ColorIndexed *int     `json:"colorIndexed,omitempty" msgpack:"colorIndexed,omitempty"`
	ColorTheme   *int     `json:"colorTheme,omitempty" msgpack:"colorTheme,omitempty"`
	ColorTint    *float64 `json:"colorTint,omitempty" msgpack:"colorTint,omitempty"`
	VertAlign    *string  `json:"vertAlign,omitempty" msgpack:"vertAlign,omitempty"`
}
type Style struct {
	Border []*Border `json:"border,omitempty" msgpack:"border,as_array,omitempty"`
	Font   *Font     `json:"font,omitempty" msgpack:"font,omitempty"`
	Fill   *Fill     `json:"fill,omitempty" msgpack:"fill,omitempty"`
}
type Border struct {
	Type  *string `json:"type,omitempty" msgpack:"type,omitempty"`
	Color *string `json:"color,omitempty" msgpack:"color,omitempty"`
	Style *int    `json:"style,omitempty" msgpack:"style,omitempty"`
}
type Fill struct {
	Type    *string   `json:"type" msgpack:"type"`
	Pattern *int      `json:"pattern,omitempty" msgpack:"pattern,omitempty"`
	Color   *[]string `json:"color,omitempty" msgpack:"color,as_array,omitempty"`
	Shading *int      `json:"shading,omitempty" msgpack:"shading,omitempty"`
}

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

// Depth 获取列数的最大层级 用于合并header
func (header *Header) Depth() int {
	if header.dept != nil {
		return *header.dept
	}
	var dept = 0
	for _, column := range header.Columns {
		if column.Columns != nil {
			_dept := depth(column.Columns)
			if _dept > dept {
				dept = _dept
			}
		}
	}
	header.dept = &dept
	return dept
}

func (header *Header) Keys() []string {
	return nil
}

func (column *Column) Depth() int {
	if column.dept != nil {
		return *column.dept
	}
	var dept = 0
	for _, column := range column.Columns {
		if column.Columns != nil {
			_dept := depth(column.Columns)
			if _dept > dept {
				dept = _dept
			}
		}
	}
	column.dept = &dept
	return dept
}

// countColumns 递归结算数量
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

func depth(columns []*Column) int {
	maxDepth := 0
	for _, column := range columns {
		d := depth(column.Columns)
		if d+1 > maxDepth {
			maxDepth = d + 1
		}
	}
	return maxDepth
}

func isNil(i interface{}) bool {
	if i == nil {
		return true
	}
	switch reflect.TypeOf(i).Kind() {
	case reflect.Ptr, reflect.Map, reflect.Array, reflect.Chan, reflect.Slice:
		return reflect.ValueOf(i).IsNil()
	}

	return false
}
