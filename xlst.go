package xlst

import (
	"errors"
	"fmt"

	"io"
	"reflect"
	"regexp"
	"strings"
	"sync"

	"github.com/aymerick/raymond"
	xlsx "github.com/tealeg/xlsx/v3"
)

var (
	rgTrim      = regexp.MustCompile(`^\{\{\s*|\s*\}\}$`)
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rgxMerge    = regexp.MustCompile(`\{\{\s*(\w+)\.\w+_merge\s*\}\}`)
	rangeRgx    = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
)

// Xlst Represents template struct
type Xlst struct {
	file      *xlsx.File
	report    *xlsx.File
	mergeMap  map[string]map[string]map[string]cellCounter
	mergeOnce sync.Once
	sync.Mutex
}

type cellCounter struct {
	cell  *xlsx.Cell
	count int
}

// Options for render has only one property WrapTextInAllCells for wrapping text
type Options struct {
	WrapTextInAllCells bool
}

// New creates new Xlst struct and returns pointer to it
func New() *Xlst {
	return &Xlst{}
}

// NewFromBinary creates new Xlst struct puts binary tempate into and returns pointer to it
func NewFromBinary(content []byte) (*Xlst, error) {
	file, err := xlsx.OpenBinary(content)
	if err != nil {
		return nil, err
	}

	res := &Xlst{file: file, mergeMap: make(map[string]map[string]map[string]cellCounter)}
	return res, nil
}

// Render renders report and stores it in a struct
func (m *Xlst) Render(in interface{}) error {
	m.Lock()
	defer m.Unlock()
	err := m.RenderWithOptions(in, nil)
	m.mergeCell()
	return err
}

// RenderWithOptions renders report with options provided and stores it in a struct
func (m *Xlst) RenderWithOptions(in interface{}, options *Options) error {
	if options == nil {
		options = new(Options)
	}
	report := xlsx.NewFile()
	for si, sheet := range m.file.Sheets {
		if _, ok := m.mergeMap[sheet.Name]; !ok {
			m.mergeMap[sheet.Name] = make(map[string]map[string]cellCounter)
		}
		ctx := getCtx(in, si)
		report.AddSheet(sheet.Name)
		cloneSheet(sheet, report.Sheets[si])

		err := m.renderRows(report.Sheets[si], getRows(sheet), ctx, options)
		if err != nil {
			return err
		}

	}
	m.report = report

	return nil
}

func getRows(sheet *xlsx.Sheet) []*xlsx.Row {
	rows := make([]*xlsx.Row, sheet.MaxRow)
	sheet.ForEachRow(func(r *xlsx.Row) error {
		rows[r.GetCoordinate()] = r
		//rows = append(rows, r)
		return nil
	})
	return rows
}

// ReadTemplate reads template from disk and stores it in a struct
func (m *Xlst) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	m.file = file
	m.mergeMap = make(map[string]map[string]map[string]cellCounter)
	return nil
}

// Save saves generated report to disk
func (m *Xlst) Save(path string) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}

	return m.report.Save(path)
}

func (m *Xlst) mergeCell() {
	m.mergeOnce.Do(func() {
		for _, sheep := range m.report.Sheet {
			for _, v := range m.mergeMap[sheep.Name] {
				for _, vv := range v {
					style := vv.cell.GetStyle()
					style.Border.Top = "thin"
					style.Border.Bottom = "thin"
					style.Border.Left = "thin"
					style.Border.Right = "thin"
					vv.cell.SetStyle(style)
					vv.cell.VMerge = vv.count
				}
			}
		}
	})
}

// Write writes generated report to provided writer
func (m *Xlst) Write(writer io.Writer) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}
	return m.report.Write(writer)
}

func (m *Xlst) renderRows(sheet *xlsx.Sheet, rows []*xlsx.Row, ctx map[string]interface{}, options *Options) error {
	for ri := 0; ri < len(rows); ri++ {
		row := rows[ri]

		rangeProp := getRangeProp(row)
		if rangeProp != "" {
			ri++

			rangeEndIndex := getRangeEndIndex(rows[ri:])
			if rangeEndIndex == -1 {
				return fmt.Errorf("End of range %q not found", rangeProp)
			}

			rangeEndIndex += ri

			rangeCtx := getRangeCtx(ctx, rangeProp)
			if rangeCtx == nil {
				return fmt.Errorf("Not expected context property for range %q", rangeProp)
			}

			for idx := range rangeCtx {
				localCtx := mergeCtx(rangeCtx[idx], ctx)
				err := m.renderRows(sheet, rows[ri:rangeEndIndex], localCtx, options)
				if err != nil {
					return err
				}
			}

			ri = rangeEndIndex

			continue
		}

		prop := getListProp(row)
		if prop == "" {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			err := renderRow(m, newRow, ctx)
			if err != nil {
				return err
			}
			continue
		}

		if !isArray(ctx, prop) {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			err := renderRow(m, newRow, ctx)
			if err != nil {
				return err
			}
			continue
		}

		arr := reflect.ValueOf(ctx[prop])
		arrBackup := ctx[prop]
		for i := 0; i < arr.Len(); i++ {
			newRow := sheet.AddRow()
			cloneRow(row, newRow, options)
			ctx[prop] = arr.Index(i).Interface()
			err := renderRow(m, newRow, ctx)
			if err != nil {
				return err
			}
		}
		ctx[prop] = arrBackup
	}
	return nil
}

func cloneCell(from, to *xlsx.Cell, options *Options) {
	*to = *from

	style := from.GetStyle()
	if options.WrapTextInAllCells {
		style.Alignment.WrapText = true
	}
	to.SetStyle(style)
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt
}

func cloneRow(from, to *xlsx.Row, options *Options) {
	if height := from.GetHeight(); height != 0 {
		to.SetHeight(height)
	}

	from.ForEachCell(func(fromCell *xlsx.Cell) error {
		toCell := to.AddCell()
		cloneCell(fromCell, toCell, options)
		return nil
	})
}

func renderCell(m *Xlst, cell *xlsx.Cell, ctx interface{}) error {
	sn := cell.Row.Sheet.Name
	bflag := false
	value := cell.String()
	if rgxMerge.MatchString(value) {
		bflag = true
	}
	tpl := strings.Replace(value, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
	template, err := raymond.Parse(tpl)
	if err != nil {
		return err
	}
	out, err := template.Exec(ctx)
	if bflag {
		key := rgxMerge.FindString(value)
		key = rgTrim.ReplaceAllString(key, "")
		isHeader := false
		if strings.HasPrefix(key, "_") || strings.HasPrefix(key, "_header_") {
			isHeader = true
		}
		if _, ok := m.mergeMap[sn][key]; !ok {
			m.mergeMap[sn][key] = make(map[string]cellCounter)
		}
		if _, ok := m.mergeMap[sn][key][out]; !ok {
			if isHeader {
				m.mergeMap[sn][key][out] = cellCounter{cell, 1}
			} else {
				m.mergeMap[sn][key][out] = cellCounter{cell, 0}
			}
		} else {
			counter := m.mergeMap[sn][key][out]
			counter.count++
			m.mergeMap[sn][key][out] = counter
		}

	}

	if err != nil {
		return err
	}
	cell.Value = out
	return nil
}

func cloneSheet(from, to *xlsx.Sheet) {
	from.Cols.ForEach(func(idx int, col *xlsx.Col) {
		to.Cols.Add(col)
	})

}

func getCtx(in interface{}, i int) map[string]interface{} {
	if ctx, ok := in.(map[string]interface{}); ok {
		return ctx
	}
	if ctxSlice, ok := in.([]interface{}); ok {
		if len(ctxSlice) > i {
			_ctx := ctxSlice[i]
			if ctx, ok := _ctx.(map[string]interface{}); ok {
				return ctx
			}
		}
		return nil
	}
	return nil
}

func getRangeCtx(ctx map[string]interface{}, prop string) []map[string]interface{} {
	val, ok := ctx[prop]
	if !ok {
		return nil
	}

	if propCtx, ok := val.([]map[string]interface{}); ok {
		return propCtx
	}

	return nil
}

func mergeCtx(local, global map[string]interface{}) map[string]interface{} {
	ctx := make(map[string]interface{})

	for k, v := range global {
		ctx[k] = v
	}

	for k, v := range local {
		ctx[k] = v
	}

	return ctx
}

func isArray(in map[string]interface{}, prop string) bool {
	val, ok := in[prop]
	if !ok {
		return false
	}
	switch reflect.TypeOf(val).Kind() {
	case reflect.Array, reflect.Slice:
		return true
	}
	return false
}

func getListProp(in *xlsx.Row) string {
	propValue := ""
	in.ForEachCell(func(c *xlsx.Cell) error {
		if propValue != "" {
			return nil
		}
		if c.Value == "" {
			return nil
		}
		if match := rgx.FindAllStringSubmatch(c.Value, -1); match != nil {
			propValue = match[0][1]
		}
		return nil
	})

	return propValue
}

func getRangeProp(in *xlsx.Row) string {
	if in.Sheet.MaxCol != 0 {
		value := in.GetCell(0).Value
		match := rangeRgx.FindAllStringSubmatch(value, -1)
		if match != nil {
			return match[0][1]
		}
	}

	return ""
}

func getRangeEndIndex(rows []*xlsx.Row) int {
	var nesting int
	for idx := 0; idx < len(rows); idx++ {
		if rows[idx].Sheet.MaxCol == 0 {
			continue
		}

		value := rows[idx].GetCell(0).Value
		if rangeEndRgx.MatchString(value) {
			if nesting == 0 {
				return idx
			}

			nesting--
			continue
		}

		if rangeRgx.MatchString(value) {
			nesting++
		}
	}

	return -1
}

func renderRow(m *Xlst, in *xlsx.Row, ctx interface{}) error {
	err := in.ForEachCell(func(cell *xlsx.Cell) error {
		err := renderCell(m, cell, ctx)
		if err != nil {
			return err
		}
		return nil
	})

	return err
}
