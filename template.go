package template

import (
	"bufio"
	"bytes"
	"fmt"
	"log"
	"regexp"
	"strconv"
	"strings"
	ttmpl "text/template"

	"github.com/tealeg/xlsx"
)

type cellMapper func(rowI int, colI int, cell *xlsx.Cell) error

type Template struct {
	sheet  *xlsx.Sheet
	curRow int
	curCol int
}

func New(sheet *xlsx.Sheet) *Template {
	return &Template{
		sheet: sheet,
	}
}

func (t *Template) Execute(data interface{}) error {
	err := t.mapCells(func(rowI int, colI int, cell *xlsx.Cell) error {
		return t.expandRanges(cell, data)
	})
	if err != nil {
		return nil
		return err
	}

	err = t.mapCells(func(rowI int, colI int, cell *xlsx.Cell) error {
		return t.applyData(cell, data)
	})
	if err != nil {
		return err
	}

	return nil
}

func (t *Template) mapCol(colI int, fn cellMapper) error {
	for rowI, row := range t.sheet.Rows {
		// cell doesn't exist in row
		if colI > (len(row.Cells) - 1) {
			return nil
		}

		cell := row.Cells[colI]
		err := fn(rowI, colI, cell)
		if err != nil {
			return err
		}
	}

	return nil
}

func (t *Template) mapCells(fn cellMapper) error {
	for rowI, row := range t.sheet.Rows {
		for colI, cell := range row.Cells {
			t.curRow = rowI
			t.curCol = colI

			err := fn(rowI, colI, cell)
			if err != nil {
				address := fmt.Sprintf("%d/%d", rowI, colI)
				return fmt.Errorf("error in template (cell %s): %s", address, err)
			}
		}
	}

	return nil
}

func (t *Template) expandRanges(cell *xlsx.Cell, data interface{}) error {
	// cell is empty: do nothing
	if cell.Value == "" {
		return nil
	}

	if isColExpander(cell) == false {
		return nil
	}

	values := splitCellValue(cell.Value)
	for i, _ := range values {
		value := values[i]
		tmpl, err := ttmpl.New("tmpl_range").
			Funcs(ttmpl.FuncMap{
				"col_range": t.colRange,
			}).
			Option("missingkey=error").
			Parse(value)
		if err != nil {
			return err
		}

		// at this point col_range is removed and the columns are copied

		var b bytes.Buffer
		w := bufio.NewWriter(&b)
		err = tmpl.Execute(w, data)
		if err != nil {
			return err
		}

		w.Flush()
		newValue := b.String()

		// if values are the same: do nothing
		if newValue == value {
			continue
		}

		values[i] = newValue
	}

	cell.Value = strings.Join(values, "")
	return nil
}

func (t *Template) applyData(cell *xlsx.Cell, data interface{}) error {
	if cell.Value == "" {
		return nil
	}

	tmpl, err := ttmpl.New("tmpl_cell").
		Option("missingkey=error").
		Parse(cell.Value)
	if err != nil {
		return err
	}

	switch cell.Type() {
	case xlsx.CellTypeString:
	case xlsx.CellTypeFormula:
	case xlsx.CellTypeNumeric:
	case xlsx.CellTypeBool:
	case xlsx.CellTypeInline:
	case xlsx.CellTypeError:
	case xlsx.CellTypeDate:
	case xlsx.CellTypeGeneral:
	default:
	}

	var b bytes.Buffer
	w := bufio.NewWriter(&b)
	err = tmpl.Execute(w, data)
	if err != nil {
		return err
		// missing data: silently fail
		return nil
	}

	w.Flush()
	newValue := b.String()

	// if values are the same: do nothing
	if newValue == cell.Value {
		return nil
	}

	// check if new value is a float
	f, err := strconv.ParseFloat(newValue, 64)
	if err == nil {
		cell.SetFloat(f)
		return nil
	}

	// treat new value as a string
	cell.Value = newValue

	return nil
}

func (t *Template) colRange(data []interface{}) string {
	// add new columns
	newColsCount := len(data) - 1
	for i := 0; i < newColsCount; i++ {
		t.copyColToRight(t.curCol)
	}

	// loop each element in data slice and fill corresponding coumn
	for i := 0; i < len(data); i++ {
		// get cells from original column
		cells := t.colCells(t.curCol)

		// get colls from added column (to the right)
		if i > 0 {
			cells = t.colCells(t.curCol + 1)
		}

		// loop each cell in column and apply data
		for _, cell := range cells {
			// fmt.Printf("\n")
			// fmt.Printf("\n")
			// fmt.Printf("%+v\n", cell.Value)
			// fmt.Printf("%+v\n", data[i])
			err := t.applyData(cell, data[i])
			if err != nil {
				log.Println(err)
			}
		}
	}

	return ""
}

func (t *Template) emptyColRange(data []interface{}) string {
	return ""
}

func (t *Template) colCells(col int) []*xlsx.Cell {
	cells := []*xlsx.Cell{}

	for _, row := range t.sheet.Rows {
		cells = append(cells, row.Cells[col])
	}

	return cells
}

func (t *Template) copyColToRight(colNumber int) {
	cells := t.colCells(colNumber)
	t.insertNewColAfter(colNumber)
	newColNumber := colNumber + 1

	// loop every row in column and copy value from original column
	err := t.mapCol(newColNumber, func(rowI int, colI int, cell *xlsx.Cell) error {
		from := cells[rowI]
		to := cell
		err := copyCell(from, to)
		return err
	})

	if err != nil {
		log.Fatal(err)
	}
}

func copyCell(from *xlsx.Cell, to *xlsx.Cell) error {
	switch from.Type() {
	case xlsx.CellTypeString:
		to.SetString(from.Value)
	case xlsx.CellTypeFormula:
		to.SetFormula(from.Value)
	case xlsx.CellTypeNumeric:
		i, err := from.Int()
		if err != nil {
			return err
		}
		to.SetInt(i)
	case xlsx.CellTypeBool:
		to.SetBool(from.Bool())
	case xlsx.CellTypeInline:
		to.Value = from.Value
	case xlsx.CellTypeError:
		to.Value = from.Value
	case xlsx.CellTypeDate:
		to.Value = from.Value
		f, err := to.Float()
		if err != nil {
			return err
		}
		to.SetDateTimeWithFormat(f, from.NumFmt)
	case xlsx.CellTypeGeneral:
		to.SetValue(from.Value)
	default:
		to.SetValue(from.Value)
	}

	style := from.GetStyle()
	to.SetStyle(style)

	return nil
}

func (t *Template) insertNewColAfter(colNumber int) {
	for _, row := range t.sheet.Rows {

		row.Cells = append(row.Cells, &xlsx.Cell{})
		copy(row.Cells[colNumber+1:], row.Cells[colNumber:])
		row.Cells[colNumber+1] = &xlsx.Cell{}
	}

	t.sheet.Cols = append(t.sheet.Cols, &xlsx.Col{})
}

func splitCellValue(cell string) []string {
	return strings.Split(cell, "\n")
}

func isColExpander(cell *xlsx.Cell) bool {
	return regexp.MustCompile(`{{\s?col_range.*?}}`).MatchString(cell.Value)
}
