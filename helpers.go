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

var (
	isTemplateRegex    = regexp.MustCompile(`{{\s?col_range.*?}}`)
	isColExpanderRegex = regexp.MustCompile(`{{\s?col_range.*?}}`)
)

func colCells(sheet *xlsx.Sheet, col int) []*xlsx.Cell {
	cells := []*xlsx.Cell{}

	for _, row := range sheet.Rows {
		cells = append(cells, row.Cells[col])
	}

	return cells
}

func splitCellValue(cell string) []string {
	return strings.Split(cell, "\n")
}

func isTemplate(cell *xlsx.Cell) bool {
	return isTemplateRegex.MatchString(cell.Value)
}

func isColExpander(value string) bool {
	return isColExpanderRegex.MatchString(value)
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

func copyColToRight(sheet *xlsx.Sheet, colNumber int) {
	from := colNumber
	to := from + 1
	copyCol(sheet, from, to)
}

func copyCol(sheet *xlsx.Sheet, from int, to int) {
	fromCells := colCells(sheet, from)

	// loop every row in column and copy value from original column
	err := mapCol(sheet, to, func(rowI int, colI int, cell *xlsx.Cell) error {
		from := fromCells[rowI]
		to := cell
		err := copyCell(from, to)
		return err
	})

	if err != nil {
		log.Fatal(err)
	}

	fromCol := sheet.Cols[from]
	toCol := sheet.Cols[to]
	toCol.SetStyle(fromCol.GetStyle())
	toCol.Width = fromCol.Width
}

func insertNewColAt(sheet *xlsx.Sheet, colNumber int) {
	// create new col
	sheet.Cols = append(sheet.Cols, &xlsx.Col{})
	copy(sheet.Cols[colNumber+1:], sheet.Cols[colNumber:])
	sheet.Cols[colNumber] = &xlsx.Col{}

	// create new cell in each row
	for _, row := range sheet.Rows {
		row.Cells = append(row.Cells, &xlsx.Cell{Row: row})
		copy(row.Cells[colNumber+1:], row.Cells[colNumber:])
		row.Cells[colNumber] = &xlsx.Cell{Row: row}
	}
}

func mapCol(sheet *xlsx.Sheet, colI int, fn cellMapper) error {
	for rowI, row := range sheet.Rows {
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

func applyData(cell *xlsx.Cell, data interface{}) error {
	if cell.Value == "" {
		return nil
	}

	tmpl, err := ttmpl.New("tmpl_cell").
		Funcs(ttmpl.FuncMap{
			"col_range": emptyColRange,
		}).
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

func mapCells(sheet *xlsx.Sheet, fn cellMapper) error {
	for rowI, row := range sheet.Rows {
		for colI, cell := range row.Cells {
			err := fn(rowI, colI, cell)
			if err != nil {
				address := fmt.Sprintf("%d/%d", rowI, colI)
				return fmt.Errorf("error in template (cell %s): %s", address, err)
			}
		}
	}

	return nil
}

func emptyColRange(data interface{}) string {
	return ""
}
