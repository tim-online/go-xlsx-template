package template

import (
	"bufio"
	"bytes"
	"fmt"
	"log"
	"strconv"
	ttmpl "text/template"

	"github.com/tealeg/xlsx"
)

type cellMapper func(*xlsx.Cell) error

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
	err := t.mapCells(sheet, func(cell *xlsx.Cell) error {
		return expandRange(cell, data)
	})
	if err != nil {
		return err
	}

	err = t.mapCells(sheet, func(cell *xlsx.Cell) error {
		return t.applyData(cell, data)
	})
	if err != nil {
		return err
	}

	return nil
}

func (t *Template) mapCells(sheet *xlsx.Sheet, fn cellMapper) error {
	log.Printf("%d rows\n", len(sheet.Rows))
	log.Printf("%d columns\n", len(sheet.Cols))

	for rowI, row := range sheet.Rows {
		log.Printf("%d cells\n", len(row.Cells))
		for colI, cell := range row.Cells {
			if cell.Value == "" {
				continue
			}

			t.curRow = rowI
			t.curCol = colI

			err := fn(cell)
			if err != nil {
				address := fmt.Sprintf("%d/%d", rowI, colI)
				return fmt.Errorf("error in template (cell %s): %s", address, err)
			}
		}
	}

	return nil
}

func expandRange(cell *xlsx.Cell, data interface{}) error {
	return nil
}

func (t *Template) applyData(cell *xlsx.Cell, data interface{}) error {
	tmpl, err := ttmpl.New("tmpl_cell").
		Funcs(t.funcMap()).
		Option("missingkey=error").
		Parse(cell.Value)
	if err != nil {
		return err
	}

	switch cell.Type() {
	case xlsx.CellTypeString:
		fmt.Println("celltype string")
	case xlsx.CellTypeFormula:
		fmt.Println("celltype formula")
	case xlsx.CellTypeNumeric:
		fmt.Println("celltype numeric")
	case xlsx.CellTypeBool:
		fmt.Println("celltype bool")
	case xlsx.CellTypeInline:
		fmt.Println("celltype inline")
	case xlsx.CellTypeError:
		fmt.Println("celltype error")
	case xlsx.CellTypeDate:
		fmt.Println("celltype date")
	case xlsx.CellTypeGeneral:
		fmt.Println("celltype general")
	}

	var b bytes.Buffer
	w := bufio.NewWriter(&b)
	err = tmpl.Execute(w, data)
	if err != nil {
		return err
		// missing data: silently fail
		fmt.Printf("after (1): %s\n", cell.Value)
		return nil
	}

	fmt.Printf("before: %s\n", cell.Value)
	w.Flush()
	newValue := b.String()

	// if values are the same: do nothing
	if newValue == cell.Value {
		fmt.Printf("after (2): %s\n", cell.Value)
		return nil
	}

	// check if new value is a float
	f, err := strconv.ParseFloat(newValue, 64)
	if err == nil {
		cell.SetFloat(f)
		fmt.Printf("after (3): %s\n", cell.Value)
		return nil
	}

	// treat new value as a string
	cell.Value = newValue
	fmt.Printf("after (4): %s\n", cell.Value)

	return nil
}

func (t *Template) colRange(data interface{}) string {
	// can't loop over %type, type must be slice or map
	fmt.Println("--------------------------------------")
	fmt.Println(data)
	fmt.Println(t.curRow)
	fmt.Println(t.curCol)
	fmt.Println("--------------------------------------")
	return ""
}

func (t *Template) funcMap() ttmpl.FuncMap {
	return ttmpl.FuncMap{
		"col_range": t.colRange,
	}
}
