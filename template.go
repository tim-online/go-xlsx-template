package template

import (
	"fmt"
	"io/ioutil"
	"strings"
	ttmpl "text/template"

	"github.com/tealeg/xlsx"
)

type cellMapper func(rowI int, colI int, cell *xlsx.Cell) error

type Template struct {
	sheet *xlsx.Sheet
}

func New(sheet *xlsx.Sheet) *Template {
	return &Template{
		sheet: sheet,
	}
}

func (t *Template) Execute(data interface{}) error {
	err := t.expandRanges(data)
	if err != nil {
		return err
	}

	err = mapCells(t.sheet, func(rowI int, colI int, cell *xlsx.Cell) error {
		return applyData(cell, data)
	})
	if err != nil {
		return err
	}

	return nil
}

func (t *Template) expandRanges(data interface{}) error {
	err := t.expandColRanges(data)
	if err != nil {
		return err
	}

	// err := t.expandRowRanges(data)
	// if err != nil {
	// 	return err
	// }

	return nil
}

func (t *Template) expandColRanges(data interface{}) error {
	// get col expanders for each col
	for colI, _ := range t.sheet.Cols {
		expanders, err := t.getColExpanders(colI, data)
		if err != nil {
			return err
		}

		for _, expander := range expanders {
			err := expander.Copy(t.sheet)
			if err != nil {
				return nil
			}
		}

		for _, expander := range expanders {
			err := expander.ApplyData(t.sheet)
			if err != nil {
				return nil
			}
		}
	}

	for _, col := range t.sheet.Cols {
		fmt.Printf("col width: %f\n", col.Width)
	}
	return nil
}

func (t *Template) getColExpanders(colI int, data interface{}) ([]colExpander, error) {
	expanders := []colExpander{}

	// loop every row in column and copy value from original column
	err := mapCol(t.sheet, colI, func(rowI int, colI int, cell *xlsx.Cell) error {
		// cell is empty: do nothing
		if cell.Value == "" {
			return nil
		}

		// threat each newline in a cell as a separate template
		values := splitCellValue(cell.Value)

		for i, _ := range values {
			value := values[i]

			// value doesn't contain a col_range function
			if isColExpander(value) == false {
				continue
			}

			exps, err := t.getColExpandersFromTemplate(colI, value, data)
			if err != nil {
				return err
			}

			if len(exps) > 0 {
				values[i] = ""
				expanders = append(expanders, exps...)
			}
		}

		cell.Value = strings.Join(values, "")
		return nil
	})
	return expanders, err
}

func (t *Template) getColExpandersFromTemplate(colI int, tmplValue string, data interface{}) ([]colExpander, error) {
	expanders := []colExpander{}

	callback := func(from int, to int, data interface{}) {
		expander := colExpander{
			from: from,
			to:   to,
			data: data,
		}
		expanders = append(expanders, expander)
	}

	colRange := t.createColRangeFunc(colI, data, callback)
	tmpl, err := ttmpl.New("tmpl_range").
		Funcs(ttmpl.FuncMap{
			"col_range": colRange,
		}).
		Option("missingkey=error").
		Parse(tmplValue)
	if err != nil {
		return expanders, err
	}

	err = tmpl.Execute(ioutil.Discard, data)

	// at this point col_range is removed
	return expanders, err
}

type callbackFn func(from int, to int, data interface{})

func (t *Template) createColRangeFunc(colI int, data interface{}, callback callbackFn) func(interface{}) string {
	fn := func(data interface{}) string {
		// catch errors
		dataSlice := data.([]interface{})

		// add new columns
		newColsCount := len(dataSlice)
		for i := 0; i < newColsCount; i++ {
			from := colI
			to := colI + i
			colData := dataSlice[i]
			callback(from, to, colData)
		}

		return ""
	}

	return fn
}
