package template

import (
	"log"
	"os"
	ttmpl "text/template"

	"github.com/tealeg/xlsx"
)

type cellMapper func(*xlsx.Cell) error

type Template struct {
	sheet *xlsx.Sheet
}

func New(sheet *xlsx.Sheet) *Template {
	return &Template{
		sheet: sheet,
	}
}

func (t *Template) Execute(data interface{}) error {
	mapCells(sheet, func(cell *xlsx.Cell) error {
		return applyData(cell, data)
	})
	return nil
}

func mapCells(sheet *xlsx.Sheet, fn cellMapper) error {
	log.Printf("%d rows\n", len(sheet.Rows))
	log.Printf("%d columns\n", len(sheet.Cols))

	for _, row := range sheet.Rows {
		log.Printf("%d cells\n", len(row.Cells))
		for _, cell := range row.Cells {
			if cell.Value == "" {
				continue
			}

			err := fn(cell)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

func applyData(cell *xlsx.Cell, data interface{}) error {
	tmpl, err := ttmpl.New("test").Parse(cell.Value)
	if err != nil {
		return err
	}

	err = tmpl.Execute(os.Stdout, data)
	if err != nil {
		return err
	}

	return nil
}
