package template

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

type colExpander struct {
	from int
	to   int
	data interface{}
}

func (e *colExpander) Copy(sheet *xlsx.Sheet) error {
	if e.to == e.from {
		return nil
	}

	fmt.Printf("insert new col at %d\n", e.to)
	insertNewColAt(sheet, e.to)
	fmt.Printf("copying col %d to %d\n", e.from, e.to)
	copyCol(sheet, e.from, e.to)
	return nil
}

func (e *colExpander) ApplyData(sheet *xlsx.Sheet) error {
	colI := e.to
	err := mapCol(sheet, colI, func(rowI int, colI int, cell *xlsx.Cell) error {
		err := applyData(cell, e.data)
		return err
	})
	return err
}
