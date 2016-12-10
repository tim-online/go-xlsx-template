package template

import (
	"github.com/tealeg/xlsx"
)

func libreOfficeFillFix(file *xlsx.File) {
	for _, sheet := range file.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				style := cell.GetStyle()
				style.ApplyFill = true
			}
		}
	}
}
