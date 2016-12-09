package template

import (
	"log"
	"testing"

	"github.com/tealeg/xlsx"
)

var (
	sheet *xlsx.Sheet
)

func setup() {
	sheet = getSheet()
}

func getSheet() *xlsx.Sheet {
	f, err := xlsx.OpenFile("template.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	if len(f.Sheets) < 1 {
		log.Fatalf("Template has no sheets")
	}

	return f.Sheets[0]
}

func teardown() {
}

func TestNewTemplate(t *testing.T) {
	setup()
	defer teardown()

	tmpl := New(sheet)

	if tmpl == nil {
		t.Errorf("Template is nil")
	}
}

func TestSimple(t *testing.T) {
	setup()
	defer teardown()

	tmpl := New(sheet)

	data := map[string]interface{}{
		"placeholder1": "Value 1",
		"placeholder2": "Value 2",
	}

	err := tmpl.Execute(data)
	if err != nil {
		t.Error(err)
	}
}
