package template

import (
	"fmt"
	"log"
	"testing"

	"github.com/tealeg/xlsx"
)

func setup() {
}

func teardown() {
}

func newSheet(cells [][]string) *xlsx.Sheet {
	sheet := &xlsx.Sheet{
		Name:     "test",
		File:     nil,
		Selected: true,
	}

	for _, rowCells := range cells {
		row := sheet.AddRow()
		for _, value := range rowCells {
			cell := row.AddCell()
			cell.SetValue(value)
		}
	}
	return sheet
}

func compareSheets(sheet1 *xlsx.Sheet, sheet2 *xlsx.Sheet) error {
	for rowI, row := range sheet1.Rows {
		for colI, _ := range row.Cells {
			cell1 := sheet1.Cell(rowI, colI)
			cell2 := sheet2.Cell(rowI, colI)
			if cell1.Value != cell2.Value {
				return fmt.Errorf("cells in row %d, column %d are not the same ('%s' vs '%s')", rowI, colI, cell1.Value, cell2.Value)
			}
		}
	}

	return nil
}

type simpleTest struct {
	cell string
	data interface{}
	want string
	err  bool
}

func TestSimple(t *testing.T) {
	setup()
	defer teardown()

	tables := []simpleTest{
		simpleTest{"", []string{}, "", false},
		simpleTest{"Should be the same", []string{}, "Should be the same", false},
		simpleTest{"{{ .test }}", map[string]string{"test": "test"}, "test", false},
		simpleTest{"Should do nothing", map[string]string{"test": "test"}, "Should do nothing", false},
		simpleTest{"{{ .test }}", map[string]string{}, "{{ .test }}", true},
		simpleTest{"{{ }}", map[string]string{}, "{{ }}", true},
		simpleTest{"0", map[string]string{}, "0", false},
		simpleTest{"0.0", map[string]string{}, "0.0", false},
	}

	for i, table := range tables {
		cell := &xlsx.Cell{}
		cell.SetValue(table.cell)
		err := applyData(cell, table.data)
		if table.err == false && err != nil {
			t.Error(err)
			continue
		}
		if table.err == true {
			if err == nil {
				t.Errorf("Wanted an error for test %d, got nothing", i)
			}
		}

		want := table.want
		got := cell.Value

		if want != got {
			t.Errorf("got '%s', want '%s' for test %d", got, want, i)
		}
	}
}

func TestSimpleColRangeExpand(t *testing.T) {
	setup()
	defer teardown()

	cells := [][]string{
		[]string{"header 1", "header 2", "{{col_range .cols}}\nheader 3", "header 4"},
		[]string{"", "static 1", "row1", "static 3"},
		[]string{"", "static 2", "row2", "static 4"},
	}

	got := newSheet(cells)

	data := map[string][]interface{}{
		"cols": []interface{}{
			map[string]interface{}{},
			map[string]interface{}{},
		},
	}

	tmpl := New(got)
	err := tmpl.Execute(data)
	if err != nil {
		t.Error(err)
	}

	if len(got.Cols) != 5 {
		t.Errorf("Wanted 5 cols, got %d cols", len(got.Cols))
	}

	want := newSheet([][]string{
		[]string{"header 1", "header 2", "header 3", "header 3", "header 4"},
		[]string{"", "static 1", "row1", "row1", "static 3"},
		[]string{"", "static 2", "row2", "row2", "static 4"},
	})

	err = compareSheets(got, want)
	if err != nil {
		t.Error(err)
	}
}

func TestColRangeExpandWithVariables(t *testing.T) {
	setup()
	defer teardown()

	cells := [][]string{
		[]string{"header 1", "header 2", "{{col_range .cols}}\nheader 3", "header 4"},
		[]string{"", "static 1", "{{ .row1 }}", "static 3"},
		[]string{"", "static 2", "{{ .row2 }}", "static 4"},
	}

	got := newSheet(cells)

	data := map[string][]interface{}{
		"cols": []interface{}{
			map[string]interface{}{
				"row1": "dynamic 1",
				"row2": "dynamic 2",
			},
			map[string]interface{}{
				"row1": "dynamic 3",
				"row2": "dynamic 4",
			},
		},
	}

	tmpl := New(got)
	err := tmpl.Execute(data)
	if err != nil {
		t.Error(err)
	}

	if len(got.Cols) != 5 {
		t.Errorf("Wanted 5 cols, got %d cols", len(got.Cols))
	}

	want := newSheet([][]string{
		[]string{"header 1", "header 2", "header 3", "header 3", "header 4"},
		[]string{"", "static 1", "dynamic 1", "dynamic 3", "static 3"},
		[]string{"", "static 2", "dynamic 2", "dynamic 4", "static 4"},
	})

	err = compareSheets(got, want)
	if err != nil {
		t.Error(err)
	}
}

func TestSchipper(t *testing.T) {
	var err error
	file, err := xlsx.OpenFile("testdata/template.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	if len(file.Sheets) < 1 {
		log.Fatalf("Template has no sheets")
	}

	libreOfficeFillFix(file)
	sheet := file.Sheets[0]

	tmpl := New(sheet)

	data := map[string][]interface{}{
		"offices": []interface{}{
			map[string]interface{}{
				"placeholder1":                                     "Value 1",
				"placeholder2":                                     "Value 2",
				"aangifte_jaar":                                    2016,
				"office_id":                                        "TO",
				"office_name":                                      "Tim_online",
				"aangifte_naam":                                    "Test",
				"aangifte_periode":                                 "12/04",
				"aangifte_code":                                    "test",
				"ob_nummer":                                        "ob-1204",
				"omzet_hoog":                                       120.0,
				"btw_hoog":                                         110.0,
				"omzet_laag":                                       90.0,
				"btw_laag":                                         30.0,
				"prive_gebruik":                                    0.0,
				"omzet_belast":                                     0.0,
				"omzet_verlegd":                                    0.0,
				"btw_verlegd":                                      0.0,
				"levering_naar_landen_buiten_de_eu":                0.0,
				"omzet_binnen_eu":                                  0.0,
				"omzet_leveringen_uit_landen_buiten_eu":            0.0,
				"btw_leveringen_uit_landen_buiten_eu":              0.0,
				"levering_van_goederen_uit_landen_binnen_eu_omzet": 0.0,
				"levering_van_goederen_uit_landen_binnen_eu_btw":   0.0,
				"verschuldigde_btw":                                0.0,
				"voorbelasting":                                    0.0,
				"subtotaal":                                        0.0,
				"totaal_te_betalen":                                0.0,
			},
			map[string]interface{}{
				"placeholder1":                                     "Value 1",
				"placeholder2":                                     "Value 2",
				"aangifte_jaar":                                    2015,
				"office_id":                                        "TO",
				"office_name":                                      "Tim_online",
				"aangifte_naam":                                    "Test",
				"aangifte_periode":                                 "12/04",
				"aangifte_code":                                    "test",
				"ob_nummer":                                        "ob-1204",
				"omzet_hoog":                                       120.0,
				"btw_hoog":                                         110.0,
				"omzet_laag":                                       90.0,
				"btw_laag":                                         30.0,
				"prive_gebruik":                                    0.0,
				"omzet_belast":                                     0.0,
				"omzet_verlegd":                                    0.0,
				"btw_verlegd":                                      0.0,
				"levering_naar_landen_buiten_de_eu":                0.0,
				"omzet_binnen_eu":                                  0.0,
				"omzet_leveringen_uit_landen_buiten_eu":            0.0,
				"btw_leveringen_uit_landen_buiten_eu":              0.0,
				"levering_van_goederen_uit_landen_binnen_eu_omzet": 0.0,
				"levering_van_goederen_uit_landen_binnen_eu_btw":   0.0,
				"verschuldigde_btw":                                0.0,
				"voorbelasting":                                    0.0,
				"subtotaal":                                        0.0,
				"totaal_te_betalen":                                0.0,
			},
		},
	}

	err = tmpl.Execute(data)
	if err != nil {
		panic(err)
	}

	file.Save("test.xlsx")
}

func saveSheet(sheet *xlsx.Sheet, path string) error {
	f := &xlsx.File{
		Sheets: []*xlsx.Sheet{sheet},
	}
	return f.Save(path)
}
