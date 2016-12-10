package template

import (
	"log"
	"testing"

	"github.com/tealeg/xlsx"
)

var (
	file  *xlsx.File
	sheet *xlsx.Sheet
)

func setup() {
	sheet = getSheet()
}

func getSheet() *xlsx.Sheet {
	var err error
	file, err = xlsx.OpenFile("testdata/template.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	if len(file.Sheets) < 1 {
		log.Fatalf("Template has no sheets")
	}

	libreOfficeFillFix(file)

	return file.Sheets[0]
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
		"placeholder1":  "Value 1",
		"placeholder2":  "Value 2",
		"aangifte_jaar": 10.0,
		"offices": map[string]string {
			"1": "een",
			"2": "twee",
		},
		"office_id": "TO",
		"office_name": "Tim_online",
		"aangifte_naam": "Test",
		"aangifte_periode": "12/04",
		"aangifte_code": "test",
		"ob_nummer": "ob-1204",
		"omzet_hoog": 120.0,
		"btw_hoog": 110.0,
		"omzet_laag": 90.0,
		"btw_laag": 30.0,
		"prive_gebruik": 0.0,
		"omzet_belast": 0.0,
		"omzet_verlegd": 0.0,
		"btw_verlegd": 0.0,
		"levering_naar_landen_buiten_de_eu": 0.0,
		"omzet_binnen_eu": 0.0,
		"omzet_leveringen_uit_landen_buiten_eu": 0.0,
		"btw_leveringen_uit_landen_buiten_eu": 0.0,
		"levering_van_goederen_uit_landen_binnen_eu_omzet": 0.0,
		"levering_van_goederen_uit_landen_binnen_eu_btw": 0.0,
		"verschuldigde_btw": 0.0,
		"voorbelasting": 0.0,
		"subtotaal": 0.0,
		"totaal_te_betalen": 0.0,
	}

	err := tmpl.Execute(data)
	if err != nil {
		t.Error(err)
	}

	file.Save("test.xlsx")
}
