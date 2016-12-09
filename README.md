# Go XLSX template

``` go
package main

import (
	"github.com/tealeg/xlsx"
	"github.com/tim-online/go-xlsx-template"
)

f, err := xlsx.OpenFile("template.xlsx")
if err != nil {
	panic(err)
}

sheet := f.Sheets[0]
tmpl := template.New(sheet)

data := map[string]interface{}{
	"placeholder1": "Value 1",
	"placeholder2": "Value 2",
}

err = tmpl.Execute(data)
if err != nil {
	panic(err)
}
```

## Test dependencies

- https://github.com/tmc/watcher
