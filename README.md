# Go XLSX template

This package combines the go
[`text/template`](https://golang.org/pkg/text/template/) with
[`tealeg/xlsx`](https://github.com/tealeg/xlsx).

The endresult is that you can use Go's template language to fill an XLSX
documents with data.


## Example

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
