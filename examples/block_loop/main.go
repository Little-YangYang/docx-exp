package main

import (
	"fmt"
	"os"

	docxexp "github.com/little-yangyang/docx-exp"
)

func main() {
	f, err := os.Open("examples/block_loop/template.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	tpl, err := docxexp.New(f, fi.Size())
	if err != nil {
		panic(err)
	}

	data := map[string]interface{}{
		"vulns": []map[string]interface{}{
			{
				"Name":      "SQL Injection",
				"HasRetest": true,
				"Desc":      "Bad SQL",
			},
			{
				"Name":      "XSS",
				"HasRetest": false,
				"Desc":      "Bad Script",
			},
		},
	}

	if err := tpl.Render(data); err != nil {
		panic(err)
	}

	out, err := os.Create("examples/block_loop/result.docx")
	if err != nil {
		panic(err)
	}
	defer out.Close()

	if err := tpl.Save(out); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/block_loop/result.docx")
}
