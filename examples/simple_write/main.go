package main

import (
	"fmt"
	"os"

	docxexp "github.com/little-yangyang/docx-exp"
)

func main() {

	f, err := os.Open("examples/simple_write/template.docx")
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
		"Name": "World",
	}

	if err := tpl.Render(data); err != nil {
		panic(err)
	}

	out, err := os.Create("examples/simple_write/result.docx")
	if err != nil {
		panic(err)
	}
	defer out.Close()

	if err := tpl.Save(out); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/simple_write/result.docx")
}
