package main

import (
	"fmt"
	"os"

	docxexp "github.com/little-yangyang/docx-exp"
)

func main() {

	f, err := os.Open("examples/row_if/template.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	tpl, err := docxexp.New(f, fi.Size())
	if err != nil {
		panic(err)
	}

	// Test Case 1: ShowRows = true
	data1 := map[string]interface{}{
		"ShowRows": true,
	}
	if err := tpl.Render(data1); err != nil {
		panic(err)
	}
	out1, err := os.Create("examples/row_if/result_true.docx")
	if err != nil {
		panic(err)
	}
	defer out1.Close()
	if err := tpl.Save(out1); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/row_if/result_true.docx")

	// Test Case 2: ShowRows = false
	// Re-open template for second run
	f2, err := os.Open("examples/row_if/template.docx")
	if err != nil {
		panic(err)
	}
	defer f2.Close()
	fi2, _ := f2.Stat()
	tpl2, err := docxexp.New(f2, fi2.Size())
	if err != nil {
		panic(err)
	}

	data2 := map[string]interface{}{
		"ShowRows": false,
	}
	if err := tpl2.Render(data2); err != nil {
		panic(err)
	}
	out2, err := os.Create("examples/row_if/result_false.docx")
	if err != nil {
		panic(err)
	}
	defer out2.Close()
	if err := tpl2.Save(out2); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/row_if/result_false.docx")
}
