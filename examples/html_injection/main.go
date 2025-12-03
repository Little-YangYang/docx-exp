package main

import (
	"fmt"
	"os"

	docxexp "github.com/little-yangyang/docx-exp"
)

func main() {

	f, err := os.Open("examples/html_injection/template.docx")
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
		"HTMLContent": docxexp.HTMLInjector{
			Content: `
				<h1>Title Level 1</h1>
				<p>This is a paragraph with an image.</p>
				<p><img src="testdata/test_image.png" /></p>
				<h2>Subtitle Level 2</h2>
				<p>Another paragraph.</p>
			`,
		},
	}

	if err := tpl.Render(data); err != nil {
		panic(err)
	}

	out, err := os.Create("examples/html_injection/result.docx")
	if err != nil {
		panic(err)
	}
	defer out.Close()

	if err := tpl.Save(out); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/html_injection/result.docx")
}
