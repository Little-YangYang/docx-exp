package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	f, err := os.Open("examples/block_loop/template.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	fi, _ := f.Stat()

	// Check zip content
	zr, err := zip.NewReader(f, fi.Size())
	if err != nil {
		panic(err)
	}
	for _, zf := range zr.File {
		fmt.Println(zf.Name)
		if zf.Name == "word/document.xml" {
			rc, _ := zf.Open()
			defer rc.Close()
			io.Copy(os.Stdout, rc)
			fmt.Println()
		}
	}
}
