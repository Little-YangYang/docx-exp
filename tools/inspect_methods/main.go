package main

import (
	"fmt"
	"reflect"

	"github.com/fumiama/go-docx"
)

func main() {
	p := &docx.Paragraph{}
	t := reflect.TypeOf(p)
	for i := 0; i < t.NumMethod(); i++ {
		fmt.Println("Paragraph." + t.Method(i).Name)
	}

	d := &docx.Docx{}
	td := reflect.TypeOf(d)
	for i := 0; i < td.NumMethod(); i++ {
		fmt.Println("Docx." + td.Method(i).Name)
	}
}
