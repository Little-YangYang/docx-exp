package docxexp

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
	"strings"
	"text/template"

	"github.com/fumiama/go-docx"
)

// Injector is the interface that must be implemented by types that want to inject content
type Injector interface {
	Inject(doc *docx.Docx, p *docx.Paragraph) ([]interface{}, error)
}

// DocxTemplate represents a docx template
type DocxTemplate struct {
	doc   *docx.Docx
	funcs template.FuncMap

	// injectors maps placeholder strings to Injector objects
	injectors map[string]Injector
}

// New creates a new DocxTemplate
func New(r io.ReaderAt, size int64) (*DocxTemplate, error) {
	// Pre-process to ensure Content_Types includes image formats
	r, size, err := ensureContentTypes(r, size)
	if err != nil {
		return nil, err
	}

	doc, err := docx.Parse(r, size)
	if err != nil {
		return nil, err
	}
	return &DocxTemplate{
		doc:       doc,
		funcs:     make(template.FuncMap),
		injectors: make(map[string]Injector),
	}, nil
}

type types struct {
	XMLName  xml.Name `xml:"Types"`
	Xmlns    string   `xml:"xmlns,attr"`
	Defaults []struct {
		Extension   string `xml:"Extension,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Default"`
	Overrides []struct {
		PartName    string `xml:"PartName,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Override"`
}

func ensureContentTypes(r io.ReaderAt, size int64) (io.ReaderAt, int64, error) {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, 0, err
	}

	var ctFound bool
	for _, f := range zr.File {
		if f.Name == "[Content_Types].xml" {
			ctFound = true
			rc, err := f.Open()
			if err != nil {
				return nil, 0, err
			}
			data, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				return nil, 0, err
			}

			var t types
			if err := xml.Unmarshal(data, &t); err != nil {
				return nil, 0, err
			}
			t.Xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"

			hasPng := false
			hasJpg := false
			hasJpeg := false
			for _, d := range t.Defaults {
				if d.Extension == "png" {
					hasPng = true
				}
				if d.Extension == "jpg" {
					hasJpg = true
				}
				if d.Extension == "jpeg" {
					hasJpeg = true
				}
			}

			if !hasPng {
				t.Defaults = append(t.Defaults, struct {
					Extension   string `xml:"Extension,attr"`
					ContentType string `xml:"ContentType,attr"`
				}{Extension: "png", ContentType: "image/png"})
			}
			if !hasJpg {
				t.Defaults = append(t.Defaults, struct {
					Extension   string `xml:"Extension,attr"`
					ContentType string `xml:"ContentType,attr"`
				}{Extension: "jpg", ContentType: "image/jpeg"})
			}
			if !hasJpeg {
				t.Defaults = append(t.Defaults, struct {
					Extension   string `xml:"Extension,attr"`
					ContentType string `xml:"ContentType,attr"`
				}{Extension: "jpeg", ContentType: "image/jpeg"})
			}

			newData, err := xml.Marshal(t)
			if err != nil {
				return nil, 0, err
			}
			// Add xml header
			newData = append([]byte(xml.Header), newData...)

			fw, err := w.Create(f.Name)
			if err != nil {
				return nil, 0, err
			}
			if _, err := fw.Write(newData); err != nil {
				return nil, 0, err
			}
		} else {
			fw, err := w.Create(f.Name)
			if err != nil {
				return nil, 0, err
			}
			rc, err := f.Open()
			if err != nil {
				return nil, 0, err
			}
			if _, err := io.Copy(fw, rc); err != nil {
				rc.Close()
				return nil, 0, err
			}
			rc.Close()
		}
	}
	if !ctFound {
		return nil, 0, fmt.Errorf("[Content_Types].xml not found")
	}
	if err := w.Close(); err != nil {
		return nil, 0, err
	}

	b := buf.Bytes()
	return bytes.NewReader(b), int64(len(b)), nil
}

// Funcs registers custom functions
func (t *DocxTemplate) Funcs(f template.FuncMap) {
	for k, v := range f {
		t.funcs[k] = v
	}
}

// Save writes the document to w
func (t *DocxTemplate) Save(w io.Writer) error {
	_, err := t.doc.WriteTo(w)
	return err
}

func (t *DocxTemplate) Render(data interface{}) error {
	newItems, err := t.traverseItems(t.doc.Document.Body.Items, data)
	if err != nil {
		return err
	}
	t.doc.Document.Body.Items = newItems
	return nil
}

func (t *DocxTemplate) traverseItems(items []interface{}, data interface{}) ([]interface{}, error) {
	var newItems []interface{}
	i := 0
	for i < len(items) {
		item := items[i]

		// Check for block start in Paragraph
		if p, ok := item.(*docx.Paragraph); ok {
			text := t.getParagraphText(p)
			// Check for {{for ...}}
			if variable, sliceExpr, isFor := t.parseForTag(text); isFor {
				// Find end tag
				endIndex, err := t.findBlockEnd(items, i+1, "endfor")
				if err != nil {
					return nil, err
				}

				// Execute Loop
				loopItems, err := t.executeLoop(items[i+1:endIndex], variable, sliceExpr, data)
				if err != nil {
					return nil, err
				}
				newItems = append(newItems, loopItems...)

				i = endIndex + 1 // Skip end tag
				continue
			}

			// Check for {{if ...}}
			if condExpr, isIf := t.parseIfTag(text); isIf {
				endIndex, err := t.findBlockEnd(items, i+1, "endif")
				if err != nil {
					return nil, err
				}

				// Execute If
				ifResult, err := t.executeIf(items[i+1:endIndex], condExpr, data)
				if err != nil {
					return nil, err
				}
				newItems = append(newItems, ifResult...)

				i = endIndex + 1
				continue
			}
		}

		// Normal processing
		switch it := item.(type) {
		case *docx.Paragraph:
			replacedItems, err := t.processParagraph(it, data)
			if err != nil {
				return nil, err
			}
			if replacedItems != nil {
				newItems = append(newItems, replacedItems...)
			} else {
				newItems = append(newItems, it)
			}
		case *docx.Table:
			if err := t.processTable(it, data); err != nil {
				return nil, err
			}
			newItems = append(newItems, it)
		default:
			newItems = append(newItems, it)
		}
		i++
	}
	return newItems, nil
}

func (t *DocxTemplate) parseForTag(text string) (string, string, bool) {
	// {{for var in slice}}
	text = strings.TrimSpace(text)
	if strings.HasPrefix(text, "{{for ") && strings.HasSuffix(text, "}}") {
		content := strings.TrimSuffix(strings.TrimPrefix(text, "{{for "), "}}")
		parts := strings.Split(content, " in ")
		if len(parts) == 2 {
			return strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1]), true
		}
	}
	return "", "", false
}

func (t *DocxTemplate) parseIfTag(text string) (string, bool) {
	// {{if cond}}
	text = strings.TrimSpace(text)
	if strings.HasPrefix(text, "{{if ") && strings.HasSuffix(text, "}}") {
		return strings.TrimSuffix(strings.TrimPrefix(text, "{{if "), "}}"), true
	}
	return "", false
}

func (t *DocxTemplate) findBlockEnd(items []interface{}, start int, endTag string) (int, error) {
	depth := 0
	for i := start; i < len(items); i++ {
		if p, ok := items[i].(*docx.Paragraph); ok {
			text := strings.TrimSpace(t.getParagraphText(p))
			if strings.HasPrefix(text, "{{for ") || strings.HasPrefix(text, "{{if ") {
				depth++
			} else if text == "{{"+endTag+"}}" {
				if depth == 0 {
					return i, nil
				}
				depth--
			} else if (endTag == "endfor" && text == "{{endfor}}") || (endTag == "endif" && text == "{{endif}}") {
				// Handle exact match or simple match
				if depth == 0 {
					return i, nil
				}
				depth--
			}
			// Check for generic end if mixed? No, strictly match.
			// Simplified check:
			if strings.Contains(text, "{{"+endTag+"}}") {
				// This is loose. Ideally strictly parse.
				// For now assume the tag is the whole paragraph text or close to it.
				if depth == 0 {
					return i, nil
				}
				depth--
			}
		}
	}
	return -1, fmt.Errorf("block end {{%s}} not found", endTag)
}

func (t *DocxTemplate) executeLoop(block []interface{}, variable, sliceExpr string, data interface{}) ([]interface{}, error) {
	slice, err := t.evaluateExpression(sliceExpr, data)
	if err != nil {
		return nil, err
	}

	var result []interface{}
	sliceVal := reflect.ValueOf(slice)
	if sliceVal.Kind() == reflect.Slice || sliceVal.Kind() == reflect.Array {
		for i := 0; i < sliceVal.Len(); i++ {
			item := sliceVal.Index(i).Interface()

			// Create context: data + variable
			ctx := map[string]interface{}{
				variable:     item,
				"__parent__": data, // Fallback?
			}

			// Clone block
			clonedBlock, err := t.cloneBlock(block)
			if err != nil {
				return nil, err
			}

			// Render block with context
			processedBlock, err := t.traverseItems(clonedBlock, ctx)
			if err != nil {
				return nil, err
			}
			result = append(result, processedBlock...)
		}
	}
	return result, nil
}

func (t *DocxTemplate) executeIf(block []interface{}, condExpr string, data interface{}) ([]interface{}, error) {
	val, err := t.evaluateExpression(condExpr, data)
	if err != nil {
		return nil, err
	}

	isTrue := false
	if val != nil {
		v := reflect.ValueOf(val)
		// fmt.Printf("DEBUG: If %s = %v (%T)\n", condExpr, val, val)
		switch v.Kind() {
		case reflect.Bool:
			isTrue = v.Bool()
		case reflect.String:
			isTrue = v.String() != ""
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			isTrue = v.Int() != 0
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			isTrue = v.Uint() != 0
		case reflect.Float32, reflect.Float64:
			isTrue = v.Float() != 0
		case reflect.Ptr, reflect.Interface:
			isTrue = !v.IsNil()
		default:
			isTrue = true // Not nil/zero
		}
	}

	if isTrue {
		clonedBlock, err := t.cloneBlock(block)
		if err != nil {
			return nil, err
		}
		return t.traverseItems(clonedBlock, data)
	}
	return nil, nil
}

func (t *DocxTemplate) cloneBlock(items []interface{}) ([]interface{}, error) {
	newItems := make([]interface{}, len(items))
	for i, it := range items {
		switch item := it.(type) {
		case *docx.Paragraph:
			newP, err := t.cloneParagraph(item)
			if err != nil {
				return nil, err
			}
			newItems[i] = newP
		case *docx.Table:
			newT, err := t.cloneTable(item)
			if err != nil {
				return nil, err
			}
			newItems[i] = newT
		default:
			newItems[i] = item
		}
	}
	return newItems, nil
}

func (t *DocxTemplate) cloneTable(tbl *docx.Table) (*docx.Table, error) {
	newT := *tbl
	newT.TableRows = make([]*docx.WTableRow, len(tbl.TableRows))
	for i, row := range tbl.TableRows {
		newRow, err := t.cloneRow(row)
		if err != nil {
			return nil, err
		}
		newT.TableRows[i] = newRow
	}
	return &newT, nil
}

func (t *DocxTemplate) processTable(table *docx.Table, data interface{}) error {
	var newRows []*docx.WTableRow

	for _, row := range table.TableRows {
		rangeCmd, rangeContent, hasRange := t.checkRowRange(row)

		if hasRange {
			slice, err := t.evaluateExpression(rangeCmd, data)
			if err != nil {
				return err
			}

			sliceVal := reflect.ValueOf(slice)
			if sliceVal.Kind() == reflect.Slice || sliceVal.Kind() == reflect.Array {
				for i := 0; i < sliceVal.Len(); i++ {
					item := sliceVal.Index(i).Interface()

					clonedRow, err := t.cloneRow(row)
					if err != nil {
						return err
					}

					t.cleanRowRangeTag(clonedRow, rangeContent)

					if err := t.processRow(clonedRow, item); err != nil {
						return err
					}

					newRows = append(newRows, clonedRow)
				}
			}
		} else {
			if err := t.processRow(row, data); err != nil {
				return err
			}
			newRows = append(newRows, row)
		}
	}

	table.TableRows = newRows
	return nil
}

func (t *DocxTemplate) processRow(row *docx.WTableRow, data interface{}) error {
	for _, cell := range row.TableCells {
		var newParagraphs []*docx.Paragraph
		for _, p := range cell.Paragraphs {
			items, err := t.processParagraph(p, data)
			if err != nil {
				return err
			}
			if items != nil {
				for _, item := range items {
					if np, ok := item.(*docx.Paragraph); ok {
						newParagraphs = append(newParagraphs, np)
					}
				}
			} else {
				newParagraphs = append(newParagraphs, p)
			}
		}
		cell.Paragraphs = newParagraphs
		for _, tbl := range cell.Tables {
			if err := t.processTable(tbl, data); err != nil {
				return err
			}
		}
	}
	return nil
}

func (t *DocxTemplate) checkRowRange(row *docx.WTableRow) (string, string, bool) {
	if len(row.TableCells) == 0 {
		return "", "", false
	}
	cell := row.TableCells[0]
	if len(cell.Paragraphs) == 0 {
		return "", "", false
	}
	p := cell.Paragraphs[0]

	text := t.getParagraphText(p)
	if strings.HasPrefix(strings.TrimSpace(text), "{{ range") {
		start := strings.Index(text, "{{ range")
		end := strings.Index(text, "}}")
		if start != -1 && end != -1 {
			content := text[start : end+2]
			cmd := strings.TrimSpace(text[start+8 : end])
			return cmd, content, true
		}
	}
	return "", "", false
}

func (t *DocxTemplate) cleanRowRangeTag(row *docx.WTableRow, tag string) {
	if len(row.TableCells) > 0 && len(row.TableCells[0].Paragraphs) > 0 {
		p := row.TableCells[0].Paragraphs[0]
		text := t.getParagraphText(p)
		newText := strings.Replace(text, tag, "", 1)
		t.replaceTextInParagraph(p, text, newText)
	}
}

func (t *DocxTemplate) cloneRow(row *docx.WTableRow) (*docx.WTableRow, error) {
	newRow := *row
	newRow.TableCells = make([]*docx.WTableCell, len(row.TableCells))
	for i, cell := range row.TableCells {
		newCell, err := t.cloneCell(cell)
		if err != nil {
			return nil, err
		}
		newRow.TableCells[i] = newCell
	}
	if row.TableRowProperties != nil {
		p := *row.TableRowProperties
		newRow.TableRowProperties = &p
	}
	return &newRow, nil
}

func (t *DocxTemplate) cloneCell(cell *docx.WTableCell) (*docx.WTableCell, error) {
	newCell := *cell
	newCell.Paragraphs = make([]*docx.Paragraph, len(cell.Paragraphs))
	for i, p := range cell.Paragraphs {
		newP, err := t.cloneParagraph(p)
		if err != nil {
			return nil, err
		}
		newCell.Paragraphs[i] = newP
	}
	newCell.Tables = make([]*docx.Table, len(cell.Tables))
	for i, tbl := range cell.Tables {
		newCell.Tables[i] = tbl // TODO: Clone
	}
	return &newCell, nil
}

func (t *DocxTemplate) cloneParagraph(p *docx.Paragraph) (*docx.Paragraph, error) {
	newP := *p
	newP.Children = make([]interface{}, len(p.Children))
	for i, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			newRun := *run
			newRun.Children = make([]interface{}, len(run.Children))
			for j, rc := range run.Children {
				if txt, ok := rc.(*docx.Text); ok {
					newTxt := *txt
					newRun.Children[j] = &newTxt
				} else {
					newRun.Children[j] = rc
				}
			}
			newP.Children[i] = &newRun
		} else {
			newP.Children[i] = child
		}
	}
	return &newP, nil
}

func (t *DocxTemplate) evaluateExpression(expr string, data interface{}) (interface{}, error) {
	path := strings.TrimPrefix(expr, ".")
	parts := strings.Split(path, ".")
	val := reflect.ValueOf(data)

	for _, part := range parts {
		if part == "" {
			continue
		}
		if val.Kind() == reflect.Ptr {
			val = val.Elem()
		}
		if val.Kind() == reflect.Interface {
			val = val.Elem()
		}
		if val.Kind() == reflect.Struct {
			val = val.FieldByName(part)
		} else if val.Kind() == reflect.Map {
			val = val.MapIndex(reflect.ValueOf(part))
		}
		if !val.IsValid() {
			return nil, fmt.Errorf("field %s not found", part)
		}
	}

	return val.Interface(), nil
}

func (t *DocxTemplate) processParagraph(p *docx.Paragraph, data interface{}) ([]interface{}, error) {
	fullText := t.getParagraphText(p)

	if !strings.Contains(fullText, "{{") {
		return nil, nil
	}

	if t.funcs["inject"] == nil {
		t.funcs["inject"] = func(v Injector) string {
			id := fmt.Sprintf("__INJECT_%p__", v)
			t.injectors[id] = v
			return id
		}
	}

	tmpl := template.New("p")
	tmpl.Funcs(t.funcs)

	if ctx, ok := data.(map[string]interface{}); ok {
		funcMap := make(template.FuncMap)
		for k, v := range ctx {
			val := v
			funcMap[k] = func() interface{} { return val }
		}
		tmpl.Funcs(funcMap)
	}

	tmpl, err := tmpl.Parse(fullText)
	if err != nil {
		return nil, err
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, data); err != nil {
		return nil, err
	}

	renderedText := buf.String()

	if strings.Contains(renderedText, "__INJECT_") {
		for id, injector := range t.injectors {
			if strings.Contains(renderedText, id) {
				items, err := injector.Inject(t.doc, p)
				if err != nil {
					return nil, err
				}

				if items != nil {
					return items, nil
				}

				renderedText = strings.Replace(renderedText, id, "", -1)
			}
		}
	}

	t.replaceTextInParagraph(p, fullText, renderedText)
	return nil, nil
}

func (t *DocxTemplate) getParagraphText(p *docx.Paragraph) string {
	fullText := ""
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			for _, runChild := range run.Children {
				if text, ok := runChild.(*docx.Text); ok {
					fullText += text.Text
				}
			}
		}
	}
	return fullText
}

func (t *DocxTemplate) replaceTextInParagraph(p *docx.Paragraph, oldText, newText string) {
	first := true
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			for _, runChild := range run.Children {
				if text, ok := runChild.(*docx.Text); ok {
					if first {
						text.Text = newText
						first = false
					} else {
						text.Text = ""
					}
				}
			}
		}
	}
}
