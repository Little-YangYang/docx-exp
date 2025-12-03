# docx-exp

A wrapper around `github.com/fumiama/go-docx` to provide advanced template rendering features for DOCX files, inspired by `docxtpl` and `docx-templates`.

ATTENTION: This project is in development and may not be stable. Use at your own risk.

## Features

- **Variable Replacement**: `{{ .Var }}` or `{{ Var }}`.
- **Loops**:
  - **Block Loops**: `{{for item in Items}} ... {{endfor}}` (supports paragraphs, tables, nested structures).
  - **Table Row Loops**: `{{ range .Items }}` inside a table row.
- **Conditionals**: `{{if Condition}} ... {{endif}}`.
- **Injection**:
  - **Images**: Inject images dynamically.
  - **HTML**: Inject HTML content (`h1`, `h2`, `h3`, `p`, `img`).
- **Robustness**: Automatically patches `[Content_Types].xml` to support image formats.

## Installation

```bash
go get github.com/little-yangyang/docx-exp
```

## Usage

### Basic Rendering

```go
package main

import (
    "os"
    docxexp "github.com/little-yangyang/docx-exp"
)

func main() {
    f, _ := os.Open("template.docx")
    defer f.Close()
    fi, _ := f.Stat()

    tpl, _ := docxexp.New(f, fi.Size())

    data := map[string]interface{}{
        "Name": "World",
    }
    tpl.Render(data)

    out, _ := os.Create("result.docx")
    defer out.Close()
    tpl.Save(out)
}
```

### Block Loops

Use `{{for var in Slice}}` to repeat a block of content.

```text
{{for item in Items}}
Item: {{item.Name}}
{{endfor}}
```

### Conditionals

Use `{{if Condition}}` to conditionally show a block.

```text
{{if IsVisible}}
This is visible.
{{endif}}
```

### Injection

Use `{{ inject .Injector }}` in your template.

#### Injecting Images

```go
data := struct {
    MyImage docxexp.Injector
}{
    MyImage: docxexp.ImageInjector{
        Path: "image.png",
        Width: 100,
        Height: 100,
    },
}
```

## Project Structure

- `examples/`: Example usage scripts.
  - `block_loop/`: Block loops (`{{for}}`, `{{if}}`).
  - `html_injection/`: HTML content injection.
  - `complex_report/`: Complex report with tables and images.
  - `simple_write/`: Basic variable replacement.
- `tools/`: Utility scripts.
- `testdata/`: Test assets.
- `client.go`, `html.go`, `image.go`: Core library code.

## Usage

See `examples/` for runnable examples.
