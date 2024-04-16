package main

import (
	"bufio"
	"bytes"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/ast"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/renderer"
	"github.com/yuin/goldmark/util"
)

func check(e error) {
	if e != nil {
		panic(e)
	}
}

func assert(check bool) {
	if !check {
		panic("Assertion Failed")
	}
}

func init() {
	// Make sure to load your metered License API key prior to using the library.
	// If you need a key, you can sign up and create a free one at https://cloud.unidoc.io
	err := license.SetMeteredKey(os.Getenv(`UNIDOC_LICENSE_API_KEY`))
	if err != nil {
		panic(err)
	}
}

func unwrap[I interface{}](o I, e error) I {
	if e != nil {
		panic(e)
	}
	return o
}

func main() {
	var inflag = flag.String("input", "/proc/self/fd/0", "Markdown file used as input")
	var outflag = flag.String("output", "./converted.docx", "Docx File to write")
	flag.Parse()
	var buf bytes.Buffer
	abs_in := unwrap(filepath.Abs(*inflag))
	md_wd := filepath.Dir(abs_in)
	abs_out := unwrap(filepath.Abs(*outflag))
	if !strings.HasPrefix(md_wd, "/proc") {
		os.Chdir(md_wd)
	}
	md_file, err := os.ReadFile(abs_in)
	check(err)
	r := myRenderer{doc: document.New(), ndBullet: make([]document.NumberingDefinition, 0)}
	defer r.doc.Close()
	md := goldmark.New(
		goldmark.WithExtensions(extension.Table),
		goldmark.WithRenderer(&r))
	fmt.Printf("%+v\n", md.Renderer())
	// fmt.Println(string(md_file))
	if err := md.Convert(md_file, &buf); err != nil {
		panic(err)
	}
	r.doc.SaveToFile(abs_out)
	// fmt.Println(buf)
}

type myRenderer struct {
	nodeRendererFuncsTmp  map[ast.NodeKind]renderer.NodeRendererFunc
	doc                   *document.Document
	currentPara           *document.Paragraph
	currentRun            *document.Run
	currentTable          *document.Table
	currentTableRow       *document.Row
	currentTableCell      *document.Cell
	currentNumberingLevel int
	ndBullet              []document.NumberingDefinition
	initSync              sync.Once
}

func (r *myRenderer) Render(w io.Writer, source []byte, n ast.Node) error {
	fmt.Println("in render")
	r.initSync.Do(
		func() {
			r.nodeRendererFuncsTmp = make(map[ast.NodeKind]renderer.NodeRendererFunc)
			r.RegisterFuncs(r)
		})
	writer, ok := w.(util.BufWriter)
	if !ok {
		writer = bufio.NewWriter(w)
	}
	err := ast.Walk(n, func(n ast.Node, entering bool) (ast.WalkStatus, error) {
		s := ast.WalkStatus(ast.WalkContinue)
		var err error
		f := r.nodeRendererFuncsTmp[n.Kind()]
		if f != nil {
			s, err = f(writer, source, n, entering)
		} else {
			panic(fmt.Sprintf("unsupported %s", n.Kind().String()))
		}
		return s, err
	})
	if err != nil {
		return err
	}

	return nil
}

func (r *myRenderer) AddOptions(opts ...renderer.Option) {
}

func (r *myRenderer) Register(kind ast.NodeKind, v renderer.NodeRendererFunc) {
	r.nodeRendererFuncsTmp[kind] = v
}
