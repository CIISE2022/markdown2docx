package main

import (
	"fmt"
	"log"

	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/common"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/measurement"
	"github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes"
	"github.com/unidoc/unioffice/schema/soo/wml"
	"github.com/yuin/goldmark/ast"

	east "github.com/yuin/goldmark/extension/ast"
	"github.com/yuin/goldmark/renderer"
	"github.com/yuin/goldmark/util"
)

func (r *myRenderer) RegisterFuncs(reg renderer.NodeRendererFuncRegisterer) {
	// blocks
	fmt.Println("In registerfunc")
	reg.Register(ast.KindDocument, r.renderDocument)
	reg.Register(ast.KindHeading, r.renderHeading)
	reg.Register(ast.KindBlockquote, r.renderBlockquote)
	reg.Register(ast.KindCodeBlock, r.renderCodeBlock)
	reg.Register(ast.KindFencedCodeBlock, r.renderFencedCodeBlock)
	reg.Register(ast.KindHTMLBlock, r.renderHTMLBlock)
	reg.Register(ast.KindList, r.renderList)
	reg.Register(ast.KindListItem, r.renderListItem)
	reg.Register(ast.KindParagraph, r.renderParagraph)
	reg.Register(ast.KindTextBlock, r.renderTextBlock)
	reg.Register(ast.KindThematicBreak, r.renderThematicBreak)

	// inlines

	reg.Register(ast.KindAutoLink, r.renderAutoLink)
	reg.Register(ast.KindCodeSpan, r.renderCodeSpan)
	reg.Register(ast.KindEmphasis, r.renderEmphasis)
	reg.Register(ast.KindImage, r.renderImage)
	reg.Register(ast.KindLink, r.renderLink)
	reg.Register(ast.KindRawHTML, r.renderRawHTML)
	reg.Register(ast.KindText, r.renderText)
	reg.Register(ast.KindString, r.renderString)

	//Tables
	reg.Register(east.KindTable, r.renderTable)
	reg.Register(east.KindTableHeader, r.renderTableHeader)
	reg.Register(east.KindTableRow, r.renderTableRow)
	reg.Register(east.KindTableCell, r.renderTableCell)
}

func (r *myRenderer) renderDocument(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		code := r.doc.Styles.AddStyle("Code", wml.ST_StyleTypeParagraph, false)
		code.SetName("Code")
		code.RunProperties().SetFontFamily("Consolas")
		code.RunProperties().SetSize(10)
		code.ParagraphProperties().SetSpacing(0, 0)
		x := code.ParagraphProperties().X() //.PBdr.Top.ValAttr = wml.ST_Border(1)
		x.PBdr = &wml.CT_PBdr{
			Top:    &wml.CT_Border{ValAttr: wml.ST_Border(3)},
			Bottom: &wml.CT_Border{ValAttr: wml.ST_Border(3)},
			Left:   &wml.CT_Border{ValAttr: wml.ST_Border(3)},
			Right:  &wml.CT_Border{ValAttr: wml.ST_Border(3)},
		}
		caption := r.doc.Styles.AddStyle("Caption", wml.ST_StyleTypeParagraph, false)
		caption.SetName("Caption")
		caption.SetBasedOn("Normal")
		normal, found := r.doc.Styles.SearchStyleByName("Normal")
		if !found {
			panic("default style not found")
		}
		normal.RunProperties().SetFontFamily("Calibri")
		caption.SetNextStyle("Normal")
		caption.ParagraphProperties().SetAlignment(wml.ST_JcCenter)
		caption.RunProperties().SetItalic(true)
		caption.RunProperties().SetSize(9)
		caption.RunProperties().SetColor(color.FromHex("#1F497D"))
		link := r.doc.Styles.AddStyle("Hyperlink", wml.ST_StyleTypeCharacter, false)
		link.SetBasedOn("Normal")
		link.RunProperties().SetUnderline(wml.ST_UnderlineSingle, color.FromHex("#1F497D"))
		link.RunProperties().SetColor(color.FromHex("#1F497D"))
		nd := r.doc.Numbering.AddDefinition()
		for i := 1; i < 8; i++ {
			lvl := nd.AddLevel()
			lvl.SetFormat(wml.ST_NumberFormatDecimal)
			lvl.SetAlignment(wml.ST_JcLeft)
			text := ""
			for j := 1; j <= i; j++ {
				text += fmt.Sprintf("%%%d.", j)
			}
			lvl.SetText(text)
			h, found := r.doc.Styles.SearchStyleById("Heading" + fmt.Sprintf("%d", i))
			if !found {
				panic("No header")
			}
			h.ParagraphProperties().X().NumPr = &wml.CT_NumPr{
				Ilvl:  &wml.CT_DecimalNumber{ValAttr: lvl.X().IlvlAttr},
				NumId: &wml.CT_DecimalNumber{ValAttr: nd.AbstractNumberID()}}
			h.RunProperties().SetColor(color.FromHex("#0070C0"))
			h.RunProperties().SetFontFamily("Calibri")
			if i == 1 {
				h.RunProperties().SetBold(true)
				t := true
				h.ParagraphProperties().X().PageBreakBefore = &wml.CT_OnOff{
					ValAttr: &sharedTypes.ST_OnOff{Bool: &t}}
			}
		}

		//Headers and Footers
		ftr := r.doc.AddFooter()
		para := ftr.AddParagraph()
		para.Properties().AddTabStop(6*measurement.Inch, wml.ST_TabJc(wml.ST_PTabAlignmentRight), wml.ST_TabTlcNone)
		run := para.AddRun()
		run.AddTab()
		run.AddText("Page ")
		run.AddField(document.FieldCurrentPage)
		run.AddText(" de ")
		run.AddField(document.FieldNumberOfPages)
		r.doc.BodySection().SetFooter(ftr, wml.ST_HdrFtrDefault)
		r.doc.Settings.SetUpdateFieldsOnOpen(true)
		htoc := r.doc.AddParagraph()
		// htoc.SetStyle("Heading1")
		htoc.AddRun().AddText("Table of Contents")
		htoc.Runs()[0].Properties().SetSize(16)
		htoc.Runs()[0].Properties().SetBold(true)
		htoc.Runs()[0].Properties().SetColor(color.FromHex("#0070C0"))
		r.doc.AddParagraph().AddRun().AddField(document.FieldTOC)
		r.doc.AddParagraph().Properties().AddSection(wml.ST_SectionMarkNextPage)
		// nothing to do
		fmt.Println("Render Document")
	}
	return ast.WalkContinue, nil
}

func addStyledPara(doc *document.Document, style string) *document.Paragraph {
	para := doc.AddParagraph()
	para.SetStyle(style)
	return &para
}

func (r *myRenderer) renderHeading(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {

	n := node.(*ast.Heading)
	if entering {
		fmt.Printf("Enter heading %d\n", n.Level)
		assert(r.currentPara == nil)
		r.currentPara = addStyledPara(r.doc, "Heading"+fmt.Sprintf("%d", n.Level))
	} else {
		fmt.Printf("Exit heading %d\n", n.Level)
		r.currentPara = nil
		r.currentRun = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderBlockquote(
	w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		fmt.Println("enter Blockquote")
		assert(r.currentPara == nil)
		r.currentPara = addStyledPara(r.doc, "Caption")
	} else {
		fmt.Println("Exit Blocquote")
		r.currentPara = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderCodeBlock(w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		fmt.Println("Enter Code")
		assert(r.currentPara == nil)
		r.currentPara = addStyledPara(r.doc, "Code")
		l := n.Lines().Len()
		for i := 0; i < l; i++ {
			line_nr := n.Lines().At(i)
			line := string(line_nr.Value(source))
			run := r.currentPara.AddRun()
			run.AddText(line)
			run.AddBreak()
		}
	} else {
		fmt.Println("Exit Code")
		r.currentPara = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderFencedCodeBlock(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	return r.renderCodeBlock(w, source, node, entering)
}

func (r *myRenderer) renderHTMLBlock(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	panic("Do not use Raw HTML")
}

func (r *myRenderer) renderList(w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	n := node.(*ast.List)

	if entering {
		fmt.Println("Enter List")
		if n.IsOrdered() {
			r.ndBullet = append(r.ndBullet, r.doc.Numbering.Definitions()[0]) //todo: add support for ordered
		} else {
			r.ndBullet = append(r.ndBullet, r.doc.Numbering.Definitions()[0])
		}
		r.currentNumberingLevel += 1
	} else {
		fmt.Println("Exit List")
		r.ndBullet = r.ndBullet[:len(r.ndBullet)-1]
		r.currentNumberingLevel -= 1
	}

	return ast.WalkContinue, nil
}

func (r *myRenderer) renderListItem(w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		// assert(r.currentPara == nil)
		fmt.Println("Enter List Item")
		p := r.doc.AddParagraph()
		p.SetNumberingLevel(r.currentNumberingLevel - 1)
		p.SetNumberingDefinition(r.ndBullet[len(r.ndBullet)-1])
		r.currentPara = &p
	} else {
		fmt.Println("Exit ListItem")
		r.currentPara = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderParagraph(w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		fmt.Println("Enter Paragraph")
		if r.currentPara == nil {
			r.currentPara = addStyledPara(r.doc, "Normal")
		}
	} else {
		fmt.Println("Exit Paragraph")
		r.currentPara = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderTextBlock(w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		fmt.Println("Enter TextBlock")
		run := r.currentPara.AddRun()
		r.currentRun = &run
	} else {
		fmt.Println("Exit TextBlock")
		if r.currentPara == nil {
			para := r.doc.AddParagraph()
			r.currentPara = &para
			r.currentRun = nil
		}
		if r.currentRun == nil {
			run := r.currentPara.AddRun()
			r.currentRun = &run
		}
		if n.NextSibling() != nil && n.FirstChild() != nil {
			// r.currentRun.AddBreak()
		}
		r.currentRun = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderThematicBreak(
	w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	panic("Do not use thematic breaks")
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderAutoLink(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	n := node.(*ast.AutoLink)
	if !entering {
		return ast.WalkContinue, nil
	}
	fmt.Println("Autolink")
	url := n.URL(source)
	label := n.Label(source)
	h1 := r.currentPara.AddHyperLink()
	h1.SetTarget(string(url))
	run := h1.AddRun()
	run.Properties().SetStyle("Hyperlink")
	run.AddText(string(label))
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderCodeSpan(w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		fmt.Println("EnterCodeSpan") //todo: make span style
		run := r.currentPara.AddRun()
		run.Properties().SetFontFamily("Consolas")
		run.Properties().SetSize(10)
		r.currentRun = &run
	} else {
		fmt.Println("Exit Codespan")
		r.currentRun = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderEmphasis(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {

	if entering {
		run := r.currentPara.AddRun()
		run.Properties().SetBold(true)
		r.currentRun = &run
	} else {
		r.currentRun = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderLink(w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	n := node.(*ast.Link)
	if entering {
		fmt.Println("enter Link")
		url := util.EscapeHTML(util.URLEscape(n.Destination, true))
		h1 := r.currentPara.AddHyperLink()
		h1.SetTarget(string(url))
		run := h1.AddRun()
		run.Properties().SetStyle("Hyperlink")
		r.currentRun = &run
	} else {
		fmt.Println("Exit link")
		r.currentRun = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderImage(w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	if !entering {
		return ast.WalkContinue, nil
	}
	n := node.(*ast.Image)
	fmt.Println("Image")
	img1, err := common.ImageFromFile(string(n.Destination))
	if err != nil {
		log.Fatalf("unable to create image: %s", err)
	}
	img1ref, err := r.doc.AddImage(img1)
	if err != nil {
		log.Fatalf("unable to add image to document: %s", err)
	}
	if r.currentPara == nil {
		para := r.doc.AddParagraph()
		r.currentPara = &para
	}
	anchored, err := r.currentPara.AddRun().AddDrawingAnchored(img1ref)
	if err != nil {
		log.Fatalf("unable to add anchored image: %s", err)
	}
	anchored.SetTextWrapTopAndBottom()
	anchored.SetHAlignment(wml.WdST_AlignHCenter)
	anchored.SetOrigin(wml.WdST_RelFromH(wml.WdST_AlignHCenter), wml.WdST_RelFromVLine)
	return ast.WalkSkipChildren, nil
}

func (r *myRenderer) renderRawHTML(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	panic("do not use raw html")
	return ast.WalkSkipChildren, nil
}

func (r *myRenderer) renderText(w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	if !entering {
		r.currentRun = nil
		return ast.WalkContinue, nil
	}
	n := node.(*ast.Text)
	if r.currentRun == nil {
		run := r.currentPara.AddRun()
		r.currentRun = &run
	}
	segment := n.Segment
	if n.IsRaw() {
		r.currentRun.AddText(string(segment.Value(source)))
	} else {
		value := segment.Value(source)
		r.currentRun.AddText(string(value))

	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderString(w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	if !entering {
		return ast.WalkContinue, nil
	}
	n := node.(*ast.String)
	fmt.Println("String")
	r.currentRun.AddText(string(n.Value))
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderTable(
	w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		table := r.doc.AddTable()
		table.Properties().SetStyle("TableNormal")
		r.currentTable = &table
	} else {
		r.currentTable = nil
	}
	return ast.WalkContinue, nil
}

// TableHeaderAttributeFilter defines attribute names which <thead> elements can have.

func (r *myRenderer) renderTableHeader(
	w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	if entering {
		row := r.currentTable.AddRow()
		r.currentTableRow = &row
	} else {
		r.currentTableRow = nil
	}
	return ast.WalkContinue, nil
}

func (r *myRenderer) renderTableRow(
	w util.BufWriter, source []byte, n ast.Node, entering bool) (ast.WalkStatus, error) {
	return r.renderTableHeader(w, source, n, entering)
}

func (r *myRenderer) renderTableCell(
	w util.BufWriter, source []byte, node ast.Node, entering bool) (ast.WalkStatus, error) {
	// n := node.(*ast.TableCell)
	if entering {
		cell := r.currentTableRow.AddCell()
		r.currentTableCell = &cell
		cell.Properties().Borders().SetAll(wml.ST_BorderBasicWideMidline, color.Black, 1)
		paragraph := cell.AddParagraph()
		r.currentPara = &paragraph
	} else {
		r.currentTableCell = nil
		r.currentPara = nil
	}
	return ast.WalkContinue, nil
}
