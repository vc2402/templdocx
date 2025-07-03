package templdocx

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/fumiama/go-docx"
	"os"
	"reflect"
	"strconv"
	"strings"
	"text/template"
	"unicode"
)

var (
	ErrNotInControlContext       = errors.New("not in control context")
	ErrInvalidControlContext     = errors.New("invalid control context")
	ErrNoClosingCommandFound     = errors.New("no closing command found")
	ErrUnsupportedCommandContext = errors.New("unsupported command context")
	ErrCommandResultError        = errors.New("command result error")
	ErrVarNotFound               = errors.New("variable not found")
	FuncVarNotFound              = errors.New("function not found")
)

const (
	controlStart        = "${{"
	controlEnd          = "}}"
	controlCommandRange = "range"
	controlCommandEnd   = "end"
)

type DocObject int

const (
	DOBody DocObject = iota
	DOParagraph
	DOTable
	DOTableRow
	DOTableCell
	DOControl
)

type Context struct {
	Processor  *DocProcessor
	ObjectType DocObject
	Parent     *Context
	Index      int
	//docx corresponding object
	Item          any
	Vars          any
	counts        map[DocObject]int
	skipExecution bool
}

type ControlObject interface {
	IsTheSame(co ControlObject) bool
	IsClosing(co *ControlTypeEnd) bool
	IsClosed() bool
	AddParam(string) error
	Execute(ctx *Context) ([]any, error)
	Describe() string
}

type CustomControlProvider func(name string) ControlObject

type DrawableProvider func(ctx *Context, params []string) ([]byte, error)

type ControlContext struct {
	Items         []any
	Parent        any
	FirstItemIdx  int
	LastItemIdx   int
	ControlName   string
	ControlObject ControlObject
}

type Error struct {
	TemplateError error
	Message       string
	Position      string
	Token         string
}

type ControlTypeEnd struct {
	command string
	params  []string
}

type controlTypeLoop struct {
	by     string
	closed bool
}

type ControlTypeDrawing struct {
	Provider DrawableProvider
	params   []string
}

type DocProcessor struct {
	doc             *docx.Docx
	values          any
	funcMap         template.FuncMap
	context         *Context
	controlProvider CustomControlProvider
}

func NewDocParser(doc *docx.Docx, values any, options ...any) *DocProcessor {
	dp := &DocProcessor{
		doc:    doc,
		values: values,
	}
	for _, option := range options {
		switch v := option.(type) {
		case CustomControlProvider:
			dp.controlProvider = v
		case template.FuncMap:
			dp.funcMap = v
		}
	}
	dp.pushContext(DOBody, doc.Document.Body)
	return dp
}

func (dp *DocProcessor) Process() error {
	_, err := dp.context.processItems(dp.doc.Document.Body.Items)
	return err
}

func (dp *DocProcessor) pushContext(tip DocObject, item any) *Context {
	ctx := &Context{
		Processor:  dp,
		ObjectType: tip,
		Parent:     dp.context,
		Item:       item,
		counts:     map[DocObject]int{},
	}
	if dp.context != nil {
		ctx.Index = dp.context.counts[tip]
		dp.context.counts[tip] = ctx.Index + 1
		ctx.Vars = dp.context.Vars
		ctx.skipExecution = dp.context.skipExecution
	} else {
		ctx.Vars = dp.values
	}
	dp.context = ctx
	return ctx
}

func (dp *DocProcessor) popContext() *Context {
	ctx := dp.context
	if ctx != nil {
		dp.context = ctx.Parent
	}
	return dp.context
}

func (c *Context) processItems(items []any) (any, error) {
	var err error
	var command any
	for i := 0; i < len(items); i++ {
		it := items[i]
		command, err = c.processItem(it)
		if err != nil {
			return command, err
		}
		if command != nil {
			doNotSkip := false
			if co, ok := command.(ControlObject); ok && co.IsClosed() {
				doNotSkip = true
			}
			if c.skipExecution || c.ObjectType != DOBody && !doNotSkip {
				return command, err
			}
			if co, ok := command.(ControlObject); ok {
				cc := &ControlContext{
					Items:         nil,
					Parent:        c.Item,
					FirstItemIdx:  i + 1,
					LastItemIdx:   i + 1,
					ControlObject: co,
				}
				c = c.Processor.pushContext(
					DOControl,
					cc,
				)
				var restItems []any
				if !co.IsClosed() {
					err = c.SkipTillControlEnd()
					if err != nil {
						return nil, Error{
							TemplateError: nil,
							Message:       err.Error(),
							Position:      c.describePosition(),
						}
					}
				}
				if cc.LastItemIdx+1 <= len(items) {
					restItems = items[cc.LastItemIdx+1:]
				}

				result, err := co.Execute(c)
				c = c.Processor.popContext()
				if err != nil {
					return nil,
						Error{
							TemplateError: nil,
							Message:       err.Error(),
							Position:      c.describePosition(),
						}
				}
				var children *[]any
				switch c.ObjectType {
				case DOBody:
					children = &c.Processor.doc.Document.Body.Items
					//c.Processor.doc.Document.Body.Items = items[0:i]
					//c.Processor.doc.Document.Body.Items = append(c.Processor.doc.Document.Body.Items, result...)
					//i = len(c.Processor.doc.Document.Body.Items)
					//if len(restItems) > 0 {
					//  c.Processor.doc.Document.Body.Items = append(c.Processor.doc.Document.Body.Items, restItems...)
					//}
					//items = c.Processor.doc.Document.Body.Items
				case DOParagraph:
					para := c.Item.(*docx.Paragraph)
					children = &para.Children
				default:
					return nil, fmt.Errorf("unexpected object type: %v", c.ObjectType)
				}
				*children = items[0:i]
				*children = append(*children, result...)
				i = len(*children)
				if len(restItems) > 0 {
					*children = append(*children, restItems...)
				}
				items = *children
			}
		}
	}
	return nil, nil
}

func (c *Context) processItem(it any) (any, error) {
	var err error
	var command any
	switch v := it.(type) {
	case *docx.Paragraph:
		command, err = c.processParagraph(v)
	case *docx.Table:
		err = c.processTable(v)
	case *docx.Text:
		command, err = c.isControl(v)
		if command == nil && err == nil {
			err = c.processText(v)
		}
	case *docx.Run:
		command, err = c.processItems(v.Children)
	case *docx.WTableRow:
		for _, cell := range v.TableCells {
			command, err = c.processItem(cell)
			if err != nil || command != nil {
				return command, err
			}
		}
	case *docx.WTableCell:
		for _, paragraph := range v.Paragraphs {
			command, err = c.processParagraph(paragraph)
			if err != nil || command != nil {
				return command, err
			}
		}
	case docx.Body:
		command, err = c.processItems(v.Items)
		// TODO add context and process command
	}
	return command, err
}

func (c *Context) processTable(table *docx.Table) error {
	c = c.Processor.pushContext(DOTable, table)
	defer c.Processor.popContext()
	rows := table.TableRows
	for rowIdx := 0; rowIdx < len(rows); rowIdx++ {
		row := rows[rowIdx]
		c = c.Processor.pushContext(DOTableRow, row)
		for cellIdx := 0; cellIdx < len(row.TableCells); cellIdx++ {
			cell := row.TableCells[cellIdx]
			c = c.Processor.pushContext(DOTableCell, cell)
			command, err := c.processItem(cell)
			if err != nil {
				c.Processor.popContext() // cell
				c.Processor.popContext() // row
				return err
			}
			if command != nil {
				if co, ok := command.(ControlObject); ok {
					if len(row.TableCells) > 1 {
						row.TableCells = append(row.TableCells[0:cellIdx], row.TableCells[cellIdx+1:len(row.TableCells)]...)
						cellIdx--
						//TODO add context and process command
					} else {
						cc := &ControlContext{
							Items:         nil,
							Parent:        table,
							FirstItemIdx:  rowIdx + 1,
							ControlObject: co,
						}
						c = c.Processor.pushContext(
							DOControl,
							cc,
						)
						err = c.SkipTillControlEnd()
						if err != nil {
							c.Processor.popContext() // command
							c.Processor.popContext() // cell
							c.Processor.popContext() // row
							return err
						}
						restRows := table.TableRows[cc.LastItemIdx+1:]
						table.TableRows = table.TableRows[0:rowIdx]

						result, err := co.Execute(c)
						c = c.Processor.popContext()
						if err != nil {
							err = Error{
								TemplateError: nil,
								Message:       err.Error(),
								Position:      c.describePosition(),
							}
							c.Processor.popContext()
							c.Processor.popContext()
							return err
						}
						for _, r := range result {
							row, ok := r.(*docx.WTableRow)
							if !ok {
								c.Processor.popContext()
								c.Processor.popContext()
								return fmt.Errorf("%w: expected *docx.WTableRow but got %T", ErrCommandResultError, r)
							}
							table.TableRows = append(table.TableRows, row)
						}
						rowIdx = len(table.TableRows)
						if len(restRows) > 0 {
							table.TableRows = append(table.TableRows, restRows...)
						}
						c = c.Processor.popContext()
						break
					}
				}
			}
			c = c.Processor.popContext()
		}
		c = c.Processor.popContext()
	}
	return nil
}

func (c *Context) processParagraph(paragraph *docx.Paragraph) (any, error) {
	c = c.Processor.pushContext(DOParagraph, paragraph)
	defer c.Processor.popContext()
	return c.processItems(paragraph.Children)
}

func (c *Context) processText(text *docx.Text) error {
	templ := template.New("text")
	if c.Processor.funcMap != nil {
		templ.Funcs(c.Processor.funcMap)
	}
	templ, err := templ.Parse(text.Text)
	if err != nil {
		return Error{
			TemplateError: err,
			Message:       "while parsing",
			Token:         text.Text,
			Position:      c.describePosition(),
		}
	}

	result := strings.Builder{}
	err = templ.Execute(&result, c.Vars)
	if err != nil {
		return Error{
			TemplateError: err,
			Message:       "while executing",
			Token:         text.Text,
			Position:      c.describePosition(),
		}
	}
	text.Text = result.String()
	return nil
}

// SkipTillControlEnd reads dic till the end of current command
func (c *Context) SkipTillControlEnd() error {
	if c.ObjectType != DOControl {
		return ErrNotInControlContext
	}
	c.skipExecution = true
	cc, ok := c.Item.(*ControlContext)
	if !ok {
		return ErrInvalidControlContext
	}
	var err error
	parent := cc.Parent
	switch v := parent.(type) {
	case *docx.Table:
		err = c.skipTillControlEndTable(cc, v)
	case docx.Body:
		err = c.skipTillControlEndBody(cc, v)
	default:
		err = ErrUnsupportedCommandContext
	}
	c.skipExecution = false
	return err
}

func (c *Context) skipTillControlEndTable(cc *ControlContext, table *docx.Table) error {
	count := 1
	for cc.LastItemIdx = cc.FirstItemIdx; cc.LastItemIdx < len(table.TableRows); cc.LastItemIdx++ {
		// let's clone it because we need to leave templates
		cloned, err := c.clone(table.TableRows[cc.LastItemIdx])
		if err != nil {
			return err
		}
		row := cloned.(*docx.WTableRow)
		if len(row.TableCells) > 0 {
			// checking only first cell
			cell := row.TableCells[0]
			command, err := c.processItem(cell)
			if err != nil {
				return err
			}
			if command != nil {
				if closing, ok := command.(*ControlTypeEnd); ok {
					if cc.ControlObject.IsClosing(closing) {
						count--
						if count == 0 {
							return nil
						}
					}
				}
				if co, ok := command.(ControlObject); ok {
					if co.IsTheSame(cc.ControlObject) {
						count++
					}
				}
			}
		}
		cc.Items = append(cc.Items, table.TableRows[cc.LastItemIdx])
	}
	return Error{
		TemplateError: ErrNoClosingCommandFound,
		Message:       "for command ",
		Token:         cc.ControlObject.Describe(),
		Position:      c.describePosition(),
	}
}

func (c *Context) skipTillControlEndBody(cc *ControlContext, body docx.Body) error {
	count := 1
	for cc.LastItemIdx = cc.FirstItemIdx; cc.LastItemIdx < len(body.Items); cc.LastItemIdx++ {
		it := body.Items[cc.LastItemIdx]
		// let's clone it because we need to leave templates
		cloned, err := c.clone(it)
		if err != nil {
			return err
		}
		command, err := c.processItem(cloned)
		if err != nil {
			return err
		}
		if command != nil {
			if closing, ok := command.(*ControlTypeEnd); ok {
				if cc.ControlObject.IsClosing(closing) {
					count--
					if count == 0 {
						return nil
					}
				}
			}
			if co, ok := command.(ControlObject); ok {
				if co.IsTheSame(cc.ControlObject) {
					count++
				}
			}

		}
		cc.Items = append(cc.Items, it)
	}
	return Error{
		TemplateError: ErrNoClosingCommandFound,
		Message:       "for command ",
		Token:         cc.ControlObject.Describe(),
		Position:      c.describePosition(),
	}
}

func (c *Context) clone(item any) (any, error) {
	var err error
	switch v := item.(type) {
	case *docx.Paragraph:
		ret := &docx.Paragraph{}
		*ret = *v
		ret.Children = make([]any, len(v.Children))
		for i, child := range v.Children {
			ret.Children[i], err = c.clone(child)
		}
		return ret, err
	case *docx.Table:
		ret := &docx.Table{}
		*ret = *v
		ret.TableRows = make([]*docx.WTableRow, len(v.TableRows))
		for i, child := range v.TableRows {
			var row any
			row, err = c.clone(child)
			if err != nil {
				break
			}
			ret.TableRows[i] = row.(*docx.WTableRow)
		}
		return ret, err
	case *docx.Text:
		ret := &docx.Text{}
		*ret = *v
		return ret, err
	case *docx.Run:
		ret := &docx.Run{}
		*ret = *v
		ret.Children = make([]any, len(v.Children))
		for i, child := range v.Children {
			ret.Children[i], err = c.clone(child)
		}
		return ret, err
	case *docx.WTableRow:
		ret := &docx.WTableRow{}
		*ret = *v
		ret.TableCells = make([]*docx.WTableCell, len(v.TableCells))
		for i, child := range v.TableCells {
			var row any
			row, err = c.clone(child)
			if err != nil {
				break
			}
			ret.TableCells[i] = row.(*docx.WTableCell)
		}
		return ret, err
	case *docx.WTableCell:
		ret := &docx.WTableCell{}
		*ret = *v
		ret.Tables = make([]*docx.Table, len(v.Tables))
		for i, child := range v.Tables {
			var row any
			row, err = c.clone(child)
			if err != nil {
				break
			}
			ret.Tables[i] = row.(*docx.Table)
		}
		if err != nil {
			return ret, err
		}
		ret.Paragraphs = make([]*docx.Paragraph, len(v.Paragraphs))
		for i, child := range v.Paragraphs {
			var row any
			row, err = c.clone(child)
			if err != nil {
				break
			}
			ret.Paragraphs[i] = row.(*docx.Paragraph)
		}
		return ret, err
	case *docx.Drawing:
		//let's leave it as is so far
		return item, nil
	default:
		//let's leave it as is so far
		return item, nil
		//return nil, errors.New("unsupported type")
	}
}

func (c *Context) isControl(text *docx.Text) (any, error) {
	t := strings.TrimSpace(text.Text)
	if strings.HasPrefix(t, controlStart) && strings.HasSuffix(t, controlEnd) {
		var co ControlObject
		var ce *ControlTypeEnd
		token := strings.Builder{}
		t = strings.TrimPrefix(t, controlStart)
		t = strings.TrimSuffix(t, controlEnd)
		finalize := func() error {
			var err error
			if co != nil {
				err = co.AddParam(token.String())
			} else if ce != nil {
				if ce.command == "" {
					ce.command = token.String()
				} else {
					ce.params = append(ce.params, token.String())
				}
			} else {
				command := token.String()
				switch command {
				case controlCommandRange:
					co = &controlTypeLoop{}
				case controlCommandEnd:
					ce = &ControlTypeEnd{}
				default:
					if c.Processor.controlProvider != nil {
						co = c.Processor.controlProvider(command)
					}
					if co == nil {
						err = errors.New(fmt.Sprintf("unknown control command: %s", command))
					}
				}
			}
			token.Reset()
			return err
		}
		for _, c := range t {
			if unicode.IsSpace(c) {
				if token.Len() == 0 {
					continue
				}
				err := finalize()
				if err != nil {
					return nil, err
				}
			} else {
				token.WriteRune(c)
			}
		}
		err := finalize()
		if err != nil {
			return nil, err
		}
		if co != nil {
			return co, nil
		}
		return ce, nil
	}
	return nil, nil
}

func (e Error) Error() string {
	var ret strings.Builder
	if e.Position != "" {
		ret.WriteString("at ")
		ret.WriteString(e.Position)
	}
	if e.Token != "" {
		ret.WriteString("current token: '")
		ret.WriteString(e.Token)
		ret.WriteString("': ")
	}
	if e.Message != "" {
		ret.WriteString(e.Message)
	}
	if e.TemplateError != nil {
		if e.Message != "" {
			ret.WriteString(": ")
		}
		ret.WriteString(e.TemplateError.Error())
	}
	return ret.String()
}

func (c *Context) describePosition() string {
	if c == nil {
		return ""
	}
	var ret strings.Builder

	var walk func(c *Context)
	walk = func(c *Context) {
		if c.Parent != nil {
			walk(c.Parent)
		}
		switch c.ObjectType {
		case DOParagraph:
			ret.WriteString("paragraph ")
			ret.WriteString(strconv.Itoa(c.Index + 1))
			ret.WriteString(": ")
		case DOTable:
			ret.WriteString("table ")
			ret.WriteString(strconv.Itoa(c.Index + 1))
			ret.WriteString(": ")
		case DOTableRow:
			ret.WriteString("row ")
			ret.WriteString(strconv.Itoa(c.Index + 1))
			ret.WriteString(": ")
		case DOTableCell:
			ret.WriteString("cell ")
			ret.WriteString(strconv.Itoa(c.Index + 1))
			ret.WriteString(": ")
		case DOBody:
			ret.WriteString("paragraph ")
			ret.WriteString(strconv.Itoa(c.counts[DOParagraph] + 1))
			ret.WriteString(": ")
		}
	}
	walk(c)
	return ret.String()
}

func (c *Context) parseVar(name string, params ...string) (any, error) {
	if strings.HasPrefix(name, ".") {
		return c.getVar(name[1:])
	}
	if c.Processor.funcMap == nil {
		return nil, fmt.Errorf("when parsing '%s': FuncMap is null", name)
	}
	return c.executeFunc(name, params...)
}

func (c *Context) getVar(name string) (any, error) {
	// todo process dots
	if asMap, ok := c.Vars.(map[string]any); ok {
		val, ok := asMap[name]
		if !ok {
			return nil, fmt.Errorf("%w: '%s'", ErrVarNotFound, name)
		}
		return val, nil
	}
	rv := reflect.ValueOf(c.Vars)
	if rv.Kind() != reflect.Struct {
		if rv.FieldByName(name).IsValid() {
			return rv.FieldByName(name).Interface(), nil
		}
	}
	return nil, fmt.Errorf("%w: '%s'", ErrVarNotFound, name)
}

func (c *Context) executeFunc(name string, params ...string) (any, error) {
	fun, ok := c.Processor.funcMap[name]
	if !ok {
		return nil, fmt.Errorf("%w: '%s'", FuncVarNotFound, name)
	}
	rv := reflect.ValueOf(fun)
	if rv.Kind() == reflect.Func {
		funcParams := make([]reflect.Value, len(params))
		for i, p := range params {
			//TODO process params types
			funcParams[i] = reflect.ValueOf(p)
		}
		results := rv.Call(funcParams)
		// TODO check if it is really error
		if len(results) > 1 && results[1].Interface() != nil {
			return nil, results[1].Interface().(error)
		}
		return results[0].Interface(), nil
	}
	return nil, fmt.Errorf("%w: '%s' is not a function", FuncVarNotFound, name)
}

func (c *controlTypeLoop) IsClosing(ctl *ControlTypeEnd) bool {
	if ctl.command == controlCommandRange && len(ctl.params) > 0 && ctl.params[0] == c.by {
		c.closed = true
		return true
	}
	return false
}

func (c *controlTypeLoop) IsClosed() bool {
	return c.closed
}

func (c *controlTypeLoop) IsTheSame(co ControlObject) bool {
	if ctl, ok := co.(*controlTypeLoop); ok && ctl.by == c.by {
		return true
	}
	return false
}

func (c *controlTypeLoop) AddParam(param string) error {
	if c.by != "" {
		return errors.New("too many params for loop")
	}
	c.by = param
	return nil
}

func (c *controlTypeLoop) Execute(ctx *Context) ([]any, error) {
	cc, ok := ctx.Item.(*ControlContext)
	if !ok {
		return nil, errors.New("invalid control type")
	}
	val, err := ctx.parseVar(c.by)
	if err != nil {
		return nil, err
	}

	ret := make([]any, 0, len(cc.Items))
	process := func() error {
		for _, it := range cc.Items {
			switch v := it.(type) {
			case *docx.WTableRow:
				cloned, err := ctx.clone(v)
				if err != nil {
					return err
				}
				clonedRow := cloned.(*docx.WTableRow)
				_, err = ctx.processItem(clonedRow)
				if err != nil {
					return err
				}
				ret = append(ret, clonedRow)
			default:
				cloned, err := ctx.clone(v)
				if err != nil {
					return err
				}
				_, err = ctx.processItem(cloned)
				if err != nil {
					return err
				}
				ret = append(ret, cloned)
			}
		}
		return nil
	}
	switch arr := val.(type) {
	case []any:
		for _, it := range arr {
			ctx.Vars = it
			err = process()
			if err != nil {
				return nil, err
			}
		}
	default:
		return nil, fmt.Errorf("unsupported array type: %T", val)
	}
	return ret, nil
}

func (c *controlTypeLoop) Describe() string {
	return fmt.Sprintf("%s %s", controlCommandRange, c.by)
}

func (c *ControlTypeDrawing) IsTheSame(co ControlObject) bool {
	return false
}

func (c *ControlTypeDrawing) IsClosing(co *ControlTypeEnd) bool {
	return false
}

func (c *ControlTypeDrawing) IsClosed() bool {
	return true
}

func (c *ControlTypeDrawing) AddParam(s string) error {
	c.params = append(c.params, s)
	return nil
}

func (c *ControlTypeDrawing) Execute(ctx *Context) ([]any, error) {
	switch ctx.Parent.ObjectType {
	case DOParagraph:
		dp := c.Provider
		if dp == nil {
			dp = FSDrawableProvider
		}
		para := ctx.Parent.Item.(*docx.Paragraph)
		bytes, err := dp(ctx, c.params)
		run, err := para.AddInlineDrawing(bytes)
		if err != nil {
			return nil, err
		}
		return []any{run}, nil
	default:
		return nil, errors.New("invalid control type")
	}
}

func (c *ControlTypeDrawing) Describe() string {
	return fmt.Sprintf("drawing: %v", c.params)
}

func FSDrawableProvider(ctx *Context, params []string) ([]byte, error) {
	if len(params) == 0 {
		return nil, errors.New("image file name not given")
	}
	paramValue, err := ctx.parseVar(params[0])
	if err != nil {
		return nil, err
	}
	fileName, ok := paramValue.(string)
	if !ok {
		return nil, fmt.Errorf("%w: '%s'", ErrVarNotFound, fileName)
	}
	file, err := os.Open(fileName)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	buf := bytes.Buffer{}
	_, err = file.WriteTo(&buf)
	return buf.Bytes(), nil
}
