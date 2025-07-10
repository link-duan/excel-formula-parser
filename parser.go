package excelformulaparser

import (
	"errors"
	"io"
	"strconv"
)

type Parser struct {
	lexer     *lexer
	token     *Token
	lookahead *Token
}

func NewParser(src string) *Parser {
	return &Parser{
		lexer: newLexer(src),
	}
}

func (p *Parser) Parse() (Node, error) {
	if err := p.advance(); err != nil {
		return nil, err
	}
	if p.token == nil {
		return nil, nil // No tokens to parse
	}
	if p.token.Type == Equal {
		if err := p.advance(); err != nil {
			return nil, err
		}
	}
	res, err := p.comparison()
	if err != nil {
		return nil, err
	}
	if p.token != nil {
		return nil, newParseError(p.token.Start, "unexpected token at end %s %v", p.token.Raw, p.token.Type)
	}
	return res, nil
}

func (p *Parser) comparison() (Node, error) {
	var left, err = p.connection()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
LOOP:
	for p.token != nil {
		switch p.token.Type {
		case Equal, NotEqual, LessThan, GreaterThan, LessThanOrEqual, GreaterThanOrEqual:
			var op = p.token
			if err := p.advance(); err != nil { // consume the operator token
				return nil, err
			}
			var right, err = p.connection()
			if err != nil {
				return nil, err
			}
			left = &BinaryExpr{
				baseNode: newBaseNode(left.Start(), right.End()),
				Left:     left,
				Operator: op,
				Right:    right,
			}
		default:
			break LOOP
		}
	}
	return left, nil
}

func (p *Parser) connection() (Node, error) {
	var left, err = p.addition()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
	for p.token != nil && p.token.Type == Concat {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		var right, err = p.addition()
		if err != nil {
			return nil, err
		}
		left = &BinaryExpr{
			baseNode: newBaseNode(left.Start(), right.End()),
			Left:     left,
			Operator: op,
			Right:    right,
		}
	}
	return left, nil
}

func (p *Parser) addition() (Node, error) {
	var left, err = p.muliplication()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
	for p.token != nil && (p.token.Type == Plus || p.token.Type == Minus) {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		var right, err = p.muliplication()
		if err != nil {
			return nil, err
		}
		left = &BinaryExpr{
			baseNode: newBaseNode(left.Start(), right.End()),
			Left:     left,
			Operator: op,
			Right:    right,
		}
	}
	return left, nil
}

func (p *Parser) muliplication() (Node, error) {
	var left, err = p.exponentiation()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
	for p.token != nil && (p.token.Type == Multiply || p.token.Type == Divide) {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		var right, err = p.exponentiation()
		if err != nil {
			return nil, err
		}
		left = &BinaryExpr{
			baseNode: newBaseNode(left.Start(), right.End()),
			Left:     left,
			Operator: op,
			Right:    right,
		}
	}
	return left, nil
}

func (p *Parser) exponentiation() (Node, error) {
	var left, err = p.percent()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
	for p.token != nil && p.token.Type == Exponentiation {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		var right, err = p.percent()
		if err != nil {
			return nil, err
		}
		left = &BinaryExpr{
			baseNode: newBaseNode(left.Start(), right.End()),
			Left:     left,
			Operator: op,
			Right:    right,
		}
	}
	return left, nil
}

func (p *Parser) percent() (Node, error) {
	var left, err = p.negation()
	if err != nil {
		return nil, err
	}
	for p.token != nil && p.token.Type == Percent {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		left = &UnaryExpr{
			baseNode: newBaseNode(left.Start(), op.End),
			Operator: op,
			Operand:  left,
		}
	}
	return left, nil
}

func (p *Parser) negation() (Node, error) {
	if p.token != nil && p.token.Type == Minus {
		var op = p.token
		if err := p.advance(); err != nil { // consume the operator token
			return nil, err
		}
		var right, err = p.negation()
		if err != nil {
			return nil, err
		}
		return &UnaryExpr{
			baseNode: newBaseNode(op.Start, op.End),
			Operator: op,
			Operand:  right,
		}, nil
	}
	return p.rangeExpr()
}

func (p *Parser) rangeExpr() (Node, error) {
	var left, err = p.primary()
	if err != nil {
		return nil, err
	}
	if p.token == nil {
		return left, nil // No more tokens, return the left node
	}
	if p.token != nil && p.token.Type == Colon {
		var begin, err = p.tryConvertToCellExpr(left)
		if err != nil {
			return nil, err
		}
		var colons []*Token
		var ends []Node
		for p.token != nil && p.token.Type == Colon {
			var op = p.token
			if err := p.advance(); err != nil { // consume the ':' token
				return nil, err
			}
			colons = append(colons, op)
			var right, err = p.primary()
			if err != nil {
				return nil, err
			}
			right, err = p.tryConvertToCellExpr(right)
			if err != nil {
				return nil, err
			}
			ends = append(ends, right)
		}
		return &RangeExpr{
			baseNode: newBaseNode(begin.Start(), ends[len(ends)-1].End()),
			Begin:    begin,
			Colons:   colons,
			Ends:     ends,
		}, nil
	}
	return left, nil
}

func (p *Parser) tryConvertToCellExpr(node Node) (*CellExpr, error) {
	if node, ok := node.(*CellExpr); ok {
		return node, nil // Already a CellExpr, no conversion needed
	}
	if node, ok := node.(*IdentExpr); ok {
		col := colNameToIndex(node.Name.Raw)
		if col < 0 {
			return nil, newParseError(node.Start(), "expected a cell reference: %s", node.Name.Raw)
		}
		return &CellExpr{
			baseNode: newBaseNode(node.Start(), node.End()),
			Ident:    node.Name,
			Row:      -1, // -1 indicates a full column reference
			Col:      col,
		}, nil
	}
	if node, ok := node.(*LiteralExpr); ok {
		if node.Value.Type != Number {
			return nil, newParseError(node.Start(), "expected a row reference, got %s", node.Value.Raw)
		}
		row, err := strconv.Atoi(node.Value.Raw)
		if err != nil {
			return nil, newParseError(node.Start(), "invalid row reference: %s", err.Error())
		}
		return &CellExpr{
			baseNode: newBaseNode(node.Start(), node.End()),
			Ident:    node.Value,
			Row:      row - 1, // Convert to zero-based index
			Col:      -1,      // -1 indicates a full row reference
		}, nil
	}
	return nil, newParseError(node.Start(), "expected a cell reference, got %T", node)
}

// Primary expressions are the basic building blocks of expressions.
// They can be:
//
//   - Literals (numbers, strings, booleans, etc.)
//
//   - Identifiers (variable or function names)
//
//   - Parenthesized expressions (expressions inside parentheses, e.g., (a + b))
//
//   - Possibly function calls or array accesses, depending on the language
func (p *Parser) primary() (Node, error) {
	if p.token == nil {
		return nil, newParseError(p.lexer.pos, "unexpected end of input")
	}
	tk := p.token
	switch p.token.Type {
	case ParenOpen:
		return p.parenthesized()
	case String, Number, BoolLiteral:
		if err := p.advance(); err != nil { // consume the token
			return nil, err
		}
		return &LiteralExpr{
			baseNode: newBaseNode(tk.Start, tk.End),
			Value:    tk,
		}, nil
	case Ident:
		var peek, err = p.peek()
		if err != nil {
			return nil, err
		}
		if peek != nil && peek.Type == ParenOpen {
			return p.functionCall()
		}
		if err := p.advance(); err != nil { // consume the token
			return nil, err
		}
		return &IdentExpr{
			baseNode: newBaseNode(tk.Start, tk.End),
			Name:     tk,
		}, nil
	case Cell:
		// parse the cell token to extract row and column information
		result, err := parseCell(tk.Raw)
		if err != nil {
			return nil, newParseError(tk.Start, "invalid cell reference: %s", err.Error())
		}
		if err := p.advance(); err != nil { // consume the token
			return nil, err
		}
		return &CellExpr{
			baseNode:    newBaseNode(tk.Start, tk.End),
			Ident:       tk,
			Row:         result.row,
			RowAbsolute: result.rowAbsolute,
			Col:         result.col,
			ColAbsolute: result.colAbsolute,
		}, nil
	case AbsoluteRow:
		row, err := strconv.Atoi(tk.Raw[1:]) // Skip the '$' character
		if err != nil {
			return nil, newParseError(tk.Start, "invalid absolute row reference: %s", err.Error())
		}
		if err := p.advance(); err != nil { // consume the token
			return nil, err
		}
		return &CellExpr{
			baseNode:    newBaseNode(tk.Start, tk.End),
			Ident:       tk,
			Row:         row - 1,
			Col:         -1, // -1 indicates a full row reference
			RowAbsolute: true,
		}, nil
	case AbsoluteColumn:
		colName := tk.Raw[1:] // Skip the '$' character
		col := colNameToIndex(colName)
		if col < 0 {
			return nil, newParseError(tk.Start, "invalid absolute column reference: %s", tk.Raw)
		}
		if err := p.advance(); err != nil { // consume the token
			return nil, err
		}
		return &CellExpr{
			baseNode:    newBaseNode(tk.Start, tk.End),
			Ident:       tk,
			Row:         -1, // -1 indicates a full column reference
			Col:         col,
			ColAbsolute: true,
		}, nil
	default:
		return nil, newParseError(tk.Start, "unexpected token %s", tk.Raw)
	}
}

func (p *Parser) functionCall() (Node, error) {
	if p.token == nil {
		return nil, newParseError(p.lexer.pos, "unexpected end of input")
	}
	if p.token.Type != Ident {
		return nil, newParseError(p.token.Start, "expected function name")
	}
	var name = p.token
	if err := p.advance(); err != nil { // consume the function name token
		return nil, err
	}
	if p.token == nil || p.token.Type != ParenOpen {
		return nil, newParseError(name.Start, "expected '(' after function name")
	}
	var paranOpen = p.token
	if err := p.advance(); err != nil { // consume the '(' token
		return nil, err
	}
	var arguments []Node
	for p.token != nil && p.token.Type != ParenClose {
		var arg, err = p.comparison()
		if err != nil {
			return nil, err
		}
		arguments = append(arguments, arg)
		if p.token == nil {
			return nil, newParseError(name.Start, "unexpected end of input, expected ')' to close function call")
		}
		if p.token.Type == Comma {
			if err := p.advance(); err != nil { // consume the ',' token
				return nil, err
			}
			if p.token == nil || p.token.Type == ParenClose {
				return nil, newParseError(name.Start, "unexpected end of input, expected argument after ','")
			}
		}
	}
	if p.token == nil || p.token.Type != ParenClose {
		return nil, newParseError(p.lexer.pos, "expected ')' to close function call")
	}
	var paranClose = p.token
	if err := p.advance(); err != nil { // consume the ')' token
		return nil, err
	}
	return &FunCallExpr{
		baseNode:   newBaseNode(name.Start, paranClose.End),
		Name:       name,
		ParanOpen:  paranOpen,
		Arguments:  arguments,
		ParanClose: paranClose,
	}, nil
}

func (p *Parser) parenthesized() (Node, error) {
	if p.token == nil || p.token.Type != ParenOpen {
		return nil, newParseError(p.token.Start, "expected '('")
	}
	var parenOpen = p.token
	var start = p.token.Start
	if err := p.advance(); err != nil { // consume the '(' token
		return nil, err
	}
	var expr, err = p.comparison()
	if err != nil {
		return nil, err
	}
	if p.token == nil || p.token.Type != ParenClose {
		return nil, newParseError(p.token.Start, "expected ')'")
	}
	var parenClose = p.token
	var end = p.token.End
	if err := p.advance(); err != nil { // consume the ')' token
		return nil, err
	}
	return &ParenthesizedExpr{
		baseNode:   newBaseNode(start, end),
		ParenOpen:  parenOpen,
		Inner:      expr,
		ParenClose: parenClose,
	}, nil
}

func (p *Parser) advance() error {
	if p.lookahead != nil {
		p.token = p.lookahead
		p.lookahead = nil
		return nil
	}
	token, err := p.lexer.next()
	if err != nil {
		if errors.Is(err, io.EOF) {
			p.token = nil
			return nil
		}
		return err
	}
	p.token = token
	return nil
}

func (p *Parser) peek() (*Token, error) {
	if p.lookahead != nil {
		return p.lookahead, nil
	}
	token, err := p.lexer.next()
	if err != nil {
		if errors.Is(err, io.EOF) {
			return nil, nil
		}
		return nil, err
	}
	p.lookahead = token
	return p.lookahead, nil
}

type parseCellResult struct {
	row         int
	col         int
	rowAbsolute bool
	colAbsolute bool
}

func parseCell(cell string) (result *parseCellResult, err error) {
	var runes = []rune(cell)
	if len(runes) == 0 {
		return nil, errors.New("cell reference cannot be empty")
	}
	var offset = 0
	result = &parseCellResult{}
	if runes[offset] == '$' {
		result.rowAbsolute = true
		offset++ // consume the '$'
	}
	var rowStarts = offset
	var rowEnds = offset
	for offset < len(runes) && isASCIILetter(runes[offset]) {
		offset += 1
		rowEnds += 1
	}
	if rowEnds == rowStarts {
		return nil, errors.New("invalid cell reference: no column specified")
	}
	if offset >= len(runes) {
		return nil, errors.New("invalid cell reference: no row specified")
	}
	if runes[offset] == '$' {
		result.colAbsolute = true
		offset++ // consume the '$'
	}
	var colStarts = offset
	var colEnds = offset
	for offset < len(runes) && isDigit(runes[offset]) {
		offset += 1
		colEnds += 1
	}
	if colEnds == colStarts {
		return nil, errors.New("invalid cell reference: no row specified")
	}
	if offset < len(runes) {
		return nil, errors.New("invalid cell reference: extra characters after row")
	}
	result.col = colNameToIndex(string(runes[colStarts:colEnds]))
	row, _ := strconv.Atoi(string(runes[rowStarts:rowEnds]))
	result.row = row - 1 // Convert to zero-based index
	return
}

func colNameToIndex(name string) int {
	if len(name) == 0 {
		return -1 // Invalid column name
	}
	var index = 0
	for _, r := range name {
		if !isASCIILetter(r) {
			return -1 // Invalid character in column name
		}
		index = index*26 + int(r-'A'+1)
	}
	return index - 1 // Convert to zero-based index
}
