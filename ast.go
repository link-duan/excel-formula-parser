package excelformulaparser

import (
	"fmt"
	"strconv"
	"strings"
)

type Token struct {
	Start, End Pos
	Type       TokenType
	Raw        string // The raw string value of the token
}

func newToken(start, end Pos, t TokenType, raw string) *Token {
	return &Token{
		Start: start,
		End:   end,
		Type:  t,
		Raw:   raw,
	}
}

type TokenType int

const (
	Ident TokenType = iota + 1
	Number
	String
	BoolLiteral          // TRUE or FALSE
	EValue               // Error value (e.g., #DIV/0!, #VALUE!, etc)
	Exclamation          // !
	BraceOpen            // {
	BraceClose           // }
	BracketOpen          // [
	BracketClose         // ]
	ParenOpen            // (
	ParenClose           // )
	Comma                // ,
	Semicolon            // ;
	ImplicitIntersection // @
	Percent              // %
	Exponentiation       // ^
	Multiply             // *
	Divide               // /
	Plus                 // +
	Colon                // :
	Minus                // -
	Concat               // &
	Equal                // =
	NotEqual             // <>
	LessThan             // <
	GreaterThan          // >
	LessThanOrEqual      // <=
	GreaterThanOrEqual   // >=
	Cell                 // e.g., A1, B2, etc
	AbsoluteRow          // Absolute row reference (e.g., $1, $2)
	AbsoluteColumn       // Absolute column reference (e.g., $A, $B)
)

type Pos struct {
	Line, Column int
}

func (p *Pos) nextLine() {
	p.Line++
	p.Column = 1
}

func (p *Pos) nextColumn() {
	p.Column++
}

func (p Pos) String() string {
	return "(" + strconv.Itoa(p.Line) + "," + strconv.Itoa(p.Column) + ")"
}

func (p Pos) Left() Pos {
	if p.Column == 1 {
		return Pos{Line: p.Line - 1, Column: 0} // Previous line, column 0
	}
	return Pos{Line: p.Line, Column: p.Column - 1}
}

type Node interface {
	Start() Pos           // Returns the position of the expression in the source code
	End() Pos             // Returns the end position of the expression in the source code
	String() string       // Returns a string representation of the node
	cannotBeImplemented() // Ensure this interface cannot be implemented by other types
}

type baseNode struct {
	start Pos
	end   Pos
}

func (e baseNode) Start() Pos {
	return e.start
}

func (e baseNode) End() Pos {
	return e.end
}

func (e baseNode) cannotBeImplemented() {}

func newBaseNode(start, end Pos) baseNode {
	return baseNode{
		start: start,
		end:   end,
	}
}

var _ Node = (*FunCallExpr)(nil)
var _ Node = (*BinaryExpr)(nil)
var _ Node = (*UnaryExpr)(nil)
var _ Node = (*LiteralExpr)(nil)
var _ Node = (*IdentExpr)(nil)
var _ Node = (*RangeExpr)(nil)

type FunCallExpr struct {
	baseNode
	Name       *Token // Function name token
	ParanOpen  *Token
	Arguments  []Node // Arguments can be any expression type
	ParanClose *Token
}

func (f FunCallExpr) String() string {
	var sb strings.Builder
	sb.WriteString("FunCallExpr(Name: ")
	sb.WriteString(f.Name.Raw)
	sb.WriteString(", Arguments: [")
	for i, arg := range f.Arguments {
		if i > 0 {
			sb.WriteString(", ")
		}
		sb.WriteString(arg.String())
	}
	sb.WriteString("])")
	return sb.String()
}

type BinaryExpr struct {
	baseNode
	Left     Node   // Left operand
	Operator *Token // Operator token (e.g., +, -, *, /)
	Right    Node   // Right operand
}

func (b *BinaryExpr) String() string {
	return fmt.Sprintf("BinaryExpr(Left: %s, Operator: %s, Right: %s)", b.Left, b.Operator.Raw, b.Right)
}

type UnaryExpr struct {
	baseNode
	Operator *Token // Operator token (e.g., !, -)
	Operand  Node   // Operand expression
}

func (u UnaryExpr) String() string {
	return fmt.Sprintf("UnaryExpr(Operator: %s, Operand: %s)", u.Operator.Raw, u.Operand)
}

type LiteralExpr struct {
	baseNode
	Value *Token // Token representing the literal value (e.g., number, string, boolean)
}

func (l LiteralExpr) String() string {
	return fmt.Sprintf("LiteralExpr(Value: %s)", l.Value.Raw)
}

type IdentExpr struct {
	baseNode
	Name *Token
}

func (i IdentExpr) String() string {
	return fmt.Sprintf("IdentExpr(Name: %s)", i.Name.Raw)
}

type ParenthesizedExpr struct {
	baseNode
	ParenOpen  *Token // The opening parenthesis
	Inner      Node   // The expression inside the parentheses
	ParenClose *Token // The closing parenthesis
}

func (p ParenthesizedExpr) String() string {
	return fmt.Sprintf("ParenthesizedExpr(Inner: %s)", p.Inner)
}

type RangeExpr struct {
	baseNode

	Begin  Node     // Start of the range (e.g., A1, A, 1)
	Colons []*Token // Colon tokens (e.g., : between A1 and B2)
	Ends   []Node   // End of the range (e.g., B2, B, 2)
}

func (r RangeExpr) String() string {
	var sb strings.Builder
	sb.WriteString("RangeExpr(")
	sb.WriteString(r.Begin.String())
	for _, end := range r.Ends {
		sb.WriteString(":")
		sb.WriteString(end.String())
	}
	sb.WriteString(")")
	return sb.String()
}

type CellExpr struct {
	baseNode
	Ident       *Token // The cell identifier (e.g., A1, B2)
	Row         int    // starts from 0. -1 indicates a full column reference (e.g., A:A)
	Col         int    // starts from 0. -1 indicates a full row reference (e.g., 1:1)
	RowAbsolute bool   // true if the row is absolute (e.g., $1, $2)
	ColAbsolute bool   // true if the column is absolute (e.g., $A, $B)
}

func (c CellExpr) String() string {
	return fmt.Sprintf("CellExpr(%s)", c.Ident.Raw)
}

type ArrayExpr struct {
	baseNode
	BraceOpen  *Token   // The opening brace token {
	Elements   [][]Node // [row][column] of expressions
	BraceClose *Token   // The closing brace token }
}

func (a ArrayExpr) String() string {
	var sb strings.Builder
	sb.WriteString("ArrayExpr(")
	for i, row := range a.Elements {
		if i > 0 {
			sb.WriteString(", ")
		}
		sb.WriteString("[")
		for j, col := range row {
			if j > 0 {
				sb.WriteString(", ")
			}
			sb.WriteString(col.String())
		}
		sb.WriteString("]")
	}
	sb.WriteString(")")
	return sb.String()
}
