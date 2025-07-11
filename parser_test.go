package excelformulaparser

import "testing"

func TestParse(t *testing.T) {
	tests := []struct {
		src      string
		expected string
	}{
		{`=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(A1) + (@ERROR.TYPE(#VALUE!) = 2)`, `BinaryExpr(Left: BinaryExpr(Left: UnaryExpr(Operator: +, Operand: IdentExpr(Name: AName)), Operator: -, Right: ParenthesizedExpr(Inner: BinaryExpr(Left: UnaryExpr(Operator: -, Operand: UnaryExpr(Operator: +, Operand: UnaryExpr(Operator: -, Operand: UnaryExpr(Operator: +, Operand: LiteralExpr(Value: -2))))), Operator: ^, Right: LiteralExpr(Value: 6)))), Operator: =, Right: BinaryExpr(Left: BinaryExpr(Left: ArrayExpr([LiteralExpr(Value: "A"), LiteralExpr(Value: "B")]), Operator: +, Right: UnaryExpr(Operator: @, Operand: FunCallExpr(Name: SUM, Arguments: [CellExpr(A1)]))), Operator: +, Right: ParenthesizedExpr(Inner: BinaryExpr(Left: UnaryExpr(Operator: @, Operand: FunCallExpr(Name: ERROR.TYPE, Arguments: [LiteralExpr(Value: #VALUE!)])), Operator: =, Right: LiteralExpr(Value: 2)))))`},
		{"=@A:A", "UnaryExpr(Operator: @, Operand: RangeExpr(CellExpr(A):CellExpr(A)))"},
		{`=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"`, `BinaryExpr(Left: FunCallExpr(Name: IF, Arguments: [BinaryExpr(Left: LiteralExpr(Value: "a"), Operator: =, Right: ArrayExpr([LiteralExpr(Value: "a"), LiteralExpr(Value: "b")], [LiteralExpr(Value: "c"), LiteralExpr(Value: #N/A)], [LiteralExpr(Value: -1), LiteralExpr(Value: TRUE)])), LiteralExpr(Value: "yes"), LiteralExpr(Value: "no")]), Operator: &, Right: LiteralExpr(Value: "  more ""test"" text"))`},
		{"={#N/A}", "ArrayExpr([LiteralExpr(Value: #N/A)])"},
		{"={1,2;3,4}", "ArrayExpr([LiteralExpr(Value: 1), LiteralExpr(Value: 2)], [LiteralExpr(Value: 3), LiteralExpr(Value: 4)])"},
		{"=-{-1}", "UnaryExpr(Operator: -, Operand: ArrayExpr([LiteralExpr(Value: -1)]))"},
		{"=--1", "UnaryExpr(Operator: -, Operand: LiteralExpr(Value: -1))"},
		{"=-A1", "UnaryExpr(Operator: -, Operand: CellExpr(A1))"},
		{"=\"abcd\"", "LiteralExpr(Value: \"abcd\")"},
		{"=10%", "UnaryExpr(Operator: %, Operand: LiteralExpr(Value: 10))"},
		{"=A1^B1", "BinaryExpr(Left: CellExpr(A1), Operator: ^, Right: CellExpr(B1))"},
		{"=A1&B1", "BinaryExpr(Left: CellExpr(A1), Operator: &, Right: CellExpr(B1))"},
		{"=1<>2", "BinaryExpr(Left: LiteralExpr(Value: 1), Operator: <>, Right: LiteralExpr(Value: 2))"},
		{"=$A:$A", "RangeExpr(CellExpr($A):CellExpr($A))"},
		{"=$1:1", "RangeExpr(CellExpr($1):CellExpr(1))"},
		{"=A:B", "RangeExpr(CellExpr(A):CellExpr(B))"},
		{"=1:1", "RangeExpr(CellExpr(1):CellExpr(1))"},
		{"=A1:B2", "RangeExpr(CellExpr(A1):CellExpr(B2))"},
		{"=A1:B2:C3", "RangeExpr(CellExpr(A1):CellExpr(B2):CellExpr(C3))"},
		{"=A1:B$2:$C$3", "RangeExpr(CellExpr(A1):CellExpr(B$2):CellExpr($C$3))"},
		{"=$A$1:C3", "RangeExpr(CellExpr($A$1):CellExpr(C3))"},
		{"=123", "LiteralExpr(Value: 123)"},
		{"=123.456", "LiteralExpr(Value: 123.456)"},
		{"=TRUE", "LiteralExpr(Value: TRUE)"},
		{"=SUM()", "FunCallExpr(Name: SUM, Arguments: [])"},
		{"=SUM(1,2)", "FunCallExpr(Name: SUM, Arguments: [LiteralExpr(Value: 1), LiteralExpr(Value: 2)])"},
		{"=1 + 2 - 3", "BinaryExpr(Left: BinaryExpr(Left: LiteralExpr(Value: 1), Operator: +, Right: LiteralExpr(Value: 2)), Operator: -, Right: LiteralExpr(Value: 3))"},
		{"=1+(2-3)", "BinaryExpr(Left: LiteralExpr(Value: 1), Operator: +, Right: ParenthesizedExpr(Inner: BinaryExpr(Left: LiteralExpr(Value: 2), Operator: -, Right: LiteralExpr(Value: 3))))"},
		{"=1+2*3", "BinaryExpr(Left: LiteralExpr(Value: 1), Operator: +, Right: BinaryExpr(Left: LiteralExpr(Value: 2), Operator: *, Right: LiteralExpr(Value: 3)))"},
	}
	for _, test := range tests {
		p := NewParser(test.src)
		node, err := p.Parse()
		if err != nil {
			t.Errorf("Parse error for '%s': %v", test.src, err)
			continue
		}
		if node.String() != test.expected {
			t.Errorf("For input '%s', expected '%s', got '%s'", test.src, test.expected, node.String())
		}
	}
}

func Test_colNameToIndex(t *testing.T) {
	tests := []struct {
		colName  string
		expected int
	}{
		{"A", 0},
		{"B", 1},
		{"Z", 25},
		{"AA", 26},
		{"AB", 27},
		{"AZ", 51},
		{"BA", 52},
	}
	for _, test := range tests {
		result := colNameToIndex(test.colName)
		if result != test.expected {
			t.Errorf("colNameToIndex(%s) = %d; want %d", test.colName, result, test.expected)
		}
	}
}
