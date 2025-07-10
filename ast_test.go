package excelformulaparser

import "testing"

func TestPos(t *testing.T) {
	var p = Pos{Line: 1, Column: 2}
	if p.String() != "(1,2)" {
		t.Errorf("Expected '(1,2)', got '%s'", p.String())
	}
	p.nextLine()
	if p.String() != "(2,1)" {
		t.Errorf("Expected '(2,1)', got '%s'", p.String())
	}
	p.nextColumn()
	if p.String() != "(2,2)" {
		t.Errorf("Expected '(2,2)', got '%s'", p.String())
	}
}

func TestIdent(t *testing.T) {
	var ident = IdentExpr{
		baseNode: newBaseNode(Pos{Line: 1, Column: 1}, Pos{Line: 1, Column: 4}),
		Name:     newToken(Pos{Line: 1, Column: 1}, Pos{Line: 1, Column: 4}, Ident, "NAME"),
	}
	if ident.String() != "IdentExpr(Name: NAME)" {
		t.Errorf("Expected 'IdentExpr(Name: NAME)', got '%s'", ident.String())
	}
}
