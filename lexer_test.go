package excelformulaparser

import (
	"testing"
	"unicode/utf8"
)

func TestLexer(t *testing.T) {
	var tests = []struct {
		src      string
		expected TokenType
	}{
		{"A$123", Cell},
		{"$A$1", Cell},
		{"A1", Cell},
		{"AA12", Cell},
		{"$ABC", AbsoluteColumn},
		{"$123", AbsoluteRow},
		{"TRUE", BoolLiteral},
		{"FALSE", BoolLiteral},
		{"123", Number},
		{"1.23", Number},
		{"0.23", Number},
		{"0.230000", Number},
		{`"hello"`, String},
		{`'world'`, String},
		{`'with quote\''`, String},
		{`"with quote\""`, String},
		{"=", Equal},
		{"!", Exclamation},
		{",", Comma},
		{"(", ParenOpen},
		{")", ParenClose},
		{"{", BraceOpen},
		{"}", BraceClose},
		{"[", BracketOpen},
		{"]", BracketClose},
		{":", Colon},
		{";", Semicolon},
		{"+", Plus},
		{"-", Minus},
		{"*", Multiply},
		{"/", Divide},
		{"a", Ident},
		{"_", Ident},
		{"abc", Ident},
		{"ABC", Ident},
		{"工作表1", Ident},
		{">", GreaterThan},
		{">=", GreaterThanOrEqual},
		{"<", LessThan},
		{"<=", LessThanOrEqual},
		{"<>", NotEqual},
	}
	for _, test := range tests {
		l := newLexer(test.src)
		token, err := l.next()
		if err != nil {
			t.Errorf("Expected token for input '%s', got error: %v", test.src, err)
			continue
		}
		if token.Type != test.expected {
			t.Errorf("For input '%s', expected token type %v, got %v", test.src, test.expected, token.Type)
		}
		if token.Raw != test.src {
			t.Errorf("For input '%s', expected raw token '%s', got '%s'", test.src, test.src, token.Raw)
		}
		// check start and end positions
		if token.Start.Line != 1 || token.Start.Column != 1 {
			t.Errorf("For input '%s', expected start position (1, 1), got (%d, %d)", test.src, token.Start.Line, token.Start.Column)
		}
		if token.End.Line != 1 || token.End.Column != utf8.RuneCountInString(test.src) {
			t.Errorf("For input '%s', expected end position (1, %d), got (%d, %d)", test.src, utf8.RuneCountInString(test.src), token.End.Line, token.End.Column)
		}
	}
}
