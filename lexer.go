package excelformulaparser

import (
	"io"
	"unicode/utf8"
)

var keywords = map[string]TokenType{
	"TRUE":  BoolLiteral,
	"FALSE": BoolLiteral,
}

type tokenState struct {
	tokenType     TokenType
	tokenStartPos Pos
	tokenEndPos   Pos
	tokenStart    int // start offset
	tokenEnd      int // end offset
}

type lexer struct {
	src    []rune
	pos    Pos
	offset int
	ch     rune
	tokenState
}

func newLexer(src string) *lexer {
	var l = &lexer{
		src: []rune(src),
	}
	l.init()
	return l
}

func (l *lexer) next() (*Token, error) {
	if len(l.src) == 0 || l.ch == -1 {
		return nil, io.EOF // EOF
	}
	for l.ch == ' ' || l.ch == '\n' || l.ch == '\t' {
		l.nextch()
	}
	var start = l.pos
	switch l.ch {
	case '"':
		return l.stringLiteral(l.ch)
	case '\'':
		return l.stringLiteral(l.ch)
	case '=':
		l.nextch()
		return newToken(start, start, Equal, "="), nil
	case '!':
		l.nextch()
		return newToken(start, start, Exclamation, "!"), nil
	case ',':
		l.nextch()
		return newToken(start, start, Comma, ","), nil
	case '(':
		l.nextch()
		return newToken(start, start, ParenOpen, "("), nil
	case ')':
		l.nextch()
		return newToken(start, start, ParenClose, ")"), nil
	case '{':
		l.nextch()
		return newToken(start, start, BraceOpen, "{"), nil
	case '}':
		l.nextch()
		return newToken(start, start, BraceClose, "}"), nil
	case '[':
		l.nextch()
		return newToken(start, start, BracketOpen, "["), nil
	case ']':
		l.nextch()
		return newToken(start, start, BracketClose, "]"), nil
	case ':':
		l.nextch()
		return newToken(start, start, Colon, ":"), nil
	case ';':
		l.nextch()
		return newToken(start, start, Semicolon, ";"), nil
	case '+':
		l.nextch()
		return newToken(start, start, Plus, "+"), nil
	case '-':
		l.nextch()
		return newToken(start, start, Minus, "-"), nil
	case '*':
		l.nextch()
		return newToken(start, start, Multiply, "*"), nil
	case '/':
		l.nextch()
		return newToken(start, start, Divide, "/"), nil
	case '$':
		return l.absoluteReference()
	case '&':
		l.nextch()
		return newToken(start, start, Concat, "&"), nil
	case '^':
		l.nextch()
		return newToken(start, start, Exponentiation, "^"), nil
	case '%':
		l.nextch()
		return newToken(start, start, Percent, "%"), nil
	case '<':
		var start = l.pos
		l.nextch()
		switch l.ch {
		case '>':
			l.nextch()
			return newToken(start, l.pos, NotEqual, "<>"), nil
		case '=':
			l.nextch()
			return newToken(start, l.pos, LessThanOrEqual, "<="), nil
		default:
			return newToken(start, l.pos, LessThan, "<"), nil
		}
	case '>':
		var start = l.pos
		l.nextch()
		switch l.ch {
		case '=':
			l.nextch()
			return newToken(start, l.pos, GreaterThanOrEqual, ">="), nil
		default:
			return newToken(start, l.pos, GreaterThan, ">"), nil
		}
	case '#':
		return l.errValue()
	case '@':
		l.nextch()
		return newToken(start, start, ImplicitIntersection, "@"), nil
	default:
		if isDigit(l.ch) {
			return l.number()
		}
		return l.ident()
	}
}

func (l *lexer) init() {
	if len(l.src) == 0 {
		l.pos = Pos{Line: 1, Column: 1}
		l.ch = -1 // EOF
		return
	}
	l.pos = Pos{Line: 1, Column: 1}
	l.ch = l.src[0]
	l.offset = 1
}

func (l *lexer) nextch() {
	if l.offset < len(l.src) {
		l.ch = l.src[l.offset]
		if l.ch == '\n' {
			l.pos.nextLine()
		} else {
			l.pos.nextColumn()
		}
		// Move to the next character
		l.offset++
	} else {
		l.ch = -1
	}
}

func (l *lexer) eat(n int) {
	for i := 0; i < n; i++ {
		l.nextch()
	}
}

func (l *lexer) absoluteReference() (*Token, error) {
	var start = l.pos
	var startOffset = l.offset - 1 // Start at the current character
	l.nextch()

	var consumeRow = func() (len int, absolute, ok bool) {
		var startOffset = l.offset - 1 // Start at the current character
		var endOffset = l.offset - 1
		if l.ch == '$' {
			if !isDigit(l.peekChar()) {
				return 0, false, false
			}
			absolute = true
			l.nextch() // Consume the '$'
		}
		for isDigit(l.ch) {
			endOffset = l.offset
			l.nextch()
		}
		if endOffset == startOffset {
			return 0, false, false // No digits found
		}
		ok = true
		len = endOffset - startOffset
		return
	}

	if isDigit(l.ch) {
		var endOffset = l.offset
		var end = l.pos
		for isDigit(l.ch) {
			end = l.pos
			endOffset = l.offset
			l.nextch()
		}
		return newToken(start, end, AbsoluteRow, string(l.src[startOffset:endOffset])), nil
	} else if isASCIILetter(l.ch) {
		var endOffset = l.offset
		var end = l.pos
		for isASCIILetter(l.ch) {
			end = l.pos
			endOffset = l.offset
			l.nextch()
		}
		digitLen, _, ok := consumeRow()
		if ok {
			end.Column += digitLen
			endOffset += digitLen
			return newToken(start, end, Cell, string(l.src[startOffset:endOffset])), nil
		}
		return newToken(start, end, AbsoluteColumn, string(l.src[startOffset:endOffset])), nil
	}
	return nil, newLexError(l.pos, "invalid absolute reference after '$'")
}

func (l *lexer) errValue() (*Token, error) {
	var values = []string{
		"#NULL!",
		"#DIV/0!",
		"#VALUE!",
		"#REF!",
		"#NAME?",
		"#NUM!",
		"#N/A",
	}
	for _, value := range values {
		_len := utf8.RuneCountInString(value)
		// check length
		if l.offset-1+_len > len(l.src) {
			continue // Not enough characters left for this error value
		}
		if string(l.src[l.offset-1:l.offset-1+_len]) == value {
			var start = l.pos
			l.eat(_len) // Consume the error value
			var end = l.pos
			return newToken(start, end, EValue, value), nil
		}
	}
	return nil, newLexError(l.pos, "unrecognized error value")
}

func (l *lexer) stringLiteral(quote rune) (*Token, error) {
	var start = l.pos
	var end = start
	l.nextch() // Consume the opening quote
	for l.ch >= 0 {
		if l.ch == '\n' {
			return nil, newLexError(l.pos, "newline in string literal")
		}
		if l.ch == quote {
			if l.peekChar() == quote {
				l.nextch()
			} else {
				break
			}
		}
		end = l.pos
		l.nextch()
	}
	if l.ch != quote {
		return nil, newLexError(l.pos, "unclosed string literal")
	}
	end = l.pos // Update end to the position after the closing quote
	l.nextch()  // Consume the closing quote
	return newToken(start, end, String, string(l.src[start.Column-1:end.Column])), nil
}

func (l *lexer) number() (*Token, error) {
	var startOffset = l.offset - 1 // Start at the current character
	var endOffset = l.offset - 1
	var start = l.pos
	var end = l.pos
	if !isDigit(l.ch) {
		return nil, newLexError(l.pos, "invalid number start")
	}
LOOP:
	for l.ch >= 0 {
		switch {
		case isDigit(l.ch):
			end = l.pos
			endOffset = l.offset
			l.nextch()
		case l.ch == '.':
			l.nextch()
			if !isDigit(l.ch) {
				return nil, newLexError(l.pos, "invalid number format")
			}
			// Continue to consume digits after the decimal point
			for isDigit(l.ch) {
				end = l.pos
				endOffset = l.offset
				l.nextch()
			}
		default:
			break LOOP
		}
	}
	return newToken(start, end, Number, string(l.src[startOffset:endOffset])), nil
}

func (l *lexer) ident() (*Token, error) {
	var startOffset = l.offset - 1 // Start at the current character
	var start = l.pos
	var end = l.pos
	if !isASCIILetter(l.ch) && l.ch != '_' && l.ch <= utf8.RuneSelf {
		return nil, newLexError(l.pos, "invalid identifier start")
	}
	var endOffset = l.offset
	l.nextch()
LOOP:
	for l.ch >= 0 {
		switch {
		case isASCIILetter(l.ch) || isDigit(l.ch) || l.ch == '.':
			end = l.pos
			endOffset = l.offset
			l.nextch()
		case l.ch == '_':
			end = l.pos
			endOffset = l.offset
			l.nextch()
		case l.ch > utf8.RuneSelf:
			end = l.pos
			endOffset = l.offset
			l.nextch()
		default:
			break LOOP
		}
	}
	var rawIdent = string(l.src[startOffset:endOffset])
	if t, ok := keywords[rawIdent]; ok {
		return newToken(start, end, t, rawIdent), nil
	}
	if isCellReference([]rune(rawIdent)) {
		return newToken(start, end, Cell, rawIdent), nil
	}
	if l.ch == '$' {
		l.nextch() // Consume the '$'
		digitsLen := l.eatDigits()
		if digitsLen == 0 {
			return nil, newLexError(l.pos, "expected digits after identifier '$'")
		}
		end.Column += digitsLen + 1 // Include the '$' in the end position
		endOffset += digitsLen + 1  // Include the '$' in the end offset
		return newToken(start, end, Cell, string(l.src[startOffset:endOffset])), nil
	}
	return newToken(start, end, Ident, rawIdent), nil
}

func (l *lexer) eatDigits() (len int) {
	var startOffset = l.offset - 1 // Start at the current character
	var endOffset = l.offset - 1
	for isDigit(l.ch) {
		endOffset = l.offset
		l.nextch()
	}
	len = endOffset - startOffset
	return
}

// Peek returns the next token without consuming it.
func (l *lexer) peekChar() rune {
	if l.offset < len(l.src) {
		return l.src[l.offset]
	}
	return -1 // EOF
}

func isASCIILetter(ch rune) bool {
	return ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z'
}

func isDigit(ch rune) bool {
	return ch >= '0' && ch <= '9'
}

func isCellReference(raw []rune) bool {
	if len(raw) < 2 {
		return false
	}
	var letterCount = 0
	// Count letters at the start
	for _, ch := range raw {
		if isASCIILetter(ch) {
			letterCount++
		} else {
			break
		}
	}
	if letterCount == 0 || letterCount >= len(raw) {
		return false // Must have at least one letter and at least one digit
	}
	// Check if the rest are digits (0-9)
	for _, ch := range raw[letterCount:] {
		if !isDigit(ch) {
			return false
		}
	}
	return true
}
