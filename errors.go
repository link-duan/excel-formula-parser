package excelformulaparser

import "fmt"

type LexError struct {
	Pos     Pos
	Message string
}

func (e *LexError) Error() string {
	return fmt.Sprintf("%s: %s", e.Pos, e.Message)
}

func newLexError(pos Pos, format string, args ...interface{}) *LexError {
	return &LexError{
		Pos:     pos,
		Message: fmt.Sprintf(format, args...),
	}
}

type ParseError struct {
	Pos     Pos
	Message string
}

func newParseError(pos Pos, format string, args ...interface{}) *ParseError {
	return &ParseError{
		Pos:     pos,
		Message: fmt.Sprintf(format, args...),
	}
}

func (e *ParseError) Error() string {
	return fmt.Sprintf("%s: %s", e.Pos, e.Message)
}
