<h1 align="center">excel-formula-parser</h1>
<p align="center"><i>A parser for excel-formula written in pure Go</i></p>
<p align="center">
  <img src="https://goreportcard.com/badge/github.com/link-duan/excel-formula-parser"/>
  <img src="https://github.com/link-duan/excel-formula-parser/actions/workflows/go.yml/badge.svg"/>
  <img src="https://raw.githubusercontent.com/link-duan/excel-formula-parser/badges/.badges/main/coverage.svg"/>
</p>

## Supportted syntax
- [x] Math operations (eg. `+ - * / ^ %`)
- [x] Cell references ( Absolute & Relative. eg. `$A$1` `A1:B2` )
- [x] Function call

## Usage
```go
ast, _ := excelformulaparser.NewParser("=SUM(A1:B2, C$3, 4:4)").Parse()
fmt.Printf("%v", ast)
```
