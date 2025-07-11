<h1 align="center">excel-formula-parser</h1>
<p align="center"><i>A parser for excel-formula written in pure Go</i></p>
<p align="center">
  <img src="https://goreportcard.com/badge/github.com/link-duan/excel-formula-parser"/>
  <img src="https://github.com/link-duan/excel-formula-parser/actions/workflows/go.yml/badge.svg"/>
  <a href="https://codecov.io/gh/link-duan/excel-formula-parser" > 
    <img src="https://codecov.io/gh/link-duan/excel-formula-parser/graph/badge.svg?token=MN0H60O7LB"/> 
  </a>
</p>

## Unsupportted syntax

- `R1C1` style cell reference syntax. (had been deprecated by excel)
- intersection operator (whitespace between references)

## Usage

```go
ast, _ := excelformulaparser.NewParser("=SUM(A1:B2, C$3, 4:4)").Parse()
fmt.Printf("%v", ast)
```
