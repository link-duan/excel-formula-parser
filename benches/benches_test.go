package benches

import (
	"testing"

	excelformulaparser "github.com/link-duan/excel-formula-parser"
	"github.com/xuri/efp"
)

const benchFormula = `=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(A1) + (@ERROR.TYPE(#VALUE!) = 2)`

func Benchmark_Parse(b *testing.B) {
	for i := 0; i < b.N; i++ {
		excelformulaparser.NewParser(benchFormula).Parse()
	}
}

func Benchmark_efp(b *testing.B) {
	for i := 0; i < b.N; i++ {
		p := efp.ExcelParser()
		p.Parse(benchFormula)
	}
}
