package unmarshalXlsx

import (
	"fmt"
	"strconv"
	"testing"
)

func TestReadXls_ReadXlsFunc(t *testing.T) {
	var ps []People
	c := ReadXls{
		FilePath:   "./123.xls",
		Bytes:      nil,
		SheetIndex: 0,
		SheetName:  "",
		RowStart:   2,
		RowEnd:     8,
		Obj:        &ps,
		Funcs: map[string]interface{}{
			"addTen": func(s string) string {
				i, _ := strconv.Atoi(s)
				return fmt.Sprintf("%d", i+10)
			},
		},
	}
	c.Init()
	err := c.ReadXlsFunc()
	if err != nil {
		fmt.Println("========= err =========")
		fmt.Println(err)
	}
	fmt.Println(ps)
}
