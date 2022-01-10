package unmarshalXls

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/shakinm/xlsReader/xls"
	"reflect"
	"strconv"
	"strings"
)

type People struct {
	Id   uint64 `json:"Id" ig:"1" `
	Name string `json:"Name" rn:"姓名"`
	Age  int    `json:"age" rn:"年龄" fn:"addTen"`
	Sex  string `json:"sex" rn:"性别" en:"0|保密;1|男;2|女"`
}

type ReadXls struct {
	FilePath   string
	Bytes      []byte
	SheetIndex int
	SheetName  string
	RowStart   int
	RowEnd     int
	Obj        interface{}
	Funcs      map[string]interface{}
}

func (c *ReadXls) Init() {
	if c.RowStart == 0 {
		c.RowStart = 1
	} else {
		c.RowStart -= 1
	}
	if c.RowEnd == 0 {
		c.RowEnd = 5000
	} else {
		c.RowEnd -= 1
	}
}

func (c *ReadXls) ReadXlsFunc() error {
	// 解析结构
	sm := unmarshalStructType(c.Obj)
	// 解析xls文件 解析失败
	var workbook xls.Workbook
	var err error
	if c.FilePath != "" {
		workbook, err = xls.OpenFile(c.FilePath)
		if err != nil {
			return err
		}
	} else {
		workbook, err = xls.OpenReader(bytes.NewReader(c.Bytes))
		if err != nil {
			return err
		}
	}

	// 获取相应sheet 下标
	sheet, err := workbook.GetSheet(c.SheetIndex)
	if err != nil {
		return err
	}
	if c.SheetName != "" {
		// 对应sheetName 校验
		if sheet.GetName() != c.SheetName {
			return errors.New("check sheet fail! ")
		}
	}

	var ms []map[string]interface{}
	for i := 0; i <= sheet.GetNumberRows(); i++ {
		if row, err := sheet.GetRow(i); err == nil {
			if c.RowStart <= i && i <= c.RowEnd {
				m := map[string]interface{}{}
				for j, cell := range row.GetCols() {
					formatIndex := workbook.GetXFbyIndex(cell.GetXFIndex())
					format := workbook.GetFormatByIndex(formatIndex.GetFormatIndex())
					mv := format.GetFormatString(cell)
					if sm[j].Type != nil {
						// 获取方法值
						if sm[j].FnV != "" {
							str, err := c.Call(sm[j].FnV, mv)
							if err == nil {
								mv = str
							}
						}
						// 获取枚举值
						if len(sm[j].EnV) > 0 {
							mv = sm[j].EnV[mv]
						}
						// 处理数据类型
						ty := sm[j].Type.Kind()
						switch ty {
						case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
							n, _ := strconv.ParseInt(mv, 10, 64)
							m[sm[j].JsV] = n
						case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
							n, _ := strconv.ParseUint(mv, 10, 64)
							m[sm[j].JsV] = n
						case reflect.Float32, reflect.Float64:
							n, _ := strconv.ParseFloat(mv, sm[j].Type.Bits())
							m[sm[j].JsV] = n
						case reflect.String:
							m[sm[j].JsV] = mv
						case reflect.Bool:
							m[sm[j].JsV], _ = strconv.ParseBool(mv)
						default:
							m[sm[j].JsV] = mv
						}
					}
				}
				ms = append(ms, m)
			}
		}
	}
	msb, _ := json.Marshal(ms)
	err = json.Unmarshal(msb, &c.Obj)
	if err != nil {
		return err
	}
	return nil
}

func (c *ReadXls) Call(funcName string, param string) (string, error) {
	var val string
	if c.Funcs[funcName] == nil {
		return "", errors.New("func not exist")
	}
	fn := reflect.ValueOf(c.Funcs[funcName])
	result := fn.Call([]reflect.Value{reflect.ValueOf(param)})
	if len(result) == 1 {
		val = result[0].String()
	}
	return val, nil
}

type StructType struct {
	Index int               // 下标
	JsV   string            // json 值
	RnV   string            // reName 值
	FnV   string            // func 值
	EnV   map[string]string // enum map 0|保密;1|男;2|女
	Type  reflect.Type      // 类型
}

func unmarshalStructType(obj interface{}) map[int]StructType {
	//结构体指针
	sm := make(map[int]StructType, 0)
	el := reflect.TypeOf(obj).Elem().Elem()
	index := 0
	for i := 0; i < el.NumField(); i++ {
		tg := el.Field(i).Tag
		igv := tg.Get("ig")
		if igv != "1" {
			enm := make(map[string]string)
			env := tg.Get("en")
			if strings.Contains(env, "|") {
				if strings.Contains(env, ";") {
					ms := strings.Split(env, ";")
					for _, k := range ms {
						es := strings.Split(k, "|")
						if len(es) == 2 {
							enm[es[0]] = es[1]
						}
					}
				} else {
					es := strings.Split(env, "|")
					if len(es) == 2 {
						enm[es[0]] = es[1]
					}
				}
			}
			sm[index] = StructType{
				Index: i,
				JsV:   tg.Get("json"),
				RnV:   tg.Get("rn"),
				FnV:   tg.Get("fn"),
				EnV:   enm,
				Type:  el.Field(i).Type,
			}
			index += 1
		}
	}
	return sm
}
