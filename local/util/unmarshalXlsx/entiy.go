package unmarshalXlsx

// ig ignore 忽略
// rn reName 重命名
// fn funcName 方法名
// en enum 枚举

type Example struct {
	Id   uint64 `json:"id" ig:"1" `
	Name string `json:"name" rn:"姓名"`
	Age  int    `json:"age" rn:"年龄" fn:"addTen"`
	Sex  string `json:"sex" rn:"性别" en:"0|保密;1|男;2|女"`
}
