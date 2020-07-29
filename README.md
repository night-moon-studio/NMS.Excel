# NMS.Excel

## 使用方法

## 1、初始化 Natasha

```C#
NatashaComponentRegister.RegistDomain<NatashaAssemblyDomain>();
NatashaComponentRegister.RegistCompiler<NatashaCSharpCompiler>();
NatashaComponentRegister.RegistSyntax<NatashaCSharpSyntax>();
```

## 2、配置读写映射

```C#
var dict = new Dictionary<string, string> {

       { "名字","Name"},
       { "年龄","Age"},
       { "性别","Sex"},
       { "描述","Description"}
       
};
//忽略 “描述” 字段
ExcelOperator.SetWritterMapping<Student>(dict, "描述");
ExcelOperator.SetReaderMapping<Student>(dict);
```

## 3、写入与读取

```C#
//写
ExcelOperator.WriteToFile("1.xlsx", students);
//读
var list = ExcelOperator.FileToEntities<Student>("1.xlsx");
```

