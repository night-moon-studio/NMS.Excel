# AzulX-NPOI

------

### 进一步封装NPOI的调用，Core2.0以上

#### 简单的文件创建以及调用：

```C#

using (ExcelFile file = new ExcelFile(filePath)){ ....; }
using (ExcelFile file = new ExcelFile(filePath,ExcelVersion.V2007)){ ....; }

```

#### 简单的Sheet操作：

```C#

//检测Sheet页是否存在

HasSheet("Page1");
HasSheet(100);


//当Sheet不存在时会自动创建

file.Select("Page1");
file.Select(100);

```

#### 简单的行列位置操作：

```C#
MoveTo(int row,int col);	 //将操作位置移动到指定位置

MoveToCol(int col);		 //移动到当前行，指定列

MoveToRow(int row);		 //移动到当前列，指定行

NextRow(bool isFirstCol = true); //移动到下一行，isFirstCol 是否将位置指向第一列
PrewRow(bool isFirstCol = true); //移动到上一行，isFirstCol 是否将位置指向第一列
```

#### 简单的赋值操作：

```C#
CurrentCell(value,style=null);		//给当前单元格赋值

NextCell(value,style=null);		//给下一个单元格赋值，并将位置移动到下一个单元格
SpeicalCell(index,value,style=null);	//给指定列的单元格赋值
```

#### 简单的属性操作(获取/赋值)

```C#
//当前行列单元格操作

StringValue = "test";
NumValue = 1.00;
DateValue = DateTime.Now;
BoolValue = false;


//当前行，下一列单元格操作

NextStringValue = "test";
NextNumValue = 1.00;
NextDateValue = DateTime.Now;
NextBoolValue = false;


//当前行，上一列单元格操作

PrewStringValue = "test";
PrewNumValue = 1.00;
PrewDateValue = DateTime.Now;
PrewBoolValue = false;
```



#### 提供简单的模板处理:

```Bash
@Name:Test1  
	@Split:|  
	@Sheet:Page1  
		@Header: @StartAt:0 head1|head2|head3|head4  
		@Content: @StartAt:0 property1|property2|property3|property4  
@End  
-----------------------------------or-----------------------------------
@Name:Test2  
@Split:@  
@Sheet:Page1  
@Header: @StartAt:3 名字@性别@年龄@备注@其他  
@Content: @StartAt:3 Name@Sex@Age@Description@other  
@End
```
#### 说明

- 每一个@End标识都是一个模板的结束，称为模块模板。  
- @Name作为模块模板缓存的Key.  
- @Split为单元格之间的分隔符.  
- @Sheet为当前操作的Sheet页名字  
- @Header为表格头部  
- @Content为表格的内容  
- @StartAt为起始行位置

#### 使用

```C#
using (ExcelFile file = new ExcelFile(filePath))
{
        ExcelStyle style = ExcelStyle.Create(file);
        file
              .LoadTemplate("1.txt") 		//加载模板文件
              .UseTemplate("Test2")  		//使用模板
              .FillHeader(style.Header())	//创建头部
              .FillCollection(students)		//填充集合
              .Save();				//保存
}
```


### 项目计划

   - 挖掘需求,支持更多的方便操作

### 更新日志

   - 2018-03-26：有想法，并完成初步封装.
   - 2018-03-27：增加和重构API，修复一些BUG.
        - 修复EOF读取失败的BUG.
        - 修复模板不能再第一列开始的BUG.
        - 增加属性赋值和获取操作.
        - 重构大量API，使其简洁，重用现有代码.
