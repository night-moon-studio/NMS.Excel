# AzulX-NPOI
对NPOI的封装，Core2.0以上

进一步封装NPOI的调用，提供简单的模板处理:
```C#
@Name:Test1  
@Split:|  
@Sheet:Page1  
@Header: @StartAt:0 head1|head2|head3|head4  
@Content: @StartAt:0 property1|property2|property3|property4  
@End  

@Name:Test2  
@Split:@  
@Sheet:Page1  
@Header: @StartAt:3 名字@性别@年龄@备注@其他  
@Content: @StartAt:3 Name@Sex@Age@Description@other  
@End
```


每一个@End标识都是一个模板的结束，称为模块模板。  

@Name作为模块模板缓存的Key.  

@Split为单元格之间的分隔符.  

@Sheet为当前操作的Sheet页名字  

@Header为表格头部  

@Content为表格的内容  




