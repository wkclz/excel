# Excel 工具类使用文档

### 介绍
***
> * 对 excel 的生成 的读取，一直都是一个很花时间的问题，现对excel的生成和读取，封装成一个工具，将加快excel 在 java中的处理。 
> * Excel 是对 poi 的封装，建立一个通用的excel模板，再通过 set,add,create等简单的操作方法生成 excel。同时也能指定起始度坐标，对标准的 Excel进行读取。


### 依赖：
***
> 1. JDK 1.8
> 1. poi-3.17.jar
> 1. poi-ooxml-3.17.jar
> 1. poi-ooxml-schemas-3.17.jar
> 1. xmlbeans-2.6.0.jar
> 1. commons-collections4-4.1.jar
* *以上依赖已经在 pom.xml 中描述*
 

### 特点：
***
> 对象化操作


### Excel生成类的使用方法：
***
1. 通过 new Excel() 创建excel对象
1. 通过set 方法设置文档基本信息。
1. 提供以下set方法：
> * setTitle(String title) 
> * setCreate_by(String create_by)
> * setCreate_by(Object create_by)【会将object转换为String】
> * setDateFrom(String dateFrom) 【null】
> * setDateTo(String dateTo) 【null】
> * setSavePath(String savePath)
> * setHeader(List<String> header)
> * setHeader(String[] header)
> * setWidth(Integer width)【在有设置 header的时候，会自动取其长度，不需要header的需要设置 width】

* 提供的get方法：
> * 以上的set方法均有对应的无参get方法。

* add方法：
> * addRow(ExcelRow row) 添加行对象【protected】【保护的方法。createXlsx过程会自动添加】
> * addRowFromCache(ExcelRow row) 【将缓存对象加入到excel。会将上一个createRow得到的对象立即添加到excel】

* 提供的create方法：
> * createRow() 创建行对象【创建row对象，创建前会将上一个行对象添加到excel】
> * createRowInCache()创建行对象【不直接添加到excel中，需要调用addRowFromCache(ExcelRow row) 对会被加入到excel】
> * CreateXlsx() 创建最终的excel.

* 通过 createRow方法创建了行对象row后，可以对行对象进行操作。

###### 提供的单元格操作方法：
> addCell(Object cellContent, boolean border, short align, int col, int row)
> cellContent：cell 文本内容。必需，将自动设置 cell 宽度。超过一定长度后自动定义为富文本框 \
> border：cell的边框，默认为true。边框大小不提供自定义 \
> align：对齐方式，需要从 ExcelUtil里面取static值。仅提货左，中两种常用取值 \
> col：合并行，需要与row结合使用。 \
> row：合并列，需要与col结合使用。 \
> col和row需成对出现或不设置。默认为1。当有行合并时，将不自动设置行宽。 \

###### cell 类型支持：
> java.lang.Integer.
> java.lang.String【默认类型】
> java.util.Date【年月日时分秒】
> java.sql.Date【年月日】
> java.lang.Double【默认保留两位小数】
> 按上述类型传入的值，将会自动设置到cell中，其他类型的值通不过上述类型判断，将用默认类型处理。

###### createXlsx的写入
> 了解了excel的写入操作，将更便于此工具的使用。
> * 对 excel对象得到的参数进行验证，如果有错误，将返回错误结果。
> * 使用 header 的长度，或者 width 的值，进行标题的合并和设置。
> * 写入文档创建时间。如果创建人有设置，加入创建人。
> * 如果时间范围存在，设置时间范围。
> * 如果存在  header，写入excel的列名
> * 依次对行和进行循环。
> * 如果有行列合并，对所在区域进行合并，并设置空字符串值。如果已经有值，说明已经被合并，找到下一个空值的cell
> * 对cell的内容进行类型判断。不同的类型，将给不同的样式。
> * 将内容写入到cell
> * 内容完全写入到cell后，将excel通过文件流的方式写入到磁盘。完成 excel生成

##### 示例代码：
***
> com.wkclz.util.excel.ExcelTest.main();


### Excel读取类的使用方法：
***
> * setStartRow(int rowNum) 
> * setStartCol(int colNum)
> * setTypes(String[] types)
> * 最后骑过 analysisXlsx() 方法，得到所有行数据。


##### 示例代码：
***
> com.wkclz.util.excel.ExcelRdTest.main();


***
> 最后，此工具类共享出来给大家使用，希望大家能够帮助一起完善，通过开源的方式互助。发现有什么bug，或者有什么想法欢迎 PR.

# 捐赠
***
> 如果您感觉这个工具很好用，希望支持作者，您还可以进行捐赠： \
> **支付宝** 
<div align=center>
<img src="http://static.wkclz.com/img/alipay.png" width="280" height="280" alt="亦菲表演机器猫"/>
</div>


***
# 更新日志
2017-11-15 22:58:33
***
1. Excel读取工具 ExcelRd 定义枚举类型ExcelRdTypeEnum：INTEGER("整形"),DOUBLE("双精浮点型"),DATE("日期型"),DATETIME("日期时间型"),STRING("字符型");读取时需要初始化字段类型。

2017-12-09 10:15:52
***
1. Excel 生成去除结果返回，换成 thros 的方式抛出错误。
2. 将 excel 生成过程独立，即将加入文件流返回
3. 结构优化
