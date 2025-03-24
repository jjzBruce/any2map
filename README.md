# any2map

`any2map `是一个将可以将数据转为Map的小工具。

目前，此系统包含有以下功能：

- Excel 转 Map
- 功能2
- 功能3
- 

## 准备

```xml
<dependency>
    <groupId>com.modern.tools</groupId>
    <artifactId>any2map</artifactId>
    <version>1.0.0</version>
</dependency>
```

## 快速开始

### Excel

使用将`Excel`转为`Map`的时候，使用文件地址作为数据源。

#### 一般将Excel转Map

```java
String filePath = “path/to/your/file.xlsx”;
// 创建转换配置
ExcelConvertConfig config = new ExcelConvertConfig(filePath);
// 创建sheet配置
SheetDataConfig sheetDataConfig = new SheetDataConfig();
// 创建日期时间数据配置，坐标(0, 5)表示是日期时间数据
ExcelDateTypeConfig edtc = new ExcelDateTypeConfig(0, 5);
sheetDataConfig.addExcelDateTypeConfig(edtc);
config.addSheetDataConfig(sheetDataConfig);
// 根据配置创建转换器
MapConverter x2ms = Any2Map.createMapConverter(config2);
// 输出结果
Map<String, Object> map = x2ms.toMap();
```

如下Excel：

<img src="./README.assets/image-20250324165642622.png" alt="image-20250324165642622" style="zoom:67%;" />

输出结果：

```json
{
  "S1": [
    {
      "A": "跨列",
      "B": "跨列",
      "C": "跨行",
      "D": "-",
      "E": "跨行跨列",
      "2000-01-11": "跨行跨列"
    },
    {
      "A": 12.0,
      "B": 1300.0,
      "C": "跨行",
      "D": -1288.0,
      "E": "跨行跨列",
      "2000-01-11": "跨行跨列"
    },
    {
      "A": true,
      "B": false,
      "C": 1300.0,
      "D": 1300.0,
      "E": 1300.0,
      "2000-01-11": 1300.0
    }
  ]
}
```



#### 将多级Head的Excel转Map



#### Excel转Map再进行分组









### To Do List

- xlsx 和 xls 的区别并需要测试
- Excel 多级Head实现
- Excel 分组实现
- Mongo 数据转Map实现
- Jdbc 数据转Map实现



