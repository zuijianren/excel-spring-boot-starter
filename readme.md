# excel 工具类

> 当前仅支持导出功能

## 特性

* 支持 一对多 导出
* 支持使用注解配置导出

## 快速开始

1. 导入依赖

    ```xml
    <dependency>
        <groupId>com.zuijianren</groupId>
        <artifactId>excel-spring-boot-starter</artifactId>
        <version>1.0-SNAPSHOT</version>
    </dependency>
    ```

2. 导出对象添加对应注解
    * 示例:

       ```java
       @ExcelSheet("学生表")
       @Data
       public class Student {
       @ExcelProperty(value = {"id"}, order = 0)
       private Integer id;
       @ExcelProperty(value = {"学生", "年龄"}, order = 2)
       private Integer age;
       @ExcelProperty(value = {"学生", "名字"}, order = 1)
       private String name;
 
       public Student(Integer id, Integer age, String name) {
             this.id = id;
             this.age = age;
             this.name = name;
         }
       }
 
       ```

3. 导出

   ```java
   ExcelWriter.createExcelWriter("a.xlsx")
                .write(Student.class, Collections.singletonList(student)) // 可以一次写入多个sheet数据
                .write(Teacher.class, Arrays.asList(teacher, teacher))
                .doWrite();
   ```

## 使用说明

### @ExcelSheet

> 标识当前对象是一个 sheet 对象

|  属性   | 作用  |
|  ----  | ----  |
| value  | 首行是否展示表名 |
| showSheetName  | 单元格 |
| showSerialNumber  | 是否展示序列号 |
| freezeHead  | 是否冻结首行及表头(新版excel可能会报错有问题) |

### @ExcelProperty

> 标识当前属性为需要导出的属性

|  属性   | 作用  |
|  ----  | ----  |
| value  | 表头 (多级) |
| order  | 表头的顺序 (从小到大) |
| converter  | 转换器 (用于处理类型转换)|
| nested  | 内嵌 |
| showCurrentName  | 是否展示当前的名字 |

### @ExcemMultiProperty

> 标识当前属性为需要导出的集合属性(渲染时将会按照一对多的方式进行渲染)
>
> 如果集合属性仅会使用一个字符串写入(例如拼接的方式), 则应使用 ExcelProperty 注解 需要实现 ExcelConverter 接口 并进行声明 才能成功写入

|  属性   | 作用  |
|  ----  | ----  |
| value  | 表头 (多级) |
| order  | 表头的顺序 (从小到大) |
| converter  | 转换器 (用于处理类型转换)|
| nested  | 内嵌 |
| showCurrentName  | 是否展示当前的名字 |

### 样式注解

#### @ExcelSheetNameStyle

> 控制表名样式

#### @ExcelSerialNumberCellStyle

> 控制序列号样式

#### @ExcelHeadCellStyle

> 控制标题样式

#### @ExcelContentCellStyle

> 控制内容样式

#### 样式注解配置属性说明

|  属性   | 作用  |
|  ----  | ----  |
| styleConfig  | 根据默认构造方法创建对象 管理样式配置<br/>可以通过配置当前属性, 自定义单元格样式, 而非局限于当前工具提供的样式类型(如有必要 可以忽略定义的所有属性 进行覆盖) |
| fontColor | 字体颜色 |
| bgColor | 背景颜色 |
| borderColor | 边框颜色 |
| borderStyle | 边框样式 |
| bold | 字体是否加粗 |
