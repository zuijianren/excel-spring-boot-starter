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

标识当前对象是一个 sheet 对象

### @ExcelProperty

标识当前属性为需要导出的属性

### @ExcemMultiProperty

标识当前属性为需要导出的集合属性(渲染时将会按照一对多的方式进行渲染)
> 如果集合属性仅会一个字符串写入(例如拼接的方式), 则应使用 ExcelProperty 注解  需要实现 ExcelConverter 接口 并进行声明 才能成功写入

### 样式注解

* @ExcelSheetNameStyle - 控制表名样式
* @ExcelSerialNumberCellStyle - 控制序列号样式
* @ExcelHeadCellStyle - 控制标题样式
* @ExcelContentCellStyle - 控制内容样式
