package com.zuijianren.excel;

import com.zuijianren.excel.annotations.ExcelProperty;
import com.zuijianren.excel.annotations.ExcelSheet;
import lombok.Data;

/**
 * @author zuijianren
 * @date 2023/3/13 14:03
 */
@ExcelSheet("学生表")
@Data
public class Student {

    private Integer id;
    @ExcelProperty(value = {"年龄"}, order = 2)
    private Integer age;
    @ExcelProperty(value = {"名字"}, order = 1)
    private String name;

    public Student(Integer id, Integer age, String name) {
        this.id = id;
        this.age = age;
        this.name = name;
    }
}
