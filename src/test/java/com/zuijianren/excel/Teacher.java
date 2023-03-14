package com.zuijianren.excel;

import com.zuijianren.excel.annotations.ExcelMultiProperty;
import com.zuijianren.excel.annotations.ExcelProperty;
import com.zuijianren.excel.annotations.ExcelSheet;
import lombok.Data;

import java.util.List;

/**
 * @author zuijianren
 * @date 2023/3/13 14:03
 */
@ExcelSheet("教师表")
@Data
public class Teacher {

    private Integer id;
    @ExcelProperty(value = {"年龄"})
    private Integer age;
    @ExcelProperty(value = {"名字"})
    private String name;

    @ExcelMultiProperty(value = "昵称")
    private List<String> nickNameList;

    @ExcelMultiProperty(value = "管理的学生", nested = true)
    private List<Student> studentList;

    public Teacher(Integer id, Integer age, String name) {
        this.id = id;
        this.age = age;
        this.name = name;
    }
}
