package com.zuijianren.excel;

import com.zuijianren.excel.annotations.ExcelMultiProperty;
import com.zuijianren.excel.annotations.ExcelProperty;
import com.zuijianren.excel.annotations.ExcelSheet;
import com.zuijianren.excel.annotations.style.ExcelSheetNameStyle;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

/**
 * @author zuijianren
 * @date 2023/3/13 14:03
 */
@ExcelSheet("教师表")
@Data
@ExcelSheetNameStyle(bgColor = IndexedColors.AQUA)
public class Teacher {

    private Integer id;
    @ExcelProperty(value = {"年龄"})
    private Integer age;


    //    @ExcelMultiProperty(value = "昵称")
    private List<String> nickNameList;


    @ExcelMultiProperty(value = "学生", nested = true, showCurrentName = true)
    private List<Student> studentList;

    @ExcelProperty(value = {"名字"})
    private String name;

    @ExcelProperty(value = {"教师"}, nested = true, showCurrentName = true)
    private Teacher2 teacher2;

    public Teacher(Integer id, Integer age, String name) {
        this.id = id;
        this.age = age;
        this.name = name;
    }
}
