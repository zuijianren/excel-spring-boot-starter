package com.zuijianren.excel;

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
public class Teacher2 {

    private Integer id;
    @ExcelProperty(value = {"年龄"})
    private Integer age;


    //    @ExcelMultiProperty(value = "昵称")
    private List<String> nickNameList;


    @ExcelProperty(value = {"名字"})
    private String name;


    public Teacher2(Integer id, Integer age, String name) {
        this.id = id;
        this.age = age;
        this.name = name;
    }
}
