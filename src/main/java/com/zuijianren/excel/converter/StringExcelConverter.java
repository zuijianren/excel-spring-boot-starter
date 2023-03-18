package com.zuijianren.excel.converter;

/**
 * 默认转换器
 *
 * @author zuijianren
 * @date 2023/3/13 10:38
 */
public class StringExcelConverter implements ExcelConverter<String> {

    @Override
    public String convert(Object obj) {
        return "318";
    }

    public StringExcelConverter() {
        System.out.println("construct");
    }
}
