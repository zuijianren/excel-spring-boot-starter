package com.zuijianren.excel.converter;

/**
 * 默认转换器
 *
 * @author zuijianren
 * @date 2023/3/13 10:38
 */
public class StringExcelConverter implements ExcelConverter<Integer, String> {

    @Override
    public String convert(Integer i) {
        return i + "";
    }

    public StringExcelConverter() {
        System.out.println("construct");
    }
}
