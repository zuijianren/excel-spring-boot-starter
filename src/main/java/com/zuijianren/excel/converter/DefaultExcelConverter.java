package com.zuijianren.excel.converter;

/**
 * 默认转换器
 *
 * @author zuijianren
 * @date 2023/3/13 10:38
 */
public class DefaultExcelConverter implements ExcelConverter<Object, String> {

    @Override
    public String convert(Object obj) {
        System.out.println("test");
        return "318";
    }
}
