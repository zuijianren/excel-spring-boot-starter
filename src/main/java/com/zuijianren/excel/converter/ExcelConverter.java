package com.zuijianren.excel.converter;

/**
 * @author zuijianren
 * @date 2023/3/13 10:38
 */
public interface ExcelConverter<T, R> {

    R convert(T t);
}
