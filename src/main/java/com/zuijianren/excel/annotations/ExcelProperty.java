package com.zuijianren.excel.annotations;


import com.zuijianren.excel.converter.DefaultExcelConverter;
import com.zuijianren.excel.converter.ExcelConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标识普通属性
 *
 * @author zuijianren
 * @date 2023/3/13 10:27
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelProperty {

    // 表头 (多级)
    String[] value() default {""};

    // 表头的顺序  从小到大
    int order() default -1;

    // 转换器 用于处理类型转换
    Class<? extends ExcelConverter<?>> converter() default DefaultExcelConverter.class;
}
