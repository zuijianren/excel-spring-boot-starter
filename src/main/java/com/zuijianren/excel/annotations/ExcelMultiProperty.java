package com.zuijianren.excel.annotations;


import com.zuijianren.excel.converter.DefaultExcelConverter;
import com.zuijianren.excel.converter.ExcelConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标识一对多的属性
 * <p>
 * 相对于 ExcelProperty  多一个内嵌属性  用于解析数据
 *
 * @author zuijianren
 * @date 2023/3/13 10:27
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelMultiProperty {

    // 表头 (多级)
    String[] value() default {""};

    // 表头的顺序  从小到大
    int order() default -1;

    // 转换器 用于处理类型转换 将当前对象转换为字符串
    Class<? extends ExcelConverter<?>> converter() default DefaultExcelConverter.class;

    // 内嵌
    boolean nested() default false;


    // 是否展示当前的名字
    boolean showCurrentName() default false;
}
