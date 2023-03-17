package com.zuijianren.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标识当前对象 为一个 ExcelSheet 对象
 *
 * @author zuijianren
 * @date 2023/3/13 10:23
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelSheet {

    // sheet 表名
    String value() default "";

    // 首行是否展示表名
    boolean showSheetName() default true;

    // 是否展示序列号
    boolean showSerialNumber() default false;

    // 是否冻结首行(sheetName)及表头
    boolean freezeHead() default true;
}
