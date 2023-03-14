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

    // 表头的顺序   从大到小
    int order() default -1;
}
