package com.zuijianren.excel.annotations.style;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel sheet name 样式配置注解
 *
 * @author zuijianren
 * @date 2023/3/16 13:02
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface ExcelSheetNameStyle {

}
