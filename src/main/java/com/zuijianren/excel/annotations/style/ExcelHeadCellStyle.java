package com.zuijianren.excel.annotations.style;

import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel 表头样式配置注解
 *
 * @author zuijianren
 * @date 2023/3/16 13:02
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface ExcelHeadCellStyle {

    /**
     * 根据默认构造方法创建对象 管理样式配置
     * <p>
     * 可以通过配置当前属性, 自定义单元格样式, 而非局限于当前工具提供的样式类型(如有必要 可以忽略定义的所有属性 进行覆盖)
     */
    Class<? extends AbstractCellStyleConfig> styleConfig() default AbstractCellStyleConfig.class;

    // 字体颜色
    IndexedColors fontColor() default IndexedColors.BLACK;

    // 背景颜色
    IndexedColors bgColor() default IndexedColors.GREY_40_PERCENT;

    // 边框颜色
    IndexedColors borderColor() default IndexedColors.BLACK;

    // 边框样式
    BorderStyle borderStyle() default BorderStyle.THIN;

    // 字体是否加粗
    boolean bold() default true;

}
