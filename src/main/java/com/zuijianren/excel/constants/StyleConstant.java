package com.zuijianren.excel.constants;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 样式常量 (用于统一管理样式)
 *
 * @author zuijianren
 * @date 2023/3/17 11:26
 */
public class StyleConstant {

    /* sheet name 默认样式配置 */
    public static final IndexedColors SheetName_BgColor = IndexedColors.GREY_25_PERCENT;
    public static final boolean SheetName_Bold = true;

    /* head 默认样式配置 */
    public static final IndexedColors Head_BgColor = IndexedColors.GREY_40_PERCENT;
    public static final boolean Head_Bold = true;

    /* serialNumber 默认样式配置 */
    public static final IndexedColors SerialNumber_BgColor = IndexedColors.GREY_40_PERCENT;
    public static final boolean SerialNumber_Bold = true;

}
