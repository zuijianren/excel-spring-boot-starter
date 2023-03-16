package com.zuijianren.excel.config.style;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel 单元格样式封装对象 规范
 *
 * @author zuijianren
 * @date 2023/3/16 13:07
 */
@Data
public abstract class AbstractCellStyleConfig {

    // 字体颜色
    private String fontColor;

    // 背景颜色
    private String bgColor;


    public CellStyle createCellStyle(XSSFWorkbook xssfWorkbook) {
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        return cellStyle;
    }
}
