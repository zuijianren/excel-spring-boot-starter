package com.zuijianren.excel.config.style;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
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
    private short fontColor = IndexedColors.BLACK.index;

    // 背景颜色
    private short bgColor = IndexedColors.WHITE.index;

    // 边框颜色
    private short borderColor = IndexedColors.BLACK.index;

    // 边框样式
    private BorderStyle borderStyle = BorderStyle.THIN;

    // 粗细
    private boolean bold = false;

    public CellStyle createCellStyle(XSSFWorkbook xssfWorkbook) {
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();

        // 设置字体样式
        Font titleFont = xssfWorkbook.createFont();
        titleFont.setBold(bold); // 设置为粗体
        titleFont.setColor(fontColor);
        cellStyle.setFont(titleFont);

        // 设置背景样式
        cellStyle.setFillBackgroundColor(bgColor);
        cellStyle.setFillForegroundColor(bgColor);
        cellStyle.setFillPattern(FillPatternType.LEAST_DOTS);

        // 设置边框
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setTopBorderColor(borderColor);
        cellStyle.setBottomBorderColor(borderColor);
        cellStyle.setLeftBorderColor(borderColor);
        cellStyle.setRightBorderColor(borderColor);

        //水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        return cellStyle;
    }
}
