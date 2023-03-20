package com.zuijianren.excel.core;

import com.zuijianren.excel.exceptions.WriteToCellException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

/**
 * @author zuijianren
 * @date 2023/3/15 13:04
 */
public class ExcelOperator {

    /**
     * 向指定单元格中写入值
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param value       值
     */
    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, boolean value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, double value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, Date value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, LocalDateTime value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, LocalDate value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, Calendar value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, String value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }

    public static XSSFCell writeCell(XSSFSheet sheet, int rowPosition, int colPosition, RichTextString value) {
        XSSFRow row = sheet.getRow(rowPosition);
        if (row == null) {
            row = sheet.createRow(rowPosition);
        }
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
        return cell;
    }


    /**
     * 写入 Object 对象的方法
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param value       值
     * @param type        对应类型
     */
    public static void writeCell(XSSFSheet sheet, int rowPosition, int colPosition, Object value, Class<?> type, CellStyle cellStyle) {
        // 如果数据为空 则按照空字符串写入
        if (value == null) {
            writeCell(sheet, rowPosition, colPosition, "", String.class, cellStyle);
            return;
        }
        XSSFCell cell = null;
        // 根据 type 进行解析
        if (Integer.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (Integer) value);
        } else if (Double.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (Double) value);
        } else if (Date.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (Date) value);
        } else if (LocalDateTime.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (LocalDateTime) value);
        } else if (LocalDate.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (LocalDate) value);
        } else if (Calendar.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (Calendar) value);
        } else if (String.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (String) value);
        } else if (RichTextString.class.isAssignableFrom(type)) {
            cell = writeCell(sheet, rowPosition, colPosition, (RichTextString) value);
        } else {
            throw new WriteToCellException(type.getName());
        }
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 写入 Object 对象的方法
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param value       值
     * @param type        对应类型
     */
    public static void writeCell(XSSFSheet sheet, int rowPosition, int colPosition, Object value, Class<?> type) {
        writeCell(sheet, rowPosition, colPosition, value, type, null);
    }


    /**
     * 纵向合并单元格
     *
     * @param sheet       sheet对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param num         合并数量
     * @param cellStyle   单元格样式
     */
    public static void mergeColCell(XSSFSheet sheet, int rowPosition, int colPosition, int num, CellStyle cellStyle) {
        // 小于2 则 无需合并
        if (num < 2) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(rowPosition, rowPosition, colPosition, colPosition + num - 1);
        sheet.addMergedRegion(region);
        // 合并后 追加样式
        for (int i = region.getFirstColumn(); i <= region.getLastColumn(); i++) {
            XSSFRow row = sheet.getRow(rowPosition);
            if (row == null) {
                row = sheet.createRow(rowPosition);
            }
            XSSFCell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 横向合并单元格
     *
     * @param sheet       sheet对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param num         合并数量
     * @param cellStyle   单元格样式
     */
    public static void mergeRowCell(XSSFSheet sheet, int rowPosition, int colPosition, int num, CellStyle cellStyle) {
        // 小于2 则 无需合并
        if (num < 2) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(rowPosition, rowPosition + num - 1, colPosition, colPosition);
        sheet.addMergedRegion(region);
        // 合并后 追加样式
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            XSSFCell cell = row.getCell(colPosition);
            if (cell == null) {
                cell = row.createCell(colPosition);
            }
            cell.setCellStyle(cellStyle);
        }
    }


}
