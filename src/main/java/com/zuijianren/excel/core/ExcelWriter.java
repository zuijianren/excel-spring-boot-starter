package com.zuijianren.excel.core;

import com.zuijianren.excel.config.ExcelConfig;
import com.zuijianren.excel.config.PropertyConfig;
import com.zuijianren.excel.config.SheetConfig;
import com.zuijianren.excel.pojo.ExcelData;
import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * excel 写入数据 工具类
 *
 * @author zuijianren
 * @date 2023/3/14 10:16
 */
@Data
public class ExcelWriter {

    private static ExcelParser parser = ExcelParser.getInstance();  // 解析器


    private OutputStream os; // 输入流
    private ExcelConfig excelConfig; // excel 文件配置

    private List<ExcelData> dataCache = new ArrayList<>(); // 渲染数据的缓存


    /**
     * 创建 ExcelWriter 对象
     *
     * @param filePath    文件路径
     * @param excelConfig excel 配置
     * @return ExcelWriter 对象
     * @throws IOException io 异常  文件未找到 创建失败时抛出
     */
    public static ExcelWriter createExcelWriter(String filePath, ExcelConfig excelConfig) throws IOException {
        File file = new File(filePath);
        return createExcelWriter(file, excelConfig);
    }

    /**
     * 创建 ExcelWriter 对象
     *
     * @param file        文件
     * @param excelConfig excel 配置
     * @return ExcelWriter 对象
     * @throws IOException io 异常 文件未找到 创建失败时抛出
     */
    public static ExcelWriter createExcelWriter(File file, ExcelConfig excelConfig) throws IOException {
        // 校验文件格式
        String fileName = file.getName();
        String suffix = fileName.substring(fileName.lastIndexOf("."));
        if (!".xlsx".equals(suffix)) {
            throw new IllegalArgumentException("文件类型错误. 仅允许写入 xlsx 类型文件");
        }
        // 判断文件是否存在
        if (!file.exists()) {
            // 如果不存在 则 创建文件
            file.createNewFile(); // 返回结果忽略
        }
        return createExcelWriter(new FileOutputStream(file), excelConfig);
    }

    /**
     * 创建 ExcelWriter 对象
     *
     * @param os          输出流
     * @param excelConfig excel 配置
     * @return ExcelWriter 对象
     */
    public static ExcelWriter createExcelWriter(OutputStream os, ExcelConfig excelConfig) {
        return new ExcelWriter(os, excelConfig);
    }

    // 构造方法
    public ExcelWriter(OutputStream os, ExcelConfig excelConfig) {
        this.os = os;
        this.excelConfig = excelConfig;
        if (excelConfig == null) {
            this.excelConfig = ExcelConfig.builder().build();
        }
    }

    /**
     * 添加 写入的数据
     *
     * @param clazz 写入的类
     * @param data  对应的数据
     * @param <T>   限制写入的类型
     * @return 当前对象 便于直接打点调用
     */
    public <T> ExcelWriter write(Class<T> clazz, Collection<T> data) {
        SheetConfig sheetConfig = parser.getSheetConfig(clazz); // 获取配置
        dataCache.add(new ExcelData(sheetConfig, data)); // 存储配置和数据  等待写入
        return this;
    }

    /**
     * 最终执行写入的方法
     * <p>
     * 文件将会被写入事先指定的位置
     *
     * @throws IOException 写入xlsx数据失败
     */
    public void doWrite() throws IOException {
        // 核心部分
        XSSFWorkbook xssfWorkbook = excelConfig.getXssfWorkbook();
        for (ExcelData excelData : dataCache) {
            SheetConfig sheetConfig = excelData.getSheetConfig();
            Collection data = excelData.getData();

            int rowNum = 0; // 行号

            XSSFSheet sheet = xssfWorkbook.createSheet();

            // 首行创建
            if (sheetConfig.isShowSheetName()) {
                mergeColCell(sheet, rowNum, 0, sheetConfig.getColNum());
                writeCell(sheet, rowNum++, 0, sheetConfig.getSheetName()); // 写入完成后 行数加1
            }

            // 表头创建
            writeHead(sheet, rowNum, sheetConfig);
            rowNum = rowNum + sheetConfig.getRowNum();

            // 内容创建
            writeContent(sheet, rowNum, data);

        }
        xssfWorkbook.write(os);
    }

    /**
     * 写入内容
     *
     * @param sheet  sheet对象
     * @param rowNum 起始行
     * @param data   数据
     */
    private void writeContent(XSSFSheet sheet, int rowNum, Collection data) {

    }

    /**
     * 写入表头
     *
     * @param sheet  sheet对象
     * @param rowNum 起始行
     */
    private void writeHead(XSSFSheet sheet, int rowNum, SheetConfig sheetConfig) {
        int size = sheetConfig.getRowNum();
        int colNum = 0;
        // 处理序号列
        if (sheetConfig.isShowSerialNumber()) {
            mergeRowCell(sheet, rowNum, colNum, size);
            XSSFRow row = sheet.createRow(rowNum++);
            XSSFCell cell = row.createCell(colNum++);
            cell.setCellValue("序号");
        }
        // 表头写入
        List<PropertyConfig> propertyConfigList = sheetConfig.getPropertyConfigList();

        // todo 抽取 并 回调

        for (PropertyConfig propertyConfig : propertyConfigList) {
            int currentRowNum = rowNum; // 当前行数 每个属性单独计算
            if (!propertyConfig.isNested()) { // 处理非内嵌属性
                String[] value = propertyConfig.getValue();
                // 每一个值占一行  最后一行占剩余所有行
                int rem = size; // 剩余可分配行数
                for (int i = 0; i < value.length; i++) {
                    //  最后一行  合并当前行
                    if (i == value.length - 1) {
                        mergeRowCell(sheet, currentRowNum, colNum, rem);
                    }
                    writeCell(sheet, currentRowNum++, colNum, value[i]);
                    rem--; // 余量减1
                }
            } else {    // 处理内嵌属性
                if (propertyConfig.isShowCurrentName()) {
                    // 写入当前表头
                    String[] value = propertyConfig.getValue();
                    int rem = size; // 剩余可分配行数
                    for (int i = 0; i < value.length; i++) {
                        writeCell(sheet, currentRowNum++, colNum, value[i]);
                        rem--; // 余量减1
                    }
                    List<PropertyConfig> childPropertyConfigList = propertyConfig.getChildPropertyConfigList();
                    // todo 回调
                }
            }
            colNum++; // 每处理完一个属性 列数加1
        }
    }

    /**
     * 向指定单元格中写入值
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param value       值
     */
    private void writeCell(XSSFSheet sheet, int rowPosition, int colPosition, String value) {
        XSSFRow row = sheet.createRow(rowPosition);
        XSSFCell cell = row.createCell(colPosition);
        cell.setCellValue(value);
    }

    /**
     * 纵向合并单元格
     *
     * @param sheet       sheet对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param num         合并数量
     */
    private void mergeColCell(XSSFSheet sheet, int rowPosition, int colPosition, int num) {
        // 小于2 则 无需合并
        if (num < 2) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(rowPosition, rowPosition, colPosition, colPosition + num - 1));
    }

    /**
     * 横向合并单元格
     *
     * @param sheet       sheet对象
     * @param rowPosition 行
     * @param colPosition 列
     * @param num         合并数量
     */
    private void mergeRowCell(XSSFSheet sheet, int rowPosition, int colPosition, int num) {
        // 小于2 则 无需合并
        if (num < 2) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(rowPosition, rowPosition + num - 1, colPosition, colPosition));
    }

}
