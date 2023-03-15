package com.zuijianren.excel.core;

import com.zuijianren.excel.config.ExcelConfig;
import com.zuijianren.excel.config.PropertyConfig;
import com.zuijianren.excel.config.SheetConfig;
import com.zuijianren.excel.pojo.ExcelData;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import static com.zuijianren.excel.core.ExcelOperator.*;

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

    private List<ExcelData> dataSource = new ArrayList<>(); // 渲染数据的缓存


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
        dataSource.add(new ExcelData(sheetConfig, data)); // 存储配置和数据  等待写入
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
        for (ExcelData excelData : dataSource) {
            SheetConfig sheetConfig = excelData.getSheetConfig();
            Collection dataList = excelData.getDataList();

            int rowPosition = 0; // 行号

            XSSFSheet sheet = xssfWorkbook.createSheet();

            // 首行创建
            if (sheetConfig.isShowSheetName()) {
                mergeColCell(sheet, rowPosition, 0, sheetConfig.getColNum());
                writeCell(sheet, rowPosition++, 0, sheetConfig.getSheetName()); // 写入完成后 行数加1
            }

            // 表头创建
            writeHead(sheet, rowPosition, sheetConfig);
            mergeSameHead(sheet, rowPosition, sheetConfig.getRowNum(), sheetConfig.getColNum());
            rowPosition = rowPosition + sheetConfig.getRowNum();

            // 内容创建
            writeContent(sheet, rowPosition, sheetConfig, dataList);

        }
        xssfWorkbook.write(os);
    }

    /**
     * 合并同名单元格(横向合并)
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行起始位置
     * @param rowRange    行范围
     * @param colRange    列范围
     */
    private void mergeSameHead(XSSFSheet sheet, int rowPosition, int rowRange, int colRange) {
        for (int i = 0; i < rowRange; i++) {
            int currentRowPosition = rowPosition + i;
            XSSFRow selectRow = sheet.getRow(currentRowPosition); // 选中行
            String flagValue = ""; // 标记值
            int repeatNum = 0; // 重复值
            for (int j = 0; j < colRange; j++) {
                XSSFCell cell = selectRow.getCell(j);
                if (cell == null) {
                    // 合并之前的重复单元格
                    mergeColCell(sheet, currentRowPosition, j - repeatNum, repeatNum);
                    flagValue = "";
                    repeatNum = 0;
                    continue; //跳过
                }
                String currentValue = cell.getStringCellValue(); // 默认标题头都是字符串类型
                assert currentValue != null;
                if (flagValue.equals(currentValue)) {
                    repeatNum++;
                } else {
                    // 合并之前的重复单元格
                    mergeColCell(sheet, currentRowPosition, j - repeatNum, repeatNum);
                    // 更新值
                    flagValue = currentValue;
                    repeatNum = 1;
                }
            }
            mergeColCell(sheet, currentRowPosition, colRange - repeatNum, repeatNum); // 退出循环前再合并一次
        }
    }

    /**
     * 写入内容
     *
     * @param sheet       sheet对象
     * @param rowPosition 起始行
     * @param dataList    数据
     */
    private void writeContent(XSSFSheet sheet, int rowPosition, SheetConfig sheetConfig, Collection dataList) {
        int index = 1; // 索引 从1开始计数
        for (Object data : dataList) {
            int colPosition = 0;
            // 序号处理
            if (sheetConfig.isShowSerialNumber()) {
                writeCell(sheet, rowPosition, colPosition++, index);
            }

            List<PropertyConfig> propertyConfigList = sheetConfig.getPropertyConfigList();

//            int i = writeData(sheet, rowPosition, colPosition, data, propertyConfigList);
        }
    }

    /**
     * 写入一条数据
     *
     * @param sheet              sheet 对象
     * @param rowPosition        行
     * @param colPosition        列
     * @param data               数据
     * @param propertyConfigList 配置集合
     * @return 数据所需行数(用于后续合并单元格)
     */
    private int writeData(XSSFSheet sheet, int rowPosition, int colPosition, Object data, List<PropertyConfig> propertyConfigList) {
        int rowNumCount = 0; // 统计当前数据所占行数
        for (PropertyConfig propertyConfig : propertyConfigList) {
            // 获取 对应数据
            Field field = propertyConfig.getField();
            Object value = null;
            try {
                value = field.get(data); // 获取对应数据
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }

            // 写入数据
            if (!propertyConfig.isMulti()) {
                // 普通属性  直接写入 并且 colPosition 加1
                // todo ExcelConverter 调用及实现
                // todo 空值 替代值 处理
                writeCell(sheet, rowPosition, colPosition++, value, propertyConfig.getType());
                rowNumCount = Math.max(rowNumCount, 1); // 仅占一行
            } else {
                Collection collection = (Collection) value; // 强转为 集合
                int currentRowPosition = rowPosition;
                // 非嵌套对象
                if (!propertyConfig.isNested()) {
                    for (Object currentValue : collection) {
                        // 写入多行数据
                        writeCell(sheet, currentRowPosition++, colPosition, currentValue, propertyConfig.getType());
                    }
                    colPosition++;
                    rowNumCount = Math.max(rowNumCount, collection.size()); // 占多行
                } else {
                    // 嵌套对象
                    if (value == null) {
                        rowNumCount = Math.max(rowNumCount, 1);
                    } else {
                        int totalChildRowNum = 0; // 对应对象所占的总行数
                        for (Object currentValue : collection) {
                            int childRowNum = writeData(sheet, currentRowPosition, colPosition, currentValue, propertyConfig.getChildPropertyConfigList());
                            totalChildRowNum += childRowNum; // 更新总行数
                            currentRowPosition += childRowNum; // 更新写入的位置
                        }
                        rowNumCount = Math.max(rowNumCount, totalChildRowNum);
                    }
                    colPosition += propertyConfig.getChildPropertyConfigList().size(); // 写入完成后 列数 对应 增加
                }
            }
        }
        return rowNumCount;
    }


    /**
     * 写入表头
     *
     * @param sheet       sheet对象
     * @param rowPosition 起始行
     */
    private void writeHead(XSSFSheet sheet, int rowPosition, SheetConfig sheetConfig) {
        int size = sheetConfig.getRowNum();
        int colPosition = 0;
        // 处理序号列
        if (sheetConfig.isShowSerialNumber()) {
            mergeRowCell(sheet, rowPosition, colPosition, size);
            writeCell(sheet, rowPosition, colPosition++, "序号");
        }
        // 表头写入
        List<PropertyConfig> propertyConfigList = sheetConfig.getPropertyConfigList();

        writeHead(sheet, propertyConfigList, rowPosition, colPosition, size);
    }

    /**
     * 写入表头  (根据 propertyConfigList 进行写入)
     *
     * @param sheet              sheet 对象
     * @param propertyConfigList 属性集合
     * @param rowPosition        行位置
     * @param colPosition        列位置
     * @param size               大小
     */
    private void writeHead(XSSFSheet sheet, List<PropertyConfig> propertyConfigList, int rowPosition, int colPosition, int size) {
        for (PropertyConfig propertyConfig : propertyConfigList) {
            int currentRowPosition = rowPosition; // 当前行数 每个属性单独计算
            if (!propertyConfig.isNested()) { // 处理非内嵌属性
                String[] value = propertyConfig.getValue();
                // 每一个值占一行  最后一行占剩余所有行
                int rem = size; // 剩余可分配行数
                for (int i = 0; i < value.length; i++) {
                    //  最后一行  合并当前剩余行
                    if (i == value.length - 1) {
                        mergeRowCell(sheet, currentRowPosition, colPosition, rem);
                    }
                    writeCell(sheet, currentRowPosition++, colPosition, value[i]);
                    rem--; // 余量减1
                }
            } else {    // 处理内嵌属性
                int rem = size; // 剩余可分配行数
                // 如果要展示当前表头 则写入当前表头
                if (propertyConfig.isShowCurrentName()) {
                    String[] value = propertyConfig.getValue();
                    for (int i = 0; i < value.length; i++) {
                        mergeColCell(sheet, currentRowPosition, colPosition, propertyConfig.getColNum()); // 横向合并
                        writeCell(sheet, currentRowPosition++, colPosition, value[i]); // 写入表头
                        rem--; // 余量减1
                    }
                }
                List<PropertyConfig> childPropertyConfigList = propertyConfig.getChildPropertyConfigList();
                writeHead(sheet, childPropertyConfigList, currentRowPosition, colPosition, rem);
            }
            colPosition = colPosition + propertyConfig.getColNum();
        }
    }


}
