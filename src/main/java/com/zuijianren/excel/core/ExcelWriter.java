package com.zuijianren.excel.core;

import com.zuijianren.excel.config.ExcelConfig;
import com.zuijianren.excel.config.PropertyConfig;
import com.zuijianren.excel.config.SheetConfig;
import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
import com.zuijianren.excel.converter.ExcelConverter;
import com.zuijianren.excel.pojo.ExcelData;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

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


    public static ExcelWriter createExcelWriter(String filePath) throws IOException {
        return createExcelWriter(filePath, null);
    }

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
        // 如果存在 则删除原文件
        file.deleteOnExit();
        file.createNewFile(); // 返回结果忽略
        return createExcelWriter(new FileOutputStream(file), excelConfig);
    }

    public static ExcelWriter createExcelWriter(OutputStream os) throws IOException {
        return createExcelWriter(os, null);
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

            XSSFSheet sheet = xssfWorkbook.createSheet(sheetConfig.getSheetName());

            // 首行创建
            if (sheetConfig.isShowSheetName()) {
                writeSheetName(sheet, sheetConfig, rowPosition++);
            }

            // 表头创建
            writeHead(sheet, rowPosition, sheetConfig);
            rowPosition = rowPosition + sheetConfig.getRowNum();

            // 冻结首行和表头
            if (sheetConfig.isFreezeHead()) {
                sheet.createFreezePane(sheetConfig.getColNum(), rowPosition, 0, 0);
            }

            // 内容创建
            writeContent(sheet, rowPosition, sheetConfig, dataList);

        }
        xssfWorkbook.write(os);
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

            int rowNum = writeData(sheet, rowPosition, colPosition, data, propertyConfigList);
            rowPosition += rowNum;
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
        Map<Integer, CellStyle> singleColumnMap = new HashMap<>();  // 存储非multi的列和样式
        for (PropertyConfig propertyConfig : propertyConfigList) {

            CellStyle cellStyle = getContentCellStyle(propertyConfig);

            // 获取转换器
            ExcelConverter converter = propertyConfig.getConverter();

            // 获取 对应数据
            Field field = propertyConfig.getField();
            Object value = getValue(field, data, converter); // 获取填充数据

            // 写入数据
            if (!propertyConfig.isMulti()) {
                if (!propertyConfig.isNested()) {
                    // 普通属性  直接写入 并且 colPosition 加1
                    writeCell(sheet, rowPosition, colPosition++, value, propertyConfig.getWriteType(), cellStyle);
                    singleColumnMap.put(colPosition - 1, cellStyle);
                } else {
                    // 嵌套属性则依次递归调用当前方法  依次写入
                    writeData(sheet, rowPosition, colPosition, value, propertyConfig.getChildPropertyConfigList());
                    colPosition += propertyConfig.getColNum();
                    for (int i = colPosition - propertyConfig.getColNum(); i < colPosition; i++) {
                        singleColumnMap.put(i, cellStyle);
                    }
                }
                rowNumCount = Math.max(rowNumCount, 1); // 仅占一行
            } else {
                Collection collection = (Collection) value; // 强转为 集合
                int currentRowPosition = rowPosition; // 标记行位置

                // 非嵌套对象
                if (!propertyConfig.isNested()) {
                    for (Object currentValue : collection) {
                        // 写入多行数据
                        writeCell(sheet, currentRowPosition++, colPosition, currentValue, propertyConfig.getWriteType(), cellStyle);
                    }
                    colPosition++;
                    rowNumCount = Math.max(rowNumCount, collection.size()); // 占多行
                } else {
                    // 嵌套对象

                    // 集合对象 判空  空对象会导致for循环报错  需要单独处理
                    if (collection == null) {
                        // 写一行数据
                        writeData(sheet, currentRowPosition, colPosition, value, propertyConfig.getChildPropertyConfigList());
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
                    colPosition += propertyConfig.getColNum(); // 写入完成后 列数 对应 增加
                }
            }
        }
        for (Integer mergeColPosition : singleColumnMap.keySet()) {
            mergeRowCell(sheet, rowPosition, mergeColPosition, rowNumCount, singleColumnMap.get(mergeColPosition));
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
            CellStyle cellStyle = getSerialNumberCellStyle(sheetConfig);
            mergeRowCell(sheet, rowPosition, colPosition, size, cellStyle);
            writeCell(sheet, rowPosition, colPosition++, "序号");
        }

        // 表头写入
        List<PropertyConfig> propertyConfigList = sheetConfig.getPropertyConfigList();
        writeHead(sheet, propertyConfigList, rowPosition, colPosition, size);

        // 合并同名表头
        mergeSameHead(sheet, rowPosition, sheetConfig.getRowNum(), sheetConfig.getColNum());
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

            CellStyle cellStyle = getHeadCellStyle(propertyConfig);

            int currentRowPosition = rowPosition; // 当前行数 每个属性单独计算
            if (!propertyConfig.isNested()) { // 处理非内嵌属性
                String[] value = propertyConfig.getValue();
                // 每一个值占一行  最后一行占剩余所有行
                int rem = size; // 剩余可分配行数
                for (int i = 0; i < value.length; i++) {
                    //  最后一行  合并当前剩余行
                    if (i == value.length - 1) {
                        mergeRowCell(sheet, currentRowPosition, colPosition, rem, cellStyle);
                    }
                    writeCell(sheet, currentRowPosition++, colPosition, value[i], String.class, cellStyle);
                    rem--; // 余量减1
                }
            } else {    // 处理内嵌属性
                int rem = size; // 剩余可分配行数
                // 如果要展示当前表头 则写入当前表头
                if (propertyConfig.isShowCurrentName()) {
                    String[] value = propertyConfig.getValue();
                    for (int i = 0; i < value.length; i++) {
                        writeCell(sheet, currentRowPosition, colPosition, value[i], String.class, cellStyle); // 写入表头
                        mergeColCell(sheet, currentRowPosition, colPosition, propertyConfig.getColNum(), cellStyle); // 横向合并
                        currentRowPosition++;
                        rem--; // 余量减1
                    }
                }
                List<PropertyConfig> childPropertyConfigList = propertyConfig.getChildPropertyConfigList();
                writeHead(sheet, childPropertyConfigList, currentRowPosition, colPosition, rem);
            }
            colPosition = colPosition + propertyConfig.getColNum();
        }
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
                    cell = selectRow.createCell(j);
                }
                String currentValue = cell.getStringCellValue(); // 默认标题头都是字符串类型
                assert currentValue != null; // For blank cells we return an empty string
                if (flagValue.equals(currentValue) && !flagValue.equals("")) {
                    repeatNum++;
                } else {
                    if (repeatNum > 1) {
                        // 获取之前单元格的样式
                        mergeRepeatCol(sheet, currentRowPosition, j, repeatNum);
                    }
                    // 更新值
                    flagValue = currentValue;
                    repeatNum = 1;
                }
            }
            // 获取之前单元格的样式
            mergeRepeatCol(sheet, currentRowPosition, colRange, repeatNum);
        }
    }

    /**
     * 合并重复的列
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行
     * @param colPosition 列位置(需要合并的单元格不包含当前列位置)
     * @param repeatNum   重复树
     */
    private void mergeRepeatCol(XSSFSheet sheet, int rowPosition, int colPosition, int repeatNum) {
        // 获取之前单元格的样式
        XSSFRow row = sheet.getRow(rowPosition);
        XSSFCell previousCell = row.getCell(colPosition - 1);
        XSSFCellStyle cellStyle = previousCell.getCellStyle();
        // 合并之前的重复单元格
        mergeColCell(sheet, rowPosition, colPosition - repeatNum, repeatNum, cellStyle);
    }


    /**
     * 写入 sheetName
     *
     * @param sheet       sheet 对象
     * @param rowPosition 行位置
     * @param sheetConfig sheet 配置
     */
    private void writeSheetName(XSSFSheet sheet, SheetConfig sheetConfig, int rowPosition) {
        CellStyle cellStyle = getSheetNameStyle(sheetConfig);
        writeCell(sheet, rowPosition, 0, sheetConfig.getSheetName(), String.class, cellStyle);
        mergeColCell(sheet, rowPosition, 0, sheetConfig.getColNum(), cellStyle);
    }


    /**
     * 获取 content 单元格样式
     *
     * @param propertyConfig property 配置
     * @return content 单元格样式
     */
    private CellStyle getContentCellStyle(PropertyConfig propertyConfig) {
        AbstractCellStyleConfig contentCellStyleConfig = propertyConfig.getContentCellStyleConfig();
        if (contentCellStyleConfig == null) {
            contentCellStyleConfig = excelConfig.getContentCellStyleConfig();
        }
        assert contentCellStyleConfig != null;
        return contentCellStyleConfig.createCellStyle(excelConfig.getXssfWorkbook());
    }

    /**
     * 获取表头的单元格样式
     *
     * @param propertyConfig property 配置
     * @return 表头的单元格样式
     */
    private CellStyle getHeadCellStyle(PropertyConfig propertyConfig) {
        // 获取 当前属性 标题 的 样式
        AbstractCellStyleConfig cellStyleConfig = propertyConfig.getHeadCellStyleConfig();
        if (cellStyleConfig == null) {
            cellStyleConfig = excelConfig.getHeadCellStyleConfig();
        }
        assert cellStyleConfig != null;
        return cellStyleConfig.createCellStyle(excelConfig.getXssfWorkbook());
    }

    /**
     * 获取 序列号 的单元格 样式
     *
     * @param sheetConfig sheet 配置
     * @return 序列号的单元格样式
     */
    private CellStyle getSerialNumberCellStyle(SheetConfig sheetConfig) {
        AbstractCellStyleConfig serialNumberStyleConfig = sheetConfig.getSerialNumberStyleConfig();
        if (serialNumberStyleConfig == null) {
            serialNumberStyleConfig = excelConfig.getSerialNumberStyleConfig();
        }
        assert serialNumberStyleConfig != null;
        return serialNumberStyleConfig.createCellStyle(excelConfig.getXssfWorkbook());
    }

    /**
     * 获取 sheetName 的单元格样式
     *
     * @param sheetConfig sheet 配置
     * @return sheetName 的单元格样式
     */
    private CellStyle getSheetNameStyle(SheetConfig sheetConfig) {
        AbstractCellStyleConfig sheetNameStyleConfig = sheetConfig.getSheetNameStyleConfig();
        // 如果没有配置样式 则获取 excelConfig 中的样式(必定包含默认样式)
        if (sheetNameStyleConfig == null) {
            sheetNameStyleConfig = excelConfig.getSheetNameStyleConfig();
        }
        assert sheetNameStyleConfig != null;
        return sheetNameStyleConfig.createCellStyle(excelConfig.getXssfWorkbook());
    }


    /**
     * 获取写入数据
     *
     * @param field     属性
     * @param data      数据对象
     * @param converter 转换器
     * @return 写入数据
     */
    private Object getValue(Field field, Object data, ExcelConverter converter) {
        Object value = null;
        if (data != null) {
            try {
                value = field.get(data); // 获取对应数据
            } catch (IllegalAccessException e) {
                e.printStackTrace(); // 忽略即可
            }
            if (converter != null) {
                value = converter.convert(value);
            }
        }
        return value;
    }

}
