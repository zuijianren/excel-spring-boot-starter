package com.zuijianren.excel.config;

import com.zuijianren.excel.config.style.*;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel 文件整体配置
 *
 * @author zuijianren
 * @date 2023/3/13 10:48
 */
@Builder
@Data
public class ExcelConfig {

    // 工作簿
    @Builder.Default
    private XSSFWorkbook xssfWorkbook = new XSSFWorkbook();


    /**
     * sheet 名字 配置
     */
    @Builder.Default
    private AbstractCellStyleConfig sheetNameStyleConfig = new SheetNameStyleConfig();

    /**
     * 序号列样式
     */
    @Builder.Default
    private AbstractCellStyleConfig serialNumberStyleConfig = new SerialNumberStyleConfig();

    /**
     * 标题单元格样式
     */
    @Builder.Default
    private AbstractCellStyleConfig headCellStyleConfig = new HeadCellStyleConfig();

    /**
     * 内容单元格样式
     */
    @Builder.Default
    private AbstractCellStyleConfig contentCellStyleConfig = new ContentCellStyleConfig();


}
