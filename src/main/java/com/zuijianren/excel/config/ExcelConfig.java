package com.zuijianren.excel.config;

import com.zuijianren.excel.config.style.AbstractCellStyleConfig;
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


    // todo 设置默认样式  便于工具使用者进行使用

    /**
     * sheet 名字 配置
     */
    private AbstractCellStyleConfig sheetNameStyleConfig;
    /**
     * 标题单元格样式
     */
    private AbstractCellStyleConfig headCellStyleConfig;

    /**
     * 内容单元格样式
     */
    private AbstractCellStyleConfig contentCellStyleConfig;

}
