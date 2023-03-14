package com.zuijianren.excel.config;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

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



}
