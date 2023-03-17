package com.zuijianren.excel.config.style;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * excel 单元格样式 封装对象
 *
 * @author zuijianren
 * @date 2023/3/16 12:39
 */
@Data
@Builder
public class SheetNameStyleConfig extends AbstractCellStyleConfig {


    public SheetNameStyleConfig() {
        /*=== 设置默认值 ===*/
        // 灰色背景
        setBgColor(IndexedColors.GREY_25_PERCENT.index);
        // 字体加粗
        setBold(true);
    }
}
