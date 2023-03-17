package com.zuijianren.excel.config.style;

import lombok.Data;

import static com.zuijianren.excel.constants.StyleConstant.*;

/**
 * excel 单元格样式 封装对象
 *
 * @author zuijianren
 * @date 2023/3/16 12:39
 */
@Data
public class SheetNameStyleConfig extends AbstractCellStyleConfig {


    public SheetNameStyleConfig() {
        /*=== 设置默认值 ===*/
        // 灰色背景
        setBgColor(SheetName_BgColor.index);
        // 字体加粗
        setBold(SheetName_Bold);
    }
}
