package com.zuijianren.excel.config.style;

import lombok.Builder;
import lombok.Data;

import static com.zuijianren.excel.constants.StyleConstant.*;

/**
 * excel 单元格样式 封装对象
 *
 * @author zuijianren
 * @date 2023/3/16 12:39
 */
@Data
@Builder
public class SerialNumberStyleConfig extends AbstractCellStyleConfig {


    public SerialNumberStyleConfig() {
        /*=== 设置默认值 ===*/
        // 灰色背景
        setBgColor(SerialNumber_BgColor.index);
        setBold(SerialNumber_Bold);
    }
}
