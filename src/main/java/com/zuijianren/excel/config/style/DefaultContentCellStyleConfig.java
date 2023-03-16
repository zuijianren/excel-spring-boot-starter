package com.zuijianren.excel.config.style;

import lombok.Builder;
import lombok.Data;

/**
 * excel 单元格样式 封装对象
 * <p>
 * content 内容的样式处理 可能相较于标题有其它处理  因此 分开管理二者
 *
 * @author zuijianren
 * @date 2023/3/16 12:39
 */
@Data
@Builder
public class DefaultContentCellStyleConfig extends AbstractCellStyleConfig {

    public DefaultContentCellStyleConfig() {
    }
}
