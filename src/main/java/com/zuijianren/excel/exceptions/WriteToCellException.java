package com.zuijianren.excel.exceptions;

/**
 * 数据写入单元格异常
 *
 * @author zuijianren
 * @date 2023/3/13 11:09
 */
public class WriteToCellException extends RuntimeException {

    public WriteToCellException(String className) {
        super("未找到数据类型: '" + className + "' 对应的写入单元格的方法");
    }

}
