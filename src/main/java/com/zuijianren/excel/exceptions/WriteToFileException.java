package com.zuijianren.excel.exceptions;

/**
 * 数据写入文件异常
 *
 * @author zuijianren
 * @date 2023/3/13 11:09
 */
public class WriteToFileException extends RuntimeException {

    public WriteToFileException() {
        super("io异常, 写入数据失败");
    }

}
