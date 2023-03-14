package com.zuijianren.excel.exceptions;

/**
 * @author zuijianren
 * @date 2023/3/14 10:10
 */
public class FileCreateException extends RuntimeException {

    public FileCreateException(Throwable cause) {
        super("目标文件创建失败");
    }
}
