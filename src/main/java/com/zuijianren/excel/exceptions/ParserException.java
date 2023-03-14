package com.zuijianren.excel.exceptions;

/**
 * 解析异常
 *
 * @author zuijianren
 * @date 2023/3/13 11:09
 */
public class ParserException extends RuntimeException {

    public ParserException(String msg, Throwable cause) {
        super(msg, cause);
    }

    public ParserException(String msg) {
        super(msg);
    }
}
