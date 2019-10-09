package com.sitech.crm.bcc.common.excelparser.interfaces;

/**
 * @Author: liaoyq
 * @Desrition:
 * @Created: 2018/9/4 19:11
 * @Modified: 2018/9/4 19:11
 * @Modified By: liaoyq
 */
public interface ParserExceptionConsumer<T> {
    /**
     * Performs this operation on the given argument.
     *
     * @param t the input argument
     */
    void accept(T t);
}
