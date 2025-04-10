package io.github.jjzbruce;

import java.util.Map;

/**
 * 转换器接口
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public interface MapConverter<T extends ConvertConfig> {

    T getConvertConfig();

    Map<String, Object> toMap();



}
