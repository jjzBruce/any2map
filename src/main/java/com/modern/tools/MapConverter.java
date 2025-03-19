package com.modern.tools;

import java.util.List;
import java.util.Map;

/**
 * 转换器接口
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public interface MapConverter<T extends ConvertConfig> {

    /**
     * 转换设置
     */
    void setConvertConfig(T config);

    /**
     * 根据源数据输出目标数据
     */
    Map<String, Object> toMap(Object source);

}
