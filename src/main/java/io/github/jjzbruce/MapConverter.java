package io.github.jjzbruce;

import java.util.Map;

/**
 * 转换器接口
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public interface MapConverter<T extends ConvertConfig> {

    /**
     * 配置
     */
    T getConvertConfig();

    /**
     * 输出结果
     */
    default Map<String, Object> toMap() {
        return toMapData().getData();
    }

    /**
     * 输出结果
     */
    DataMapWrapper toMapData();

}
