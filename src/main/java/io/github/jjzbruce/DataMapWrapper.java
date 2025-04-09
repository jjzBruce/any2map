package io.github.jjzbruce;

import java.util.Map;

/**
 * Map数据
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.1
 */
public class DataMapWrapper {

    /**
     * 数据头
     */
    private MapHeaders headers;

    /**
     * 数据体
     */
    private Map<String, Object> data;

    public DataMapWrapper(MapHeaders headers, Map<String, Object> data) {
        this.headers = headers;
        this.data = data;
    }

    public MapHeaders getHeaders() {
        return headers;
    }

    public Map<String, Object> getData() {
        return data;
    }

}
