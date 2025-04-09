package io.github.jjzbruce;

import java.io.Serializable;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 数据头
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.1
 */
public class MapHeaders implements Serializable {

    protected Set<MapHeader> headers = new LinkedHashSet<>();

    /**
     * 添加标题
     */
    public boolean addHeader(MapHeader header) {
        return headers.add(header);
    }

    public MapHeader getHeader(String headerName, int level) {
        return headers.stream().filter(x -> x.getHeader().equals(headerName) && x.getLevel() == level)
                .findFirst().orElse(null);
    }

    public Set<MapHeader> getHeaders() {
        return headers;
    }
}
