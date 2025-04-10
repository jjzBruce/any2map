package io.github.jjzbruce;

import java.io.Serializable;
import java.util.*;

/**
 * 数据头
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.1
 */
public class MapHeaders implements Serializable {

    protected Set<MapHeader> headers = new LinkedHashSet<>();

    public boolean addHead(MapHeader header) {
        return headers.add(header);
    }

    public MapHeader getHeader(String headerName, int level) {
        return headers.stream().filter(x -> x.getName().equals(headerName) && x.getLevel() == level)
                .findFirst().orElse(null);
    }

    public MapHeader getHeader(String headerName, int level, int index) {
        int i = 0;
        for (MapHeader header : headers) {
            if (Objects.equals(header.getName(), headerName) && header.getLevel() == level) {
                if (index == (i++)) {
                    return header;
                }
            }
        }
        return null;
    }

    public Set<MapHeader> getHeaders() {
        return headers;
    }
}
