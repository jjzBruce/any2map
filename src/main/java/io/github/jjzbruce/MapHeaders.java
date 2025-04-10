package io.github.jjzbruce;

import java.io.Serializable;
import java.util.*;
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
     * 添加标题，如果存在则更新
     */
    public boolean addUpdateHeader(MapHeader header) {
        MapHeader mh = getHeader(header.getName(), header.getLevel());
        if(mh == null) {
            return headers.add(header);
        } else {
            mh.setParentHeader(header.getParentHeader());
            mh.setLeaf(header.isLeaf());
            String[] mhKts = mh.getKeyTypes();
            String[] keyTypes = header.getKeyTypes();
            if(keyTypes != null || keyTypes.length > 0) {
                List<String> collect = Arrays.stream(keyTypes).filter(x -> {
                    for (String kt : mhKts) {
                        if (kt.equals(x)) {
                            return false;
                        }
                    }
                    return true;
                }).collect(Collectors.toList());
                if(!collect.isEmpty()) {
                    String[] mhKts2 = new String[mhKts.length + collect.size()];
                    System.arraycopy(mhKts, 0, mhKts2, 0, mhKts.length);
                    Iterator<String> it = collect.iterator();
                    for (int i = mhKts.length; i < mhKts2.length; i++) {
                        mhKts2[i] = it.next();
                    }
                }
            }
        }
        return headers.add(header);
    }

    public MapHeader getHeader(String headerName, int level) {
        return headers.stream().filter(x -> x.getName().equals(headerName) && x.getLevel() == level)
                .findFirst().orElse(null);
    }

    public Set<MapHeader> getHeaders() {
        return headers;
    }
}
