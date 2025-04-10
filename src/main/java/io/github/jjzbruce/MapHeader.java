package io.github.jjzbruce;

import java.util.Objects;

/**
 * 数据头
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.1
 */
public class MapHeader {

    /**
     * 数据头
     */
    private String name;

    /**
     * 父键
     */
    private MapHeader parentHeader;

    /**
     * 层级
     */
    private int level;

    /**
     * 是否叶子
     */
    private boolean leaf;

    /**
     * 表示此键映射数据的类型
     */
    private String keyType;

    public MapHeader(String name, MapHeader parentHeader, int level, boolean leaf, String keyType) {
        this.name = name;
        this.parentHeader = parentHeader;
        this.level = level;
        this.leaf = leaf;
        this.keyType = keyType;
    }

    public void setParentHeader(MapHeader parentHeader) {
        this.parentHeader = parentHeader;
    }

    public void setLeaf(boolean leaf) {
        this.leaf = leaf;
    }


    public MapHeader getParentHeader() {
        return parentHeader;
    }

    public int getLevel() {
        return level;
    }

    public boolean isLeaf() {
        return leaf;
    }

    public String getName() {
        return name;
    }

    public String getKeyType() {
        return keyType;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null || getClass() != obj.getClass()) {
            return false;
        }
        MapHeader other = (MapHeader) obj;
        return name.equals(other.name) && level == other.level &&
                (Objects.equals(parentHeader, other.parentHeader));
    }

    @Override
    public int hashCode() {
        int result = name != null ? name.hashCode() : 0;
        result = 31 * result + level;
        result = 31 * result + (parentHeader != null ? parentHeader.hashCode() : 0);
        return result;
    }

    public static MapHeader ofRoot(String header) {
        return new MapHeader(header, null, 0, false, null);
    }

    public static MapHeader ofChild(String header, MapHeader parentHeader) {
        return new MapHeader(header, parentHeader, parentHeader.level + 1, false, null);
    }

    public static MapHeader ofLeaf(String header, MapHeader parentHeader, String keyType) {
        return new MapHeader(header, parentHeader, parentHeader.level + 1, true, keyType);
    }

    public static MapHeader ofLeaf(String header, String keyType) {
        return new MapHeader(header, null, 0, true, keyType);
    }

}
