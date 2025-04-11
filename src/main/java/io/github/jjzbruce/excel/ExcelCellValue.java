package io.github.jjzbruce.excel;

import io.github.jjzbruce.MapKeyType;

/**
 * Excel Cell Value
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.1
 */
public class ExcelCellValue {

    private Object value;

    private MapKeyType mapKeyType;

    public ExcelCellValue(Object value, MapKeyType mapKeyType) {
        this.value = value;
        this.mapKeyType = mapKeyType;
    }

    public Object getValue() {
        return value;
    }

    public String getStringValue() {
        if(value == null) {
            return null;
        }
        return value + "";
    }

    public MapKeyType getMapKeyType() {
        return mapKeyType;
    }

    @Override
    public String toString() {
        return "{" + value + ":" + mapKeyType + '}';
    }
}
