package com.modern.tools.xlsx;

import java.util.Objects;

/**
 * SheetDataConfig
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class SheetDataConfig {

    private Integer sheetIndex = 0;
    private String sheetKey;
    private SheetDataRange sheetDataRange;

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getSheetKey() {
        return sheetKey;
    }

    public void setSheetKey(String sheetKey) {
        this.sheetKey = sheetKey;
    }

    public SheetDataRange getSheetDataRange() {
        return sheetDataRange;
    }

    public void setSheetDataRange(SheetDataRange sheetDataRange) {
        this.sheetDataRange = sheetDataRange;
    }

    @Override
    public boolean equals(Object obj) {
        // 自反性检查
        if (this == obj) {
            return true;
        }
        // 检查对象是否为 null 或者类型是否不匹配
        if (obj == null || getClass() != obj.getClass()) {
            return false;
        }
        SheetDataConfig other = (SheetDataConfig) obj;
        // 比较所有有意义的属性，避免空指针异常
        return Objects.equals(sheetIndex, other.sheetIndex);
    }

    @Override
    public int hashCode() {
        return Objects.hash(sheetIndex, sheetKey, sheetDataRange);
    }

    @Override
    public String toString() {
        return "SheetDataConfig{" +
                "sheetIndex=" + sheetIndex +
                ", sheetKey='" + sheetKey + '\'' +
                ", sheetDataRange=" + sheetDataRange +
                '}';
    }
}


