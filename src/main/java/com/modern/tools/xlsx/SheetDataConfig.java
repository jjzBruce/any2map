package com.modern.tools.xlsx;

/**
 * SheetDataConfig
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class SheetDataConfig {

    private Integer sheetIndex = 0;
    private SheetDataRange sheetDataRange;

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public SheetDataRange getSheetDataRange() {
        return sheetDataRange;
    }

    public void setSheetDataRange(SheetDataRange sheetDataRange) {
        this.sheetDataRange = sheetDataRange;
    }

}


