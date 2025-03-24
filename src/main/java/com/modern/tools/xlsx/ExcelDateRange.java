package com.modern.tools.xlsx;

import org.apache.poi.ss.util.CellReference;

/**
 * Excel 时间位置
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class ExcelDateRange {

    int firstRowNum;
    int firstColNum;
    int lastRowNum;
    int lastColNum;

    public ExcelDateRange(int rowNum, int colNum) {
        this.firstRowNum = rowNum;
        this.firstColNum = colNum;
        this.lastRowNum = rowNum;
        this.lastColNum = colNum;
    }

    public ExcelDateRange(String startRef, String endRef) {
        CellReference firstCr = new CellReference(startRef);
        this.firstRowNum = firstCr.getRow();
        this.firstColNum = firstCr.getCol();
        CellReference lastCr = new CellReference(endRef);
        this.lastRowNum = lastCr.getRow();
        this.lastColNum = lastCr.getCol();
    }

    public boolean isInRange(int row, int col) {
        return row >= firstRowNum && row <= lastRowNum && col >= firstColNum && col <= lastColNum;
    }

}
