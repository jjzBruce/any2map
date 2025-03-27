package io.github.jjzbruce.excel;

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

    public boolean isInRange(int row, int col) {
        return row >= firstRowNum && row <= lastRowNum && col >= firstColNum && col <= lastColNum;
    }

}
