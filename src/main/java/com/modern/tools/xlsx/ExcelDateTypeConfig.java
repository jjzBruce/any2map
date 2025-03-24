package com.modern.tools.xlsx;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel 时间类型设置
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class ExcelDateTypeConfig {

    private List<ExcelDateRange> ExcelDateTypeRanges = new ArrayList<>();

    private String dateFormat = "yyyy-MM-dd";

    public ExcelDateTypeConfig(int rowNum, int colNum) {
        ExcelDateRange excelDateRange = new ExcelDateRange(rowNum, colNum);
        ExcelDateTypeRanges.add(excelDateRange);
    }

    public void setCoordinates(String startRef, String endRef) {
        ExcelDateRange excelDateRange = new ExcelDateRange(startRef, endRef);
        ExcelDateTypeRanges.add(excelDateRange);
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public boolean isInRange(int row, int col) {
        if(ExcelDateTypeRanges == null) {
            return false;
        }
        for (ExcelDateRange excelDateTypeRange : ExcelDateTypeRanges) {
            if(excelDateTypeRange.isInRange(row, col)) {
                return true;
            }
        }
        return false;
    }

}
