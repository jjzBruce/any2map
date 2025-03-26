package io.github.jjzbruce.excel;

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

    private String dateFormat;

    public ExcelDateTypeConfig(int rowNum, int colNum) {
        this(rowNum, colNum, "yyyy-MM-dd");
    }

    public ExcelDateTypeConfig(int rowNum, int colNum, String dateFormat) {
        ExcelDateRange excelDateRange = new ExcelDateRange(rowNum, colNum);
        ExcelDateTypeRanges.add(excelDateRange);
        this.dateFormat = dateFormat;
    }

    public ExcelDateTypeConfig(int[][] coordinates, String dateFormat) {
        for (int i = 0; i < coordinates.length; i++) {
            ExcelDateRange excelDateRange = new ExcelDateRange(coordinates[i][0], coordinates[i][1]);
            ExcelDateTypeRanges.add(excelDateRange);
        }
        this.dateFormat = dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public boolean isInRange(int row, int col) {
        if (ExcelDateTypeRanges == null) {
            return false;
        }
        for (ExcelDateRange excelDateTypeRange : ExcelDateTypeRanges) {
            if (excelDateTypeRange.isInRange(row, col)) {
                return true;
            }
        }
        return false;
    }

}
