package io.github.jjzbruce.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * SheetDataConfig
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class SheetDataConfig {

    private Integer sheetIndex = 0;
    private SheetDataRangeConfig sheetDataRange;

    public SheetDataConfig() {
        this(0, new SheetDataRangeConfig());
    }

    public SheetDataConfig(SheetDataRangeConfig sheetDataRange) {
        this(0, sheetDataRange);
    }
    public SheetDataConfig(Integer sheetIndex) {
        this(sheetIndex, new SheetDataRangeConfig());
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public SheetDataConfig(Integer sheetIndex, SheetDataRangeConfig sheetDataRange) {
        this.sheetIndex = sheetIndex;
        this.sheetDataRange = sheetDataRange;
    }

    /**
     * Excel解析数据类型设置
     */
    private List<ExcelDateTypeConfig> excelDateTypeConfigs = new ArrayList<>();

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public SheetDataRangeConfig getSheetDataRange() {
        return sheetDataRange;
    }

    public void addExcelDateTypeConfig(ExcelDateTypeConfig excelDateTypeConfig) {
        this.excelDateTypeConfigs.add(excelDateTypeConfig);
    }

    public ExcelDateTypeConfig getExcelDataType(int row, int col) {
        if(excelDateTypeConfigs == null) {
            return null;
        }
        for(ExcelDateTypeConfig dateTypeConfig : excelDateTypeConfigs) {
            if(dateTypeConfig.isInRange(row, col)) {
                return dateTypeConfig;
            }
        }
        return null;
    }

}


