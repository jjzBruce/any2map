package com.modern.tools.xlsx;

import com.modern.tools.ConvertConfig;

import java.util.*;

/**
 * Xlsx Convert Config
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class XlsxConvertConfig implements ConvertConfig {

    private Map<Integer, SheetDataConfig> sheetDataConfigs = new TreeMap<>(Comparator.comparingInt(x -> x));

    private SheetDataRange defaultDataRange = new SheetDataRange();

    public SheetDataRange getDefaultDataRange() {
        return defaultDataRange;
    }

    public void setDefaultDataRange(SheetDataRange defaultDataRange) {
        this.defaultDataRange = defaultDataRange;
    }

    public Map<Integer, SheetDataConfig> getSheetDataConfigs() {
        return sheetDataConfigs;
    }

    public void addSheetDataConfig(SheetDataConfig sheetDataConfig) {
        this.sheetDataConfigs.put(sheetDataConfig.getSheetIndex(), sheetDataConfig);
    }

    public void setSheetDataConfigs(Map<Integer, SheetDataConfig> sheetDataConfigs) {
        this.sheetDataConfigs = sheetDataConfigs;
    }


}
