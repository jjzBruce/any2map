package com.modern.tools.excel;

import com.modern.tools.ConvertConfig;

import java.util.*;

/**
 * Excel Convert Config
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class ExcelConvertConfig implements ConvertConfig {
    /**
     * 数据源
     */
    private Object source;

    /**
     * sheet 数据配置
     */
    private Map<Integer, SheetDataConfig> sheetDataConfigs = new TreeMap<>(Comparator.comparingInt(x -> x));

    /**
     * sheet 中数据范围的默认配置
     */
    private SheetDataRange defaultDataRange = new SheetDataRange();

    private Class<? extends AbstractExcelMapConverter> delegateImpl = Excel2MapConverterBySax.class;

    public ExcelConvertConfig(Object source) {
        this.source = source;
    }

    public ExcelConvertConfig(Object source, Class<? extends AbstractExcelMapConverter> delegateImpl) {
        this.source = source;
        if(delegateImpl != null) {
            this.delegateImpl = delegateImpl;
        }
    }

    public ExcelConvertConfig(Object source, SheetDataRange defaultDataRange) {
        this.source = source;
        if(defaultDataRange != null) {
            this.defaultDataRange = defaultDataRange;
        }
    }

    public SheetDataRange getDefaultDataRange() {
        return defaultDataRange;
    }

    public Map<Integer, SheetDataConfig> getSheetDataConfigs() {
        return sheetDataConfigs;
    }

    public void addSheetDataConfig(SheetDataConfig sheetDataConfig) {
        this.sheetDataConfigs.put(sheetDataConfig.getSheetIndex(), sheetDataConfig);
    }

    public Class<? extends AbstractExcelMapConverter> getDelegateImpl() {
        return delegateImpl;
    }

    public Object getSource() {
        return source;
    }
}
