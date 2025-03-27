package io.github.jjzbruce.excel;

import io.github.jjzbruce.ConvertConfig;

import java.util.*;
import java.util.stream.Collectors;

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
    private List<SheetDataConfig> sheetDataConfigs = new ArrayList<>();

    /**
     * sheet 中数据范围的默认配置
     */
    private SheetDataRangeConfig defaultDataRange = new SheetDataRangeConfig();

    private Class<? extends AbstractExcelMapConverter> delegateImpl = Excel2MapConverterByEvent.class;

    public ExcelConvertConfig(Object source) {
        this.source = source;
    }

    public ExcelConvertConfig(Object source, Class<? extends AbstractExcelMapConverter> delegateImpl) {
        this.source = source;
        if (delegateImpl != null) {
            this.delegateImpl = delegateImpl;
        }
    }

    public ExcelConvertConfig(Object source, SheetDataRangeConfig defaultDataRange) {
        this.source = source;
        if (defaultDataRange != null) {
            this.defaultDataRange = defaultDataRange;
        }
    }

    public SheetDataRangeConfig getDefaultDataRange() {
        return defaultDataRange;
    }

    public Map<Integer, SheetDataConfig> getSheetDataConfigs() {
        return this.sheetDataConfigs.stream().collect(Collectors.toMap(SheetDataConfig::getSheetIndex,
                x -> x, (m1, m2) -> m1));
    }

    public void addSheetDataConfig(SheetDataConfig sheetDataConfig) {
        this.sheetDataConfigs.add(sheetDataConfig);
    }

    public Class<? extends AbstractExcelMapConverter> getDelegateImpl() {
        return delegateImpl;
    }

    public Object getSource() {
        return source;
    }
}
