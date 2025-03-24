package io.github.jjzbruce.excel;

import io.github.jjzbruce.MapConverter;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * AbstractExcelMapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public abstract class AbstractExcelMapConverter implements MapConverter<ExcelConvertConfig> {

    /**
     * 配置
     */
    protected ExcelConvertConfig config;

    /**
     * 列标题缓存，key:列下标
     * TODO 需要实现多层下标的情况
     */
    protected Map<Integer, String> headValueCache = new HashMap<>();

    protected AbstractExcelMapConverter(ExcelConvertConfig config) {
        this.config = config;
    }

    public ExcelConvertConfig getConfig() {
        return config;
    }

//    /**
//     * 转换设置
//     */
//    @Override
//    public void setConvertConfig(ExcelConvertConfig config) {
//        this.config = config;
//    }

    protected void fillData(SheetDataRange sheetDataRange, int rowNum, int colNum, Object value,
                            Map<String, Object> map, Consumer<Object> afterSetMapData) {
        if (sheetDataRange == null) {
            return;
        }
        if (colNum >= sheetDataRange.getDataColumnStart() && colNum < sheetDataRange.getDataColumnEnd()) {
            if (rowNum >= sheetDataRange.getHeadRowStart() && rowNum < sheetDataRange.getHeadRowEnd()) {
                // 添加到标题
                setHeadTitle(colNum, String.valueOf(value));
            } else if (rowNum >= sheetDataRange.getDataRowStart() && rowNum < sheetDataRange.getDataRowEnd()) {
                // 填充数据
                setMapData(colNum, value, map, afterSetMapData);
            }
        }
    }

    protected void setMapData(int colNum, Object value, Map<String, Object> lineMap, Consumer<Object> after) {
        String head = headValueCache.get(colNum);
        lineMap.put(head, value);
        if (after != null) {
            after.accept(value);
        }
    }


    /**
     * 维护标题缓存
     */
    protected void setHeadTitle(int colNum, String value) {
        if (value == null) {
            return;
        }
        headValueCache.put(colNum, value);
        // TODO 多行标题的时候使用
        // 维护扩列标题，一般 [1, 1]: 标题1，[1, 4]: 标题2；那么[1, 2], [1, 3] 的标题为标题1
        int jj = colNum - 1;
        List<Integer> needs = new LinkedList<>();
        while (!headValueCache.containsKey(jj) && jj >= 0) {
            needs.add(jj--);
        }
        if (!needs.isEmpty()) {
            String addTitle = headValueCache.get(jj);
            if (addTitle != null) {
                needs.forEach(jjj -> headValueCache.put(jjj, addTitle));
            }
        }
    }

    @Override
    public ExcelConvertConfig getConvertConfig() {
        return config;
    }
}
