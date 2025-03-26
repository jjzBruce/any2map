package io.github.jjzbruce.excel;

import io.github.jjzbruce.MapConverter;

import java.util.*;
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

//    /**
//     * 列标题缓存，key:列下标
//     * TODO 需要实现多层下标的情况
//     */
//    protected Map<Integer, String> headValueCache = new HashMap<>();

    /**
     * 当前sheet下标
     */
    protected Integer currentSheetIndex;

    /**
     * 标题信息
     */
    protected ExcelHead excelHead;

    protected AbstractExcelMapConverter(ExcelConvertConfig config) {
        this.config = config;
    }

    public ExcelConvertConfig getConfig() {
        return config;
    }

    protected void init(Integer currentSheetIndex) {
        this.currentSheetIndex = currentSheetIndex;
        SheetDataConfig sheetDataConfig = config.getSheetDataConfigs().get(this.currentSheetIndex);
        SheetDataRange sheetDataRange = sheetDataConfig.getSheetDataRange();
        this.excelHead = new ExcelHead(sheetDataRange.getHeadRowStart(), sheetDataRange.getHeadRowEnd());
    }

    protected void fillData(SheetDataRange sheetDataRange, int rowNum, int colNum, Object value,
                            Map<String, Object> map, Consumer<Object> afterSetMapData) {
        if (sheetDataRange == null) {
            return;
        }
        if (colNum >= sheetDataRange.getDataColumnStart() && colNum < sheetDataRange.getDataColumnEnd()) {
            if (rowNum >= sheetDataRange.getHeadRowStart() && rowNum < sheetDataRange.getHeadRowEnd()) {
                // 添加到标题
                setHeadTitle(rowNum, colNum, String.valueOf(value));
            } else if (rowNum >= sheetDataRange.getDataRowStart() && rowNum < sheetDataRange.getDataRowEnd()) {
                // 填充数据
                setMapData(colNum, value, map, afterSetMapData);
            }
        }
    }

    protected void setMapData(int colNum, Object value, Map<String, Object> lineMap, Consumer<Object> after) {
        String[] heads = excelHead.getHeads(colNum);
        Map<String, Object> tmp = lineMap;
        for (int i = 0; i < heads.length; i++) {
            String head = heads[i];
            if (i < heads.length - 1) {
                Map<String, Object> childMap = (Map<String, Object>) tmp.get(head);
                if (childMap == null) {
                    childMap = new LinkedHashMap<>();
                    tmp.put(head, childMap);
                }
                tmp = childMap;
            } else {
                tmp.put(head, value);
            }
        }
        if (after != null) {
            after.accept(value);
        }
    }

//    protected void setMapData(int colNum, Object value, Map<String, Object> lineMap, Consumer<Object> after) {
//        String head = headValueCache.get(colNum);
//        lineMap.put(head, value);
//        if (after != null) {
//            after.accept(value);
//        }
//    }

    protected void setHeadTitle(int rowNum, int colNum, String value) {
        if (value == null) {
            return;
        }
        excelHead.setHeads(rowNum, colNum, value);
    }


//    /**
//     * 维护标题缓存
//     */
//    protected void setHeadTitle(int colNum, String value) {
//        if (value == null) {
//            return;
//        }
//        headValueCache.put(colNum, value);
//        // TODO 多行标题的时候使用
//        // 维护扩列标题，一般 [1, 1]: 标题1，[1, 4]: 标题2；那么[1, 2], [1, 3] 的标题为标题1
//        int jj = colNum - 1;
//        List<Integer> needs = new LinkedList<>();
//        while (!headValueCache.containsKey(jj) && jj >= 0) {
//            needs.add(jj--);
//        }
//        if (!needs.isEmpty()) {
//            String addTitle = headValueCache.get(jj);
//            if (addTitle != null) {
//                needs.forEach(jjj -> headValueCache.put(jjj, addTitle));
//            }
//        }
//    }

    @Override
    public ExcelConvertConfig getConvertConfig() {
        return config;
    }
}
