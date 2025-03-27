package io.github.jjzbruce.excel;

import io.github.jjzbruce.MapConverter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * AbstractExcelMapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public abstract class AbstractExcelMapConverter implements MapConverter<ExcelConvertConfig> {

    private Logger log = LoggerFactory.getLogger(AbstractExcelMapConverter.class);

    /**
     * 配置
     */
    protected ExcelConvertConfig config;

    /**
     * 当前sheet下标
     */
    protected Integer currentSheetIndex;

    /**
     * 标题信息
     */
    protected ExcelHead excelHead;

    /**
     * 是否存在分组
     */
    protected boolean existGroup;

    /**
     * 分组
     */
    protected ExcelGroup excelGroup;

    protected AbstractExcelMapConverter(ExcelConvertConfig config) {
        this.config = config;
    }

    protected void init(Integer currentSheetIndex) {
        this.currentSheetIndex = currentSheetIndex;
        SheetDataConfig sheetDataConfig = config.getSheetDataConfigs().get(this.currentSheetIndex);
        SheetDataRangeConfig sheetDataRange = sheetDataConfig.getSheetDataRange();
        this.excelHead = new ExcelHead(sheetDataRange.getHeadRowStart(), sheetDataRange.getHeadRowEnd());
        this.excelGroup = new ExcelGroup(sheetDataRange.getDataRowStart(),
                sheetDataRange.getGroupColumnStart(), sheetDataRange.getGroupColumnEnd());
        this.existGroup = sheetDataRange.getGroupColumnEnd() > sheetDataRange.getGroupColumnStart();
    }

    protected void fillData(SheetDataRangeConfig sheetDataRange, int rowNum, int colNum, Object value,
                            Map<String, Object> map, Consumer<Object> afterSetMapData) {
        if (sheetDataRange == null) {
            return;
        }
        int headRowStart = sheetDataRange.getHeadRowStart();
        int headRowEnd = sheetDataRange.getHeadRowEnd();

        int groupColumnStart = sheetDataRange.getGroupColumnStart();
        int groupColumnEnd = sheetDataRange.getGroupColumnEnd();

        int dataRowStart = sheetDataRange.getDataRowStart();
        int dataRowEnd = sheetDataRange.getDataRowEnd();
        int dataColumnStart = sheetDataRange.getDataColumnStart();
        int dataColumnEnd = sheetDataRange.getDataColumnEnd();

        if (colNum >= dataColumnStart && colNum < dataColumnEnd) {
            if (rowNum >= headRowStart && rowNum < headRowEnd) {
                // 添加到标题
                setHeadTitle(rowNum, colNum, String.valueOf(value));
            } else if (rowNum >= dataRowStart && rowNum < dataRowEnd) {
                // 填充数据
                setMapData(colNum, value, map, afterSetMapData);
            }
        } else if (rowNum >= dataRowStart && colNum >= groupColumnStart && colNum < groupColumnEnd) {
            // 命中分组和设置分组信息
            excelGroup.setGroups(rowNum, colNum, String.valueOf(value));
        }
    }

    /**
     * 若存在分组信息的情况下设置分组信息
     */
    protected Object setGroupIfExist(List<Map<String, Object>> mapList) {
        if (this.existGroup) {

            Map<String, Object> groupMap = new LinkedHashMap<>();
            for (int i = 0; i < mapList.size(); i++) {
                int offsetRowNum = i + excelGroup.getBeginRowNum();
                String[] groups = excelGroup.getGroups(offsetRowNum);
                if(log.isTraceEnabled()) {
                    log.trace("分组信息: {}", Arrays.stream(groups).collect(Collectors.joining(",")));
                }
                Map<String, Object> tmp ;
                if(groupMap.containsKey(groups[0])) {
                    tmp = (Map<String, Object>) groupMap.get(groups[0]);
                } else {
                    tmp = new LinkedHashMap<>();
                    groupMap.put(groups[0], tmp);
                }
                for (int j = 1; j < groups.length; j++) {
                    String group = groups[j];
                    if (j < groups.length - 1) {
                        Map<String, Object> childMap = (Map<String, Object>) tmp.get(group);
                        if (childMap == null) {
                            childMap = new LinkedHashMap<>();
                            tmp.put(group, childMap);
                        }
                        tmp = childMap;
                    } else {
                        tmp.put(group, mapList.get(i));
                    }
                }
            }
            return groupMap;
        } else {
            return mapList;
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

    protected void setHeadTitle(int rowNum, int colNum, String value) {
        if (value == null) {
            return;
        }
        excelHead.setHeads(rowNum, colNum, value);
    }

    @Override
    public ExcelConvertConfig getConvertConfig() {
        return config;
    }
}
