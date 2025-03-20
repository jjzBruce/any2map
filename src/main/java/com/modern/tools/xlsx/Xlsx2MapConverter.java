package com.modern.tools.xlsx;

import com.modern.tools.MapConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.BiPredicate;

/**
 * Xlsx To Map
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverter implements MapConverter<XlsxConvertConfig> {

    private Logger log = LoggerFactory.getLogger(Xlsx2MapConverter.class);

    private XlsxConvertConfig config = new XlsxConvertConfig();

    public XlsxConvertConfig getConfig() {
        return config;
    }

    /**
     * 转换设置
     *
     * @param config 配置
     */
    @Override
    public void setConvertConfig(XlsxConvertConfig config) {
        this.config = config;
    }

    /**
     * 输出目标 Map
     *
     * @return Map
     */
    @Override
    public Map<String, Object> toMap(Object source) {
        InputStream is = null;
        if (source instanceof InputStream) {
            is = (InputStream) source;
        }
        Map<String, Object> map = new LinkedHashMap<>();
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(is);
        } catch (Throwable e) {
            try {
                workbook = new HSSFWorkbook(is);
            } catch (Throwable e1) {
                throw new IllegalArgumentException("不支持的文件格式");
            }
        }
        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
        for (Integer sheetNo : sheetDataConfigs.keySet()) {
            List<Map<String, Object>> listMap = new ArrayList<>();
            SheetDataConfig sheetDataConfig = sheetDataConfigs.get(sheetNo);
            SheetDataRange sheetDataRange = sheetDataConfig.getSheetDataRange();
            if (sheetDataRange == null) {
                sheetDataRange = config.getDefaultDataRange();
            }
            Sheet sheet = workbook.getSheetAt(0);
            convertSheetData(sheet, sheetDataRange.getHeadRowStart(),
                    sheetDataRange.getDataRowStart(), sheetDataRange.getDataRowEnd(), sheetDataRange.getDataColumnStart(), sheetDataRange.getDataColumnEnd(),
                    listMap
            );
            map.put(sheet.getSheetName(), listMap);
        }

        return map;
    }

    public void convertSheetData(Sheet sheet, Integer headRowStart,
                                 Integer dataRowStart, Integer dataRowEnd, Integer dataColumnStart, Integer dataColumnEnd,
                                 List<Map<String, Object>> mapList, BiPredicate<Object, Object>... skipRowTest) {
        long start = System.currentTimeMillis();
        // 列标题缓存，key:列下标
        Map<Integer, String> headValueCache = new HashMap<>();
        // 最小的有效函数
        int rowStartNum = Math.min(headRowStart, dataRowStart);
        // 缓存跨行夸列的信息，key：坐标(x,y)
        Map<String, Object> cellMergedValueCache = new HashMap<>();
        // 提取数据
        row:
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (rowNum < rowStartNum) {
                continue;
            }
            // 匹配列标题
            if (headRowStart == rowNum) {
                Integer lastCellNum;
                if (dataColumnEnd == null) {
                    lastCellNum = Short.valueOf(row.getLastCellNum()).intValue();
                } else {
                    lastCellNum = dataColumnEnd;
                }
                for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    int j = cell.getColumnIndex();
                    if (j >= lastCellNum) {
                        break;
                    }
                    if (j >= dataColumnStart) {
                        setHeadTitle(j, getCellString(cell), headValueCache);
                    }
                }
                continue;
            }
            if (dataRowEnd != null && row.getRowNum() >= dataRowEnd) {
                break;
            }
            if (row.getRowNum() >= dataRowStart) {
                Map<String, Object> map = new LinkedHashMap<>();
                //遍历所有的列
                for (Cell cell : row) {
                    int i = cell.getRowIndex(), j = cell.getColumnIndex();
                    if (j < dataColumnStart || (j - dataColumnStart) >= headValueCache.size()) {
                        continue;
                    }

                    String xy = i + "," + j;
                    Object cellValue;
                    if (cellMergedValueCache.containsKey(xy)) {
                        cellValue = cellMergedValueCache.get(xy);
                    } else {
                        cellValue = getCellValue(cell);
                        CellRangeAddress cellMerged = getCellMerged(sheet, cell);
                        if (cellMerged != null) {
                            for (int k = cellMerged.getFirstRow(); k <= cellMerged.getLastRow(); k++) {
                                for (int l = cellMerged.getFirstColumn(); l <= cellMerged.getLastColumn(); l++) {
                                    String xxyy = k + "," + l;
                                    cellMergedValueCache.put(xxyy, cellValue);
                                }
                            }
                        }
                    }
                    String head = headValueCache.get(j);
                    // 不处理的情况
                    if (skipRowTest != null) {
                        for (BiPredicate<Object, Object> test : skipRowTest) {
                            if (test.test(head, cellValue)) {
                                continue row;
                            }
                        }
                    }
                    map.put(head, cellValue);
                }
                if (map.size() >= headValueCache.size() - 1) {
                    mapList.add(map);
                }
            }
        }
        log.info("处理sheet耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * 维护标题缓存（需要按照顺序访问excel数据次方法才能有效）
     *
     * @param headValueCache 标题缓存，key: 列下标
     */
    private void setHeadTitle(int columnIndex, String cellValue, Map<Integer, String> headValueCache) {
        if (cellValue == null) {
            return;
        }
        headValueCache.put(columnIndex, cellValue);
        // 维护扩列标题，一般 [1, 1]: 标题1，[1, 4]: 标题2；那么[1, 2], [1, 3] 的标题为标题1
        int jj = columnIndex - 1;
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

    private String getCellString(Cell cell) {
        Object cellValue = getCellValue(cell);
        if (cellValue == null) {
            return null;
        }
        if (cellValue instanceof Date) {
            return new SimpleDateFormat("yyyy-MM-dd").format(cellValue);
        } else {
            return cellValue.toString();
        }
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case _NONE:
                return null;
            case BLANK:
                return "";
            default:
                log.error("无法解析的Cell，坐标: ({}, {})， 类型: {}", cell.getRowIndex(), cell.getColumnIndex(), cell.getCellType());
                return null;
        }
    }

    /**
     * 获取单元格的合并信息
     *
     * @param sheet 工作表
     * @param cell  单元格
     * @return 如果单元格位于合并区域内合并信息
     */
    public CellRangeAddress getCellMerged(Sheet sheet, Cell cell) {
        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return mergedRegion;
            }
        }
        return null;
    }
}
