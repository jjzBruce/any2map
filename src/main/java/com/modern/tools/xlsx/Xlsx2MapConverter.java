package com.modern.tools.xlsx;

import com.modern.tools.MapConverter;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.WeekFields;
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
    public List<Map<String, Object>> toListMap(Object source) {
        InputStream is = null;
        if (source instanceof InputStream) {
            is = (InputStream) source;
        }

        List<Map<String, Object>> listMap = new ArrayList<>();
        try (Workbook workbook = StreamingReader.builder().rowCacheSize(10 * 10).bufferSize(1024 * 4)
                //打开资源，可以是InputStream或者是File，注意：只能打开.xlsx格式的文件
                .open(is)) {
            int numberOfSheets = workbook.getNumberOfSheets();
            Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
            for (int i = 0; i < numberOfSheets; i++) {
                if (sheetDataConfigs.keySet().contains(i)) {
                    SheetDataConfig sheetDataConfig = sheetDataConfigs.get(i);
                    SheetDataRange sheetDataRange = sheetDataConfig.getSheetDataRange();
                    if (sheetDataRange == null) {
                        sheetDataRange = config.getDefaultDataRange();
                    }
                    Sheet sheet = workbook.getSheetAt(i);
                    convertSheetData(sheet, sheetDataConfig.getSheetKey(),
                            sheetDataRange.getHeadRowStart(),
                            sheetDataRange.getDataRowStart(), sheetDataRange.getDataRowEnd(), sheetDataRange.getDataColumnStart(), sheetDataRange.getDataColumnEnd(),
                            listMap
                    );
                }
            }
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            return null;
        }
        return listMap;
    }

    public void convertSheetData(Sheet sheet, String sheetName,
                                 Integer headRowStart,
                                 Integer dataRowStart, Integer dataRowEnd, Integer dataColumnStart, Integer dataColumnEnd,
                                 List<Map<String, Object>> mapList, BiPredicate<Object, Object>... skipRowTest) {
        long start = System.currentTimeMillis();
        // 列标题缓存，key:列下标
        Map<Integer, String> headValueCache = new HashMap<>();
        // 最小的有效函数
        int minRowIndex = Math.min(headRowStart, dataRowStart);
        // 提取数据
        row:
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < minRowIndex) {
                continue;
            }

            // 匹配列标题
            if (headRowStart == rowNum) {
                Integer lastCellNum;
                if(dataColumnEnd == null) {
                    lastCellNum = Short.valueOf(row.getLastCellNum()).intValue();
                } else {
                    lastCellNum = dataColumnEnd;
                }
                for (Cell cell : row) {
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

                    Object cellValue = getCellValue(cell);
                    String head1 = headValueCache.get(j);

                    // 不处理的情况
                    if (skipRowTest != null) {
                        for (BiPredicate<Object, Object> test : skipRowTest) {
                            if (test.test(head1, cellValue)) {
                                continue row;
                            }
                        }
                    }

                    // 数据齐平处理
                    if (Objects.equals("", cellValue)) {
                        cellValue = 0D;
                    }

                    map.put(head1, cellValue);
                    // 对日期的数据再次存入 周次 与 月 信息
                    if ("日期".equals(head1)) {
                        try {
                            Date date = (Date) cellValue;
                            LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                            map.put("周次", "WK" + localDate.get(WeekFields.of(Locale.getDefault()).weekOfWeekBasedYear()));
                            map.put("月", localDate.getMonth().getValue());
                        } catch (Throwable ignore) {
                            // NOP
                        }
                    }
                }
                if (map.size() >= headValueCache.size() - 1) {
                    map.put("sheet", sheetName);
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
        CellType cellType = cell.getCellType();
        if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                return new SimpleDateFormat("yyyy-MM-dd").format(date);
            } else {
                return cell.getStringCellValue();
            }
        } else {
            return cell.getStringCellValue();
        }
    }

    private Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();

        if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                return cell.getStringCellValue();
            }
        } else {
            return cell.getStringCellValue();
        }
    }

}
