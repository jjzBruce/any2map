package io.github.jjzbruce.excel;

import io.github.jjzbruce.DataMapWrapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.BiPredicate;

import static io.github.jjzbruce.MapKeyType.*;

/**
 * Xlsx To Map
 * 非默认转换器，谨慎使用！
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Excel2MapConverter extends AbstractExcelMapConverter {
    private Logger log = LoggerFactory.getLogger(Excel2MapConverter.class);

    public Excel2MapConverter(ExcelConvertConfig config) {
        super(config);
    }

    /**
     * toMapData
     * @since 1.0.1
     * @return DataMapWrapper
     */
    @Override
    public DataMapWrapper toMapData() {
        Object source = config.getSource();
        Objects.nonNull(source);
        long start = System.currentTimeMillis();
        InputStream is;
        if (source instanceof String) {
            try {
                is = new FileInputStream(String.valueOf(source));
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
        } else {
            throw new UnsupportedOperationException("需要传入文件路径");
        }
        Map<String, Object> map = new LinkedHashMap<>();
        Workbook workbook;
        long create;
        try {
            workbook = WorkbookFactory.create(is);
            create = System.currentTimeMillis();
            if (log.isTraceEnabled()) {
                log.trace("创建Workbook耗时: {}", create - start);
            }
        } catch (Throwable e) {
            throw new IllegalArgumentException("不支持的文件格式");
        }

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();

        long prepare = System.currentTimeMillis();
        if (log.isTraceEnabled()) {
            log.trace("输出Map准备阶段耗时: {}", prepare - create);
        }

        for (Integer sheetNo : sheetDataConfigs.keySet()) {
            init(sheetNo);
            List<Map<String, Object>> listMap = new ArrayList<>();
            SheetDataConfig sheetDataConfig = sheetDataConfigs.get(sheetNo);
            SheetDataRangeConfig sheetDataRange = sheetDataConfig.getSheetDataRange();
            if (sheetDataRange == null) {
                sheetDataRange = config.getDefaultDataRange();
            }
            Sheet sheet = workbook.getSheetAt(sheetNo);
            convertSheetData(sheet, sheetDataRange, evaluator,
                    sheetDataRange.getHeadRowStart(),
                    sheetDataRange.getDataRowStart(), sheetDataRange.getDataRowEnd(), sheetDataRange.getDataColumnStart(), sheetDataRange.getDataColumnEnd(),
                    listMap
            );
            map.put(sheet.getSheetName(), listMap);
        }

        if (log.isTraceEnabled()) {
            log.trace("输出Map数据耗时: {}", System.currentTimeMillis() - prepare);
        }
        return new DataMapWrapper(this.excelHead, map);
    }

    public void convertSheetData(Sheet sheet, SheetDataRangeConfig sheetDataRange, FormulaEvaluator evaluator, Integer headRowStart,
                                 Integer dataRowStart, Integer dataRowEnd, Integer dataColumnStart, Integer dataColumnEnd,
                                 List<Map<String, Object>> mapList, BiPredicate<Object, Object>... skipRowTest) {
        long start = System.currentTimeMillis();
        // 列标题缓存，key:列下标
        Map<Integer, String> headValueCache = new HashMap<>();
        // 最小的有效函数
        int rowStartNum = Math.min(headRowStart, dataRowStart);
        // 缓存跨行夸列的信息，key：坐标(x,y)
        Map<String, ExcelCellValue> cellMergedValueCache = new HashMap<>();

        // 提取数据
        row:
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (rowNum < rowStartNum) {
                continue;
            }
            // 最后一个列下标（不包含）
            int maxCellNum = Math.min(row.getLastCellNum(), sheetDataRange.getDataColumnEnd());
            // 匹配列标题
            if (headRowStart == rowNum) {
                for (int colNum = 0; colNum < maxCellNum; colNum++) {
                    Cell cell = row.getCell(colNum);
                    fillData(sheetDataRange, rowNum, colNum, getCellString(evaluator, cell),
                            null, null);
                }
            }

            else if (row.getRowNum() >= dataRowStart) {
                Map<String, Object> map = new LinkedHashMap<>();
                //遍历所有的列
                for (int colNum = 0; colNum < maxCellNum; colNum++) {
                    Cell cell = row.getCell(colNum);
                    String xy = rowNum + "," + colNum;
                    ExcelCellValue cellValue;
                    if (cellMergedValueCache.containsKey(xy)) {
                        cellValue = cellMergedValueCache.get(xy);
                    } else {
                        cellValue = getCellValue(evaluator, cell);
                    }

                    String head = headValueCache.get(colNum);
                    if (skipRowTest != null) {
                        for (BiPredicate<Object, Object> test : skipRowTest) {
                            if (test.test(head, cellValue)) {
                                continue row;
                            }
                        }
                    }

                    fillData(sheetDataRange, rowNum, colNum, cellValue,
                            map, cv -> {
                                CellRangeAddress cellMerged = getCellMerged(sheet, cell);
                                if (cellMerged != null) {
                                    for (int k = cellMerged.getFirstRow(); k <= cellMerged.getLastRow(); k++) {
                                        for (int l = cellMerged.getFirstColumn(); l <= cellMerged.getLastColumn(); l++) {
                                            String xxyy = k + "," + l;
                                            cellMergedValueCache.put(xxyy, cv);
                                        }
                                    }
                                }
                            });
                }
                if (map.size() >= headValueCache.size() - 1) {
                    mapList.add(map);
                }
            }
        }
        if (log.isTraceEnabled()) {
            log.trace("处理sheet[{}] 耗时：{}", sheet.getSheetName(), System.currentTimeMillis() - start);
        }
    }

    private ExcelCellValue getCellString(FormulaEvaluator evaluator, Cell cell) {
        Object cellValue = getCellValue(evaluator, cell);
        if (cellValue == null) {
            return null;
        }
        if (cellValue instanceof Date) {
            new ExcelCellValue(new SimpleDateFormat("yyyy-MM-dd").format(cellValue), STRING);
        } else {
            new ExcelCellValue(cellValue.toString(), STRING);
        }
        return null;
    }

    private ExcelCellValue getCellValue(FormulaEvaluator evaluator, Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return new ExcelCellValue(cell.getStringCellValue(), STRING);
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new ExcelCellValue(cell.getDateCellValue(), DATE);
                } else {
                    return new ExcelCellValue(cell.getNumericCellValue(), NUMBER);
                }
            case BOOLEAN:
                return new ExcelCellValue(cell.getBooleanCellValue(), BOOLEAN);
            case FORMULA:
                CellValue evalVal = evaluator.evaluate(cell);
                switch (evalVal.getCellType()) {
                    case NUMERIC:
                        return new ExcelCellValue(evalVal.getNumberValue(), NUMBER);
                    case STRING:
                        return new ExcelCellValue(evalVal.getStringValue(), STRING);
                    case BOOLEAN:
                        return new ExcelCellValue(evalVal.getBooleanValue(), BOOLEAN);
                    default:
                        return null;
                }
            case _NONE:
                return null;
            case BLANK:
                return new ExcelCellValue("", STRING);
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
