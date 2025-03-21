//package com.modern.tools.xlsx;
//
//import com.modern.tools.MapConverter;
//import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.util.CellReference;
//import org.apache.poi.util.XMLHelper;
//import org.apache.poi.xssf.eventusermodel.XSSFReader;
//import org.apache.poi.xssf.model.SharedStringsTable;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//import org.xml.sax.*;
//import org.xml.sax.helpers.DefaultHandler;
//
//import javax.xml.parsers.ParserConfigurationException;
//import java.io.InputStream;
//import java.text.SimpleDateFormat;
//import java.util.*;
//import java.util.function.BiPredicate;
//
///**
// * Xlsx To Map
// * 基于：XSSF and SAX (Event API)
// * 参考：https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
// *
// * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
// * @since 1.0.0
// */
//public class Xlsx2MapConverterBySax implements MapConverter<XlsxConvertConfig> {
//
//    private Logger log = LoggerFactory.getLogger(Xlsx2MapConverterBySax.class);
//
//    private XlsxConvertConfig config = new XlsxConvertConfig();
//
//    /**
//     * 列标题缓存，key:列下标
//     */
//    private Map<Integer, String> headValueCache = new HashMap<>();
//
//    public XlsxConvertConfig getConfig() {
//        return config;
//    }
//
//    /**
//     * 转换设置
//     */
//    @Override
//    public void setConvertConfig(XlsxConvertConfig config) {
//        this.config = config;
//    }
//
//    /**
//     * 输出目标 Map
//     *
//     * @return Map
//     */
//    @Override
//    public Map<String, Object> toMap(Object source) {
//        long start = System.currentTimeMillis();
//        Objects.nonNull(source);
//
//        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
//
//        OPCPackage pkg;
//        try {
//            pkg = OPCPackage.open(source + "");
//            XSSFReader r = new XSSFReader(pkg);
//            // 抓取跨行跨列的信息
//            XMLReader parser1 = XMLHelper.newXMLReader();
//            ScanCellRangeAddressHandler scanCellRangeAddressHandler = new ScanCellRangeAddressHandler();
//            parser1.setContentHandler(scanCellRangeAddressHandler);
////            SharedStringsTable sst = (SharedStringsTable) r.getSharedStringsTable();
////            XMLReader parser = fetchSheetParser(sst);
//            Iterator<InputStream> sheets = r.getSheetsData();
//            int sheetIndex = 0;
//            while (sheets.hasNext()) {
//                InputStream sheet = sheets.next();
//                if (sheetDataConfigs.containsKey(sheetIndex++)) {
//                    InputSource sheetSource = new InputSource(sheet);
//                    parser1.parse(sheetSource);
//                    System.out.println(scanCellRangeAddressHandler.getMergedRegions());
//                    sheet.close();
//                }
//            }
//
//        } catch (Throwable e) {
//            //TODO 合理
//            e.printStackTrace();
//        }
//        if (log.isDebugEnabled()) {
//            log.debug("输出Map数据耗时: {}", System.currentTimeMillis() - start);
//        }
//        return null;
//    }
//
//    public void convertSheetData(Sheet sheet, FormulaEvaluator evaluator, Integer headRowStart,
//                                 Integer dataRowStart, Integer dataRowEnd, Integer dataColumnStart, Integer dataColumnEnd,
//                                 List<Map<String, Object>> mapList, BiPredicate<Object, Object>... skipRowTest) {
//        long start = System.currentTimeMillis();
//        // 列标题缓存，key:列下标
//        Map<Integer, String> headValueCache = new HashMap<>();
//        // 最小的有效函数
//        int rowStartNum = Math.min(headRowStart, dataRowStart);
//        // 缓存跨行夸列的信息，key：坐标(x,y)
//        Map<String, Object> cellMergedValueCache = new HashMap<>();
//        // 提取数据
//        row:
//        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
//            Row row = sheet.getRow(rowNum);
//            if (rowNum < rowStartNum) {
//                continue;
//            }
//            // 匹配列标题
//            if (headRowStart == rowNum) {
//                Integer lastCellNum;
//                if (dataColumnEnd == null) {
//                    lastCellNum = Short.valueOf(row.getLastCellNum()).intValue();
//                } else {
//                    lastCellNum = dataColumnEnd;
//                }
//                for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
//                    Cell cell = row.getCell(cellNum);
//                    int j = cell.getColumnIndex();
//                    if (j >= lastCellNum) {
//                        break;
//                    }
//                    if (j >= dataColumnStart) {
//                        setHeadTitle(j, getCellString(evaluator, cell), headValueCache);
//                    }
//                }
//                continue;
//            }
//            if (dataRowEnd != null && row.getRowNum() >= dataRowEnd) {
//                break;
//            }
//            if (row.getRowNum() >= dataRowStart) {
//                Map<String, Object> map = new LinkedHashMap<>();
//                //遍历所有的列
//                for (Cell cell : row) {
//                    int i = cell.getRowIndex(), j = cell.getColumnIndex();
//                    if (j < dataColumnStart || (j - dataColumnStart) >= headValueCache.size()) {
//                        continue;
//                    }
//
//                    String xy = i + "," + j;
//                    Object cellValue;
//                    if (cellMergedValueCache.containsKey(xy)) {
//                        cellValue = cellMergedValueCache.get(xy);
//                    } else {
//                        cellValue = getCellValue(evaluator, cell);
//                        CellRangeAddress cellMerged = getCellMerged(sheet, cell);
//                        if (cellMerged != null) {
//                            for (int k = cellMerged.getFirstRow(); k <= cellMerged.getLastRow(); k++) {
//                                for (int l = cellMerged.getFirstColumn(); l <= cellMerged.getLastColumn(); l++) {
//                                    String xxyy = k + "," + l;
//                                    cellMergedValueCache.put(xxyy, cellValue);
//                                }
//                            }
//                        }
//                    }
//                    String head = headValueCache.get(j);
//                    // 不处理的情况
//                    if (skipRowTest != null) {
//                        for (BiPredicate<Object, Object> test : skipRowTest) {
//                            if (test.test(head, cellValue)) {
//                                continue row;
//                            }
//                        }
//                    }
//                    map.put(head, cellValue);
//                }
//                if (map.size() >= headValueCache.size() - 1) {
//                    mapList.add(map);
//                }
//            }
//        }
//        if (log.isDebugEnabled()) {
//            log.debug("处理sheet[{}] 耗时：{}", sheet.getSheetName(), System.currentTimeMillis() - start);
//        }
//    }
//
//    /**
//     * 维护标题缓存（需要按照顺序访问excel数据次方法才能有效）
//     *
//     * @param headValueCache 标题缓存，key: 列下标
//     */
//    private void setHeadTitle(int columnIndex, String cellValue, Map<Integer, String> headValueCache) {
//        if (cellValue == null) {
//            return;
//        }
//        headValueCache.put(columnIndex, cellValue);
//        // 处理跨列的标题逻辑
//        // 维护扩列标题，一般 [1, 1]: 标题1，[1, 4]: 标题2；那么[1, 2], [1, 3] 的标题为标题1
//        int jj = columnIndex - 1;
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
//
//    private String getCellString(FormulaEvaluator evaluator, Cell cell) {
//        Object cellValue = getCellValue(evaluator, cell);
//        if (cellValue == null) {
//            return null;
//        }
//        if (cellValue instanceof Date) {
//            return new SimpleDateFormat("yyyy-MM-dd").format(cellValue);
//        } else {
//            return cellValue.toString();
//        }
//    }
//
//    private Object getCellValue(FormulaEvaluator evaluator, Cell cell) {
//        if (cell == null) {
//            return null;
//        }
//        switch (cell.getCellType()) {
//            case STRING:
//                return cell.getStringCellValue();
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    return cell.getDateCellValue();
//                } else {
//                    return cell.getNumericCellValue();
//                }
//            case BOOLEAN:
//                return cell.getBooleanCellValue();
//            case FORMULA:
//                CellValue evalVal = evaluator.evaluate(cell);
//                switch (evalVal.getCellType()) {
//                    case NUMERIC:
//                        return evalVal.getNumberValue();
//                    case STRING:
//                        return evalVal.getStringValue();
//                    case BOOLEAN:
//                        return evalVal.getBooleanValue();
//                    default:
//                        return null;
//                }
////                return cell.getStringCellValue();
//            case _NONE:
//                return null;
//            case BLANK:
//                return "";
//            default:
//                log.error("无法解析的Cell，坐标: ({}, {})， 类型: {}", cell.getRowIndex(), cell.getColumnIndex(), cell.getCellType());
//                return null;
//        }
//    }
//
//    /**
//     * 获取单元格的合并信息
//     *
//     * @param sheet 工作表
//     * @param cell  单元格
//     * @return 如果单元格位于合并区域内合并信息
//     */
//    public CellRangeAddress getCellMerged(Sheet sheet, Cell cell) {
//        int numMergedRegions = sheet.getNumMergedRegions();
//        for (int i = 0; i < numMergedRegions; i++) {
//            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
//            if (mergedRegion.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
//                return mergedRegion;
//            }
//        }
//        return null;
//    }
//
//    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
//        XMLReader parser = XMLHelper.newXMLReader();
//        ContentHandler handler = new SheetHandler(sst);
//        parser.setContentHandler(handler);
//        return parser;
//    }
//
//    public class SheetHandler extends DefaultHandler {
//        private SharedStringsTable sst;
//        private String lastContents;
//        private boolean nextIsString;
//
//        public SheetHandler(SharedStringsTable sst) {
//            this.sst = sst;
//        }
//
//        @Override
//        public void startElement(String uri, String localName, String name,
//                                 Attributes attributes) throws SAXException {
//            if (name.equals("c")) {
//                // c => cell
//                String cellRef = attributes.getValue("r");
//
//                // Print the cell reference
//                System.out.print(attributes.getValue("r") + " - ");
//                // Figure out if the value is an index in the SST
//                String cellType = attributes.getValue("t");
//                if (cellType != null && cellType.equals("s")) {
//                    nextIsString = true;
//                } else {
//                    nextIsString = false;
//                }
//            }
//            // Clear contents cache
//            lastContents = "";
//        }
//
//        @Override
//        public void endElement(String uri, String localName, String name) throws SAXException {
//            // Process the last contents as required.
//            // Do now, as characters() may be called more than once
//            if (nextIsString) {
//                int idx = Integer.parseInt(lastContents);
//                lastContents = sst.getItemAt(idx).getString();
//                nextIsString = false;
//            }
//            // v => contents of a cell
//            // Output after we've seen the string contents
//            if (name.equals("v")) {
//                System.out.println(lastContents);
//            }
//        }
//
//        @Override
//        public void characters(char[] ch, int start, int length) {
//            lastContents += new String(ch, start, length);
//        }
//
//        private CellRangeAddress parseCellRange(String ref) {
//            // 解析类似 "A1:B2" 的引用为合并区域
//            String[] parts = ref.split(":");
//            if (parts.length != 2) return null;
//            return new CellRangeAddress(
//                    CellReference.convertColStringToIndex(parts[0].split("\\d+")[0]),
//                    Integer.parseInt(parts[0].replaceAll("[^\\d]", "")),
//                    CellReference.convertColStringToIndex(parts[1].split("\\d+")[0]),
//                    Integer.parseInt(parts[1].replaceAll("[^\\d]", ""))
//            );
//        }
//
//        private CellRangeAddress getMergedRegion(String cellRef) {
//            // 将单元格引用转换为行列索引
//            CellReference cr = new CellReference(cellRef);
//            int row = cr.getRow();
//            int col = cr.getCol();
//            // 查找对应的合并区域
//            return mergedRegions.stream()
//                    .filter(r -> r.isInRange(row, col))
//                    .findFirst()
//                    .orElse(null);
//        }
//
//    }
//
//    private CellRangeAddress getMergedRegion(List<CellRangeAddress> mergedRegions, String cellRef) {
//        // 将单元格引用转换为行列索引
//        CellReference cr = new CellReference(cellRef);
//        int row = cr.getRow();
//        int col = cr.getCol();
//        // 查找对应的合并区域
//        return mergedRegions.stream()
//                .filter(r -> r.isInRange(row, col))
//                .findFirst()
//                .orElse(null);
//    }
//
//    /**
//     * 将单元格引用转换为行列索引
//     *
//     * @param ref 单元格引用
//     */
//    private CellRangeAddress parseCellRange(String ref) {
//        String[] parts = ref.split(":");
//        if (parts.length != 2) return null;
//        int firstRow = Integer.parseInt(parts[0].replaceAll("[^\\d]", ""));
//        int lastRow = Integer.parseInt(parts[1].replaceAll("[^\\d]", ""));
//        int firstCol = CellReference.convertColStringToIndex(parts[0].split("\\d+")[0]);
//        int lastCol = CellReference.convertColStringToIndex(parts[1].split("\\d+")[0]);
//        return new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
//    }
//
//    /**
//     * 处理跨行跨列信息
//     */
//    private class ScanCellRangeAddressHandler extends DefaultHandler {
//        private final List<CellRangeAddress> mergedRegions = new ArrayList<>();
//
//        public List<CellRangeAddress> getMergedRegions() {
//            return mergedRegions;
//        }
//
//        @Override
//        public void startElement(String uri, String localName, String name,
//                                 Attributes attributes) throws SAXException {
//            if ("mergeCell".equals(name)) {
//                String currentRef = attributes.getValue("ref");
//                String firstRef = currentRef.split(":")[0];
//                CellRangeAddress cellRangeAddress = getMergedRegion(mergedRegions, firstRef);
//                if (cellRangeAddress == null) {
//                    cellRangeAddress = parseCellRange(currentRef);
//                    mergedRegions.add(cellRangeAddress);
//                }
//            }
//        }
//    }
//}
