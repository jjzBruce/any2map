package com.modern.tools.xlsx;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.util.*;

/**
 * Xlsx To Map，读取两次，第一次获取合并区域信息，第二次整合数据
 * 基于：XSSF and SAX (Event API)
 * 参考：https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverterBySax extends AbstractExcelMapConverter {
    private Logger log = LoggerFactory.getLogger(Xlsx2MapConverterBySax.class);

    /**
     * 输出目标 Map
     *
     * @return Map
     */
    @Override
    public Map<String, Object> toMap(Object source) {
        long start = System.currentTimeMillis();
        Objects.nonNull(source);
        Map<String, Object> map = new LinkedHashMap<>();
        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(source + "");
            XSSFReader xssfReader = new XSSFReader(pkg);
            // 抓取跨行跨列的信息
            XMLReader mergeParser = XMLHelper.newXMLReader();
            ScanCellRangeAddressHandler scanCellRangeAddressHandler = new ScanCellRangeAddressHandler();
            mergeParser.setContentHandler(scanCellRangeAddressHandler);

            SharedStringsTable sst = (SharedStringsTable) xssfReader.getSharedStringsTable();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext()) {
                InputStream sheet = sheets.next();
                String sheetName = sheets.getSheetName();
                SheetDataConfig sheetDataConfig = sheetDataConfigs.get(sheetIndex);
                if (sheetDataConfig != null) {
                    List<Map<String, Object>> mapList = new ArrayList<>();
                    map.put(sheetName, mapList);

                    InputSource sheetSource = new InputSource(sheet);
                    mergeParser.parse(sheetSource);
                    if (log.isDebugEnabled()) {
                        log.debug("sheet[{}]合并区域是:[{}]", sheetIndex, scanCellRangeAddressHandler.getMergedRegions());
                    }
                    XMLReader parser = XMLHelper.newXMLReader();
                    SheetHandler handler = new SheetHandler(sst, sheetDataConfig.getSheetDataRange(),
                            scanCellRangeAddressHandler.getMergedRegions());
                    parser.setContentHandler(handler);
                    parser.parse(sheetSource);
                    sheet.close();
                }
                sheetIndex += 1;
            }

        } catch (Throwable e) {
            //TODO 合理
            e.printStackTrace();
        }
        if (log.isDebugEnabled()) {
            log.debug("输出Map数据耗时: {}", System.currentTimeMillis() - start);
        }
        return map;
    }

    public class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private SheetDataRange sheetDataRange;
        private String lastContents;
        private boolean nextIsString;
        private int currentRowNum;
        private int currentColNum;
        private String currentCellRef;
        private List<Map<String, Object>> mapList;
        /**
         * 合并区域的坐标映射到最左上方坐标
         */
        private final Map<String, String> mergeMap = new HashMap<>();
        /**
         * 记录合并区域左上方对应的值
         */
        private final Map<String, Object> mergeValueMap = new HashMap<>();

        public SheetHandler(SharedStringsTable sst, SheetDataRange sheetDataRange,
                            List<CellRangeAddress> mergedRegions) {
            this.sst = sst;
            this.sheetDataRange = sheetDataRange;
            if (CollectionUtils.isNotEmpty(mergedRegions)) {
                mergedRegions.forEach(cellRangeAddress -> {
                    String value = cellRangeAddress.getFirstRow() + "," + cellRangeAddress.getFirstColumn();
                    for (int i = cellRangeAddress.getFirstRow(); i <= cellRangeAddress.getLastRow(); i++) {
                        for (int j = cellRangeAddress.getFirstColumn(); j <= cellRangeAddress.getLastColumn(); j++) {
                            mergeMap.put(i + "," + j, value);
                        }
                    }
                });
                if(log.isDebugEnabled()) {
                    log.debug("mergeMap: {}", mergeMap);
                }
            }
            this.mapList = new ArrayList<>();
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if (name.equals("c")) {
                // c => cell
                currentCellRef = attributes.getValue("r");
                currentRowNum = CellReference.convertColStringToIndex(currentCellRef.split("\\d+")[0]);
                currentColNum = Integer.parseInt(currentCellRef.replaceAll("[^\\d]", ""));

                // Print the cell reference
                System.out.print(attributes.getValue("r") + " - ");
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
            // Clear contents cache
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
            Map<String, Object> lineMap;
            try {
                lineMap = mapList.get(currentRowNum);
            } catch (Throwable e) {
                lineMap = new LinkedHashMap<>();
                mapList.add(currentRowNum, lineMap);
            }

            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = sst.getItemAt(idx).getString();
                nextIsString = false;
            }

            Object value;
            String ij = currentRowNum + "," + currentColNum;
            if (name.equals("v")) {
                // v => contents of a cell
                value = lastContents;
            } else {
                if (mergeMap.containsKey(ij)) {
                    // 在合并区域的情况下取值
                    value = mergeValueMap.get(mergeMap.get(ij));
                } else {
                    value = null;
                }
            }

            fillData(sheetDataRange, currentRowNum, currentColNum, lastContents, lineMap, v -> {
                if (mergeMap.containsKey(ij) && ij.equals(mergeMap.get(ij))) {
                    // 表示该值是合并区域的左上方的值
                    mergeValueMap.put(ij, value);
                }
            });

            currentRowNum = -1;
            currentColNum = -1;
            currentCellRef = "";
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }
    }

    private CellRangeAddress getMergedRegion(List<CellRangeAddress> mergedRegions, String cellRef) {
        // 将单元格引用转换为行列索引
        CellReference cr = new CellReference(cellRef);
        int row = cr.getRow();
        int col = cr.getCol();
        // 查找对应的合并区域
        return mergedRegions.stream()
                .filter(r -> r.isInRange(row, col))
                .findFirst()
                .orElse(null);
    }

    /**
     * 将单元格引用转换为行列索引
     *
     * @param ref 单元格引用
     */
    private CellRangeAddress parseCellRange(String ref) {
        String[] parts = ref.split(":");
        if (parts.length != 2) return null;
        int firstRow = Integer.parseInt(parts[0].replaceAll("[^\\d]", ""));
        int lastRow = Integer.parseInt(parts[1].replaceAll("[^\\d]", ""));
        int firstCol = CellReference.convertColStringToIndex(parts[0].split("\\d+")[0]);
        int lastCol = CellReference.convertColStringToIndex(parts[1].split("\\d+")[0]);
        return new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
    }

    /**
     * 处理跨行跨列信息
     */
    private class ScanCellRangeAddressHandler extends DefaultHandler {
        private final List<CellRangeAddress> mergedRegions = new ArrayList<>();

        public List<CellRangeAddress> getMergedRegions() {
            return mergedRegions;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if ("mergeCell".equals(name)) {
                String currentRef = attributes.getValue("ref");
                String firstRef = currentRef.split(":")[0];
                CellRangeAddress cellRangeAddress = getMergedRegion(mergedRegions, firstRef);
                if (cellRangeAddress == null) {
                    cellRangeAddress = parseCellRange(currentRef);
                    mergedRegions.add(cellRangeAddress);
                }
            }
        }
    }
}
