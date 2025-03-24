package com.modern.tools.excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Xlsx To Map，读取一次，将数据整合成一个二维数组，在整理成Map
 * 基于：XSSF and SAX (Event API)
 * 参考：https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Excel2MapConverterBySax extends AbstractExcelMapConverter {
    private Logger log = LoggerFactory.getLogger(Excel2MapConverterBySax.class);

    public Excel2MapConverterBySax(ExcelConvertConfig config) {
        super(config);
    }

    @Override
    public Map<String, Object> toMap() {
        Object source = config.getSource();
        Objects.nonNull(source);
        return toMap1(source);
    }

    private Map<String, Object> toMap1(Object source) {
        long start = System.currentTimeMillis();
        Objects.nonNull(source);
        Map<String, Object> map = new LinkedHashMap<>();
        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(source + "");
            XSSFReader xssfReader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) xssfReader.getSharedStringsTable();
            XMLReader parser = XMLHelper.newXMLReader();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext()) {
                long sheetStart = System.currentTimeMillis();
                InputStream sheet = sheets.next();
                String sheetName = sheets.getSheetName();
                SheetDataConfig sheetDataConfig = sheetDataConfigs.get(sheetIndex);
                if (sheetDataConfig != null) {
                    List<Map<String, Object>> mapList = new ArrayList<>();
                    map.put(sheetName, mapList);
                    InputSource sheetSource = new InputSource(sheet);
                    SheetHandler handler = new SheetHandler(sst, sheetDataConfig);
                    parser.setContentHandler(handler);
                    parser.parse(sheetSource);
                    sheet.close();

                    List<Object[]> listArray = handler.getListArray();
                    for (int i = 0; i < listArray.size(); i++) {
                        SheetDataRange sheetDataRange = sheetDataConfig.getSheetDataRange();
                        if (sheetDataRange == null) {
                            sheetDataRange = config.getDefaultDataRange();
                        }
                        Map<String, Object> lineMap = new LinkedHashMap<>();
                        Object[] objects = listArray.get(i);
                        for (int j = 0; j < objects.length; j++) {
                            fillData(sheetDataRange, i, j, objects[j], lineMap, null);
                        }
                        if (!lineMap.isEmpty()) {
                            mapList.add(lineMap);
                        }
                    }

                    if (log.isDebugEnabled()) {
                        log.debug("解析Excel Sheet[{}] 耗时: [{}] 数据: {}", sheetName, System.currentTimeMillis() - sheetStart,
                                handler.getListArray());
                    }
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
        private SheetDataConfig sheetDataConfig;
        private String lastContents;
        private int currentRowNum;
        private int currentColNum;
        private String cellType;
        private List<Object[]> listArray = new LinkedList<>();
        private List<Object> firstRowList = new LinkedList<>();
        private int maxColNum;

        public List<Object[]> getListArray() {
            return listArray;
        }

        public SheetHandler(SharedStringsTable sst, SheetDataConfig sheetDataConfig) {
            this.sst = sst;
            this.sheetDataConfig = sheetDataConfig;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {

            if ("c".equals(name)) {
                String ref = attributes.getValue("r");
                CellReference cr = new CellReference(ref);
                currentRowNum = cr.getRow();
                currentColNum = cr.getCol();
                // c => cell
                if (currentColNum == 0 && currentRowNum > currentColNum) {
                    // 列从头开始，行好大于列号的时候代表换行了
                    if (currentRowNum == 1) {
                        // 第0行扫描完成，可以获取最大列数
                        maxColNum = firstRowList.size();
                        // 第0行填充到数组中
                        Object[] firstLine = firstRowList.toArray(new Object[firstRowList.size()]);
                        listArray.add(firstLine);
                    }
                    Object[] line = new Object[maxColNum];
                    listArray.add(line);
                }
                cellType = attributes.getValue("t");
                log.debug("({},{}) cellType: {}", currentRowNum, currentColNum, cellType);
            } else if ("mergeCell".equals(name)) {
                // 处理合并区域取值，取左上值即可
                String mergeRef = attributes.getValue("ref");
                String[] split = mergeRef.split(":");
                // 左上角坐标
                String firstRef = split[0];
                CellReference firstCr = new CellReference(firstRef);
                int firstRowNum = firstCr.getRow();
                int firstColNum = firstCr.getCol();
                // 右下角坐标
                String lastRef = split[1];
                CellReference lastCr = new CellReference(lastRef);
                int lastRowNum = lastCr.getRow();
                int lastColNum = lastCr.getCol();

                // 算出当前坐标
                Object mergeValue = listArray.get(firstRowNum)[firstColNum];
                for (int i = firstRowNum; i <= lastRowNum; i++) {
                    for (int j = firstColNum; j <= lastColNum; j++) {
                        listArray.get(i)[j] = mergeValue;
                    }
                }
                if (log.isDebugEnabled()) {
                    log.debug("坐标({},{}) 至 ({},{}) 是合并区域，值为: {}",
                            firstRowNum, firstColNum, lastRowNum, lastColNum, mergeValue);
                }
            }
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
            if (name.equals("v")) {
                Object value = lastContents;
                if (cellType != null) {
                    switch (cellType) {
                        case "b":
                            // 布尔值。单元格内的 v 标签的值为 0 或 1，分别代表 false 和 true
                            value = "1".equals(lastContents);
                            break;
                        case "n":
                            // 数字。可能是整数、小数等。
                            value = Double.valueOf(lastContents);
                            break;
                        case "s":
                            // 共享字符串（Shared String）。
                            // Excel 会把所有的字符串存储在一个共享字符串表中，每个字符串有一个对应的索引。
                            // 当 cellType 为 s 时，v 标签的值是共享字符串表的索引，你需要通过这个索引从共享字符串表中获取实际的字符串内容。
                            int idx = Integer.parseInt(lastContents);
                            value = sst.getItemAt(idx).getString();
                            break;
                        case "str":
                            // 公式字符串。v 标签的值就是公式计算得到的字符串结果。f 标签是计算公式
                            value = lastContents;
                            break;
                        case "e":
                            // 错误值。v 标签的值是错误代码
                            value = lastContents;
                            break;
                        case "d":
                            // 日期。v 标签的值是错误代码
                            value = lastContents;
                            break;
                    }
                }

                if (!("s").equals(cellType)) {
                    // 处理自定义的数据类型
                    ExcelDateTypeConfig excelDataType = sheetDataConfig.getExcelDataType(currentRowNum, currentColNum);
                    if (excelDataType != null) {
                        try {
                            double dateValue = Double.parseDouble(lastContents);
                            Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(dateValue);
                            value = new SimpleDateFormat(excelDataType.getDateFormat()).format(date);
                        } catch (Throwable ignored) {
                            log.error("数据转为时间错误 ({},{}): {}", currentRowNum, currentColNum, lastContents);
                        }
                    } else if(!("b").equals(cellType) && !("d").equals(cellType)) {
                        // 尝试当成 数字来处理
                        try {
                            value = Double.parseDouble(lastContents);
                        } catch (Throwable ignored) {
                            log.error("数据转为数字错误 ({},{}): {}", currentRowNum, currentColNum, lastContents);
                        }
                    }
                }

                // 填充值
                if (currentRowNum == 0) {
                    firstRowList.add(value);
                } else {
                    listArray.get(listArray.size() - 1)[currentColNum] = value;
                }
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }
    }

}
