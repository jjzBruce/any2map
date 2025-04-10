package io.github.jjzbruce.excel;

import io.github.jjzbruce.DataMapWrapper;
import io.github.jjzbruce.MapKeyType;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.cont.ContinuableRecord;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;
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

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Xlsx To Map，通过事件读取
 * <pre>
 * XSSF
 * XSSF and SAX (Event API)
 * https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 *
 * HSSF
 * Event API (HSSF Only)
 * https://poi.apache.org/components/spreadsheet/how-to.html#event_api
 * </pre>
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Excel2MapConverterByEvent extends AbstractExcelMapConverter {
    private Logger log = LoggerFactory.getLogger(Excel2MapConverterByEvent.class);

    public Excel2MapConverterByEvent(ExcelConvertConfig config) {
        super(config);
    }

    /**
     * toMapData
     * @since 1.0.1
     * @return DataMapWrapper
     */
    @Override
    public DataMapWrapper toMapData() {
        Map<String, Object> data;
        Object source = config.getSource();
        Objects.nonNull(source);
        String filePath = (String) source;
        if (filePath.endsWith(".xlsx")) {
            data = xlsx2Map(filePath);
        } else if (filePath.endsWith(".xls")) {
            data = xls2Map(filePath);
        } else {
            throw new UnsupportedOperationException("不支持的文件格式: " + filePath);
        }
        return new DataMapWrapper(this.excelHead, data);
    }

    private Map<String, Object> xls2Map(String filePath) {
        long start = System.currentTimeMillis();
        Map<String, Object> map = new LinkedHashMap<>();
        try (
                FileInputStream fin = new FileInputStream(filePath);
                POIFSFileSystem poifs = new POIFSFileSystem(fin);
                InputStream din = poifs.createDocumentInputStream("Workbook")) {
            HSSFRequest req = new HSSFRequest();
            Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
            HssfDataListener hssfDataListener = new HssfDataListener(sheetDataConfigs);
            req.addListenerForAllRecords(hssfDataListener);
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.processEvents(req, din);
            Map<Integer, List<ExcelCellValue[]>> listArrayMap = hssfDataListener.getListArrayMap();
            List<String> sheetNames = hssfDataListener.getSheetNames();
            for (Integer index : sheetDataConfigs.keySet()) {
                init(index);
                SheetDataConfig sheetDataConfig = sheetDataConfigs.get(index);
                if (sheetDataConfig != null) {
                    List<Map<String, Object>> mapList = new ArrayList<>();
                    SheetDataRangeConfig sheetDataRange = sheetDataConfig.getSheetDataRange();
                    if (sheetDataRange == null) {
                        sheetDataRange = config.getDefaultDataRange();
                    }
                    List<ExcelCellValue[]> lineArray = listArrayMap.get(index);
                    for (int i = 0; i < lineArray.size(); i++) {
                        Map<String, Object> lineMap = new LinkedHashMap<>();
                        ExcelCellValue[] objects = lineArray.get(i);
                        for (int j = 0; j < objects.length; j++) {
                            fillData(sheetDataRange, i, j, objects[j], lineMap, null);
                        }
                        if (!lineMap.isEmpty()) {
                            mapList.add(lineMap);
                        }
                    }
                    map.put(sheetNames.get(index), setGroupIfExist(mapList));
                }
            }
            if (log.isDebugEnabled()) {
                log.debug("解析Excel Hssf 耗时: [{}] 数据: {}", System.currentTimeMillis() - start, listArrayMap);
            }
        } catch (Throwable e) {
            //TODO 合理解释？
            e.printStackTrace();
        }
        if (log.isDebugEnabled()) {
            log.debug("输出Map数据耗时: {}", System.currentTimeMillis() - start);
        }
        return map;
    }

    private Map<String, Object> xlsx2Map(String filePath) {
        long start = System.currentTimeMillis();
        Map<String, Object> map = new LinkedHashMap<>();
        Map<Integer, SheetDataConfig> sheetDataConfigs = config.getSheetDataConfigs();
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(filePath);
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
                    init(sheetIndex);
                    List<Map<String, Object>> mapList = new ArrayList<>();
                    InputSource sheetSource = new InputSource(sheet);
                    XssfDataHandler handler = new XssfDataHandler(sst, sheetDataConfig);
                    parser.setContentHandler(handler);
                    parser.parse(sheetSource);
                    sheet.close();
                    List<ExcelCellValue[]> listArray = handler.getListArray();
                    for (int i = 0; i < listArray.size(); i++) {
                        SheetDataRangeConfig sheetDataRange = sheetDataConfig.getSheetDataRange();
                        if (sheetDataRange == null) {
                            sheetDataRange = config.getDefaultDataRange();
                        }
                        Map<String, Object> lineMap = new LinkedHashMap<>();
                        ExcelCellValue[] objects = listArray.get(i);
                        for (int j = 0; j < objects.length; j++) {
                            fillData(sheetDataRange, i, j, objects[j], lineMap, null);
                        }
                        if (!lineMap.isEmpty()) {
                            mapList.add(lineMap);
                        }
                    }
                    map.put(sheetName, setGroupIfExist(mapList));
                    if (log.isTraceEnabled()) {
                        log.trace("解析Excel Sheet[{}] 耗时: [{}] 数据: {}", sheetName, System.currentTimeMillis() - sheetStart,
                                handler.getListArray());
                    }
                }
                sheetIndex += 1;
            }

        } catch (Throwable e) {
            //TODO 合理
            e.printStackTrace();
        }
        if (log.isTraceEnabled()) {
            log.trace("输出Map数据耗时: {}", System.currentTimeMillis() - start);
        }
        return map;
    }

    /**
     * <pre>
     *     解析顺序：所有的sheet信息 -> sheet[0]数据 -> sheet[1]数据 -> ...
     * </pre>
     */
    class HssfDataListener implements HSSFListener {
        private Map<Integer, SheetDataConfig> sheetDataConfigs;
        private SSTRecord sstrec;
        private List<String> sheetNames = new ArrayList<>();
        /**
         * 当前的sheet下标，初始值是 -2
         * 通过 {@link BOFRecord}来更新，每次 +1
         * 当 sheet name 加载之前， 该值会更新为 -1。获取所有的 sheet name 之后，在数据加载之前，值会更新为 0
         */
        private int sheetIndex = -2;
        private Map<Integer, List<ExcelCellValue[]>> listArrayMap = new HashMap<>();
        private List<ExcelCellValue[]> listArray = null;
        /**
         * 列长
         */
        private int colLength = 0;

        public List<String> getSheetNames() {
            return sheetNames;
        }

        private String getSheetName() {
            return sheetNames.get(sheetIndex);
        }

        public HssfDataListener(Map<Integer, SheetDataConfig> sheetDataConfigs) {
            this.sheetDataConfigs = sheetDataConfigs;
        }

        public Map<Integer, List<ExcelCellValue[]>> getListArrayMap() {
            return listArrayMap;
        }

        private void init() {
            this.listArray = new ArrayList<>();
            this.colLength = 0;
            listArrayMap.put(sheetIndex, listArray);
        }

        public void processRecord(org.apache.poi.hssf.record.Record record) {
            short sid = record.getSid();
            switch (sid) {
                // the BOFRecord can represent either the beginning of a sheet or the workbook
                // sheet信息 -> sheet[0]数据 -> sheet[1]数据 -> ...
                case BOFRecord.sid:
                    ++sheetIndex;
                    if (sheetIndex >= 0) {
                        init();
                    }
                    if (log.isTraceEnabled()) {
                        log.trace("BOFRecord sheetIndex: {}", sheetIndex);
                    }
                    break;
                case BoundSheetRecord.sid:
                    BoundSheetRecord bsr = (BoundSheetRecord) record;
                    sheetNames.add(bsr.getSheetname());
                    log.trace("hssf event add new sheet, sheet names: {}", sheetNames);
                    break;
                case SSTRecord.sid:
                    sstrec = (SSTRecord) record;
                    break;
            }
            if (sheetDataConfigs.containsKey(sheetIndex)) {
                SheetDataConfig sheetDataConfig = sheetDataConfigs.get(sheetIndex);
                if (record instanceof CellRecord) {
                    CellRecord cr = (CellRecord) record;
                    int rowNum = cr.getRow();
                    int colNum = cr.getColumn();
                    String cellType = null;
                    ExcelCellValue excelCellValue = null;
                    switch (sid) {
                        case BoolErrRecord.sid:
                            BoolErrRecord brr = (BoolErrRecord) record;
                            excelCellValue = new ExcelCellValue( brr.getBooleanValue(), MapKeyType.BOOLEAN);
                            cellType = "BoolErrRecord";
                            break;
                        case FormulaRecord.sid:
                            FormulaRecord fr = (FormulaRecord) record;
                            excelCellValue = new ExcelCellValue(fr.getValue(), MapKeyType.NUMBER);
                            cellType = "FormulaRecord";
                            break;
                        case LabelSSTRecord.sid:
                            LabelSSTRecord lrec = (LabelSSTRecord) record;
                            excelCellValue = new ExcelCellValue(sstrec.getString(lrec.getSSTIndex()).getString(), MapKeyType.STRING);
                            cellType = "LabelSSTRecord";
                            break;
                        case NumberRecord.sid:
                            NumberRecord nr = (NumberRecord) record;
                            excelCellValue = new ExcelCellValue(nr.getValue(), MapKeyType.NUMBER);
                            cellType = "NumberRecord";

                            ExcelDateTypeConfig excelDataType = sheetDataConfig.getExcelDataType(rowNum, colNum);
                            if (excelDataType != null) {
                                try {
                                    double dateValue = Double.parseDouble(excelCellValue.getValue() + "");
                                    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(dateValue);
                                    excelCellValue = new ExcelCellValue(new SimpleDateFormat(excelDataType.getDateFormat()).format(date),
                                            MapKeyType.DATE);
                                } catch (Throwable ignored) {
                                    if (log.isTraceEnabled()) {
                                        log.trace("数据转为时间错误 ({},{}): {}", rowNum, colNum, excelCellValue.getValue());
                                    }
                                }
                            } else if (BoolErrRecord.sid != sid && LabelSSTRecord.sid != sid) {
                                // 尝试当成 数字来处理
                                try {
                                    excelCellValue = new ExcelCellValue(Double.parseDouble(excelCellValue.getValue() + ""), MapKeyType.NUMBER);
                                } catch (Throwable ignored) {
                                    if (log.isTraceEnabled()) {
                                        log.trace("数据转为数字错误 ({},{}): {}", rowNum, colNum, excelCellValue.getValue());
                                    }
                                }
                            }

                            break;
                        case RKRecord.sid:
                            cellType = "RKRecord";
                            break;
                    }
                    if (log.isTraceEnabled()) {
                        log.trace("({},{}) cellType: {}", rowNum, colNum, cellType);
                    }

                    if(excelCellValue == null) {
                        excelCellValue = new ExcelCellValue(null, null);
                    }

                    // 填充值
                    while (rowNum + 1 > listArray.size()) {
                        listArray.add(new ExcelCellValue[colLength]);
                    }
                    ExcelCellValue[] lineValues = listArray.get(rowNum);
                    // 更新列长
                    if (colNum >= colLength) {
                        colLength = colNum + 1;
                    }
                    // 更新数字长度
                    if (lineValues.length < colLength) {
                        lineValues = Arrays.copyOf(lineValues, colLength);
                    }
                    lineValues[colNum] = excelCellValue;
                    listArray.set(rowNum, lineValues);
                } else if (record instanceof StandardRecord) {
                    // 处理合并区域
                    if (sid == MergeCellsRecord.sid) {
                        // 在基本信息解析完之后才触发，根据合并区域对二维表进行数据整理
                        MergeCellsRecord mcr = (MergeCellsRecord) record;
                        for (int i = 0; i < mcr.getNumAreas(); i++) {
                            CellRangeAddress areaAt = mcr.getAreaAt(i);
                            log.debug("sheet下标: {}, 合并区域: {}", sheetIndex, areaAt.formatAsString());
                            ExcelCellValue mergeValue = listArray.get(areaAt.getFirstRow())[areaAt.getFirstColumn()];
                            for (int j = areaAt.getFirstRow(); j <= areaAt.getLastRow(); j++) {
                                for (int k = areaAt.getFirstColumn(); k <= areaAt.getLastColumn(); k++) {
                                    ExcelCellValue[] lines = listArray.get(j);
                                    if (lines.length < colLength) {
                                        lines = Arrays.copyOf(lines, colLength);
                                        listArray.set(i, lines);
                                    }
                                    lines[k] = mergeValue;
                                }
                            }
                        }
                    }
                } else if (record instanceof ContinuableRecord) {
                    if (sid == SSTRecord.sid) {
                        sstrec = (SSTRecord) record;
                    }
                }
            }
        }
    }


    public class XssfDataHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private SheetDataConfig sheetDataConfig;
        private String lastContents;
        private int currentRowNum;
        private int currentColNum;
        private String cellType;
        private List<ExcelCellValue[]> listArray = new LinkedList<>();
        /**
         * 列长
         */
        private int colLength = 0;

        public List<ExcelCellValue[]> getListArray() {
            return listArray;
        }

        public XssfDataHandler(SharedStringsTable sst, SheetDataConfig sheetDataConfig) {
            this.sst = sst;
            this.sheetDataConfig = sheetDataConfig;
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if ("c".equals(name)) {
                // c => cell
                String ref = attributes.getValue("r");
                CellReference cr = new CellReference(ref);
                currentRowNum = cr.getRow();
                currentColNum = cr.getCol();
                cellType = attributes.getValue("t");
                if (log.isTraceEnabled()) {
                    log.trace("({},{}) cellType: {}", currentRowNum, currentColNum, cellType);
                }
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
                ExcelCellValue mergeValue = listArray.get(firstRowNum)[firstColNum];
                for (int i = firstRowNum; i <= lastRowNum; i++) {
                    if(listArray.get(i).length < colLength) {
                        listArray.set(i, Arrays.copyOf(listArray.get(i), colLength));
                    }
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
                ExcelCellValue excelCellValue = null;
                if (cellType != null) {
                    switch (cellType) {
                        case "b":
                            // 布尔值。单元格内的 v 标签的值为 0 或 1，分别代表 false 和 true
                            excelCellValue = new ExcelCellValue("1".equals(lastContents), MapKeyType.BOOLEAN);
                            break;
                        case "n":
                            // 数字。可能是整数、小数等。
                            excelCellValue = new ExcelCellValue(Double.valueOf(lastContents), MapKeyType.NUMBER);
                            break;
                        case "s":
                            // 共享字符串（Shared String）。
                            // Excel 会把所有的字符串存储在一个共享字符串表中，每个字符串有一个对应的索引。
                            // 当 cellType 为 s 时，v 标签的值是共享字符串表的索引，你需要通过这个索引从共享字符串表中获取实际的字符串内容。
                            int idx = Integer.parseInt(lastContents);
                            excelCellValue = new ExcelCellValue(sst.getItemAt(idx).getString(), MapKeyType.STRING);
                            break;
                        case "str":
                            // 公式字符串。v 标签的值就是公式计算得到的字符串结果。f 标签是计算公式
                            excelCellValue = new ExcelCellValue(lastContents, MapKeyType.STRING);
                            break;
                        case "e":
                            // 错误值。v 标签的值是错误代码
                            excelCellValue = new ExcelCellValue(lastContents, MapKeyType.STRING);
                            break;
                        case "d":
                            // 日期。v 标签的值是错误代码
                            excelCellValue = new ExcelCellValue(lastContents, MapKeyType.STRING);
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
                            excelCellValue = new ExcelCellValue(new SimpleDateFormat(excelDataType.getDateFormat()).format(date), MapKeyType.DATE);
                        } catch (Throwable ignored) {
                            log.debug("数据转为时间错误 ({},{}): {}", currentRowNum, currentColNum, lastContents);
                        }
                    } else if (!("b").equals(cellType) && !("d").equals(cellType)) {
                        // 尝试当成 数字来处理
                        try {
                            excelCellValue = new ExcelCellValue(Double.parseDouble(lastContents), MapKeyType.NUMBER);
                        } catch (Throwable ignored) {
                            log.debug("数据转为数字错误 ({},{}): {}", currentRowNum, currentColNum, lastContents);
                        }
                    }
                }

                if(excelCellValue == null) {
                    excelCellValue = new ExcelCellValue(null, null);
                }

                // 填充值
                while (currentRowNum + 1 > listArray.size()) {
                    listArray.add(new ExcelCellValue[colLength]);
                }
                ExcelCellValue[] lineValues = listArray.get(currentRowNum);
                // 更新列长
                if (currentColNum >= colLength) {
                    colLength = currentColNum + 1;
                }
                // 更新数字长度
                if (lineValues.length < colLength) {
                    lineValues = Arrays.copyOf(lineValues, colLength);
                }
                lineValues[currentColNum] = excelCellValue;
                listArray.set(currentRowNum, lineValues);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }
    }

}
