package com.modern.tools.xlsx;

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
import java.util.*;

/**
 * Xlsx To Map，读取一次，将数据整合成一个二维数组，在整理成Map
 * 基于：XSSF and SAX (Event API)
 * 参考：https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverterBySax2 extends AbstractExcelMapConverter {
    private Logger log = LoggerFactory.getLogger(Xlsx2MapConverterBySax2.class);

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
                    SheetHandler handler = new SheetHandler(sst);
                    parser.setContentHandler(handler);
                    parser.parse(sheetSource);
                    sheet.close();
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
        private String lastContents;
        private boolean nextIsString;
        private int currentRowNum;
        private int currentColNum;
        private List<Object[]> listArray = new LinkedList<>();
        private List<Object> firstRowList = new LinkedList<>();
        private int maxColNum;

        public List<Object[]> getListArray() {
            return listArray;
        }

        public SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
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
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
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
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = sst.getItemAt(idx).getString();
                nextIsString = false;
            }
            if (name.equals("v")) {
                Object value = lastContents;
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
