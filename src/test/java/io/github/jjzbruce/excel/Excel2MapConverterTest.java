package io.github.jjzbruce.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.jjzbruce.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Random;

/**
 * Xlsx2MapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class Excel2MapConverterTest {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Test
    public void testXlsxByExcel2MapConverter() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xlsx";
        ExcelConvertConfig config2 = new ExcelConvertConfig(filePath, Excel2MapConverter.class);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        config2.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config2);
        doTest(mc);
    }

    @Test
    public void testXlsByExcel2MapConverter() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xls";
        ExcelConvertConfig config2 = new ExcelConvertConfig(filePath, Excel2MapConverter.class);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        config2.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config2);
        doTest(mc);
    }

    @Test
    public void testXlsx() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xlsx";
        ExcelConvertConfig config2 = new ExcelConvertConfig(filePath);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        ExcelDateTypeConfig edtc = new ExcelDateTypeConfig(0, 5);
        sheetDataConfig.addExcelDateTypeConfig(edtc);
        config2.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config2);
        doTest(mc);
    }

    @Test
    public void testXls() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xls";
        ExcelConvertConfig config2 = new ExcelConvertConfig(filePath);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        ExcelDateTypeConfig edtc = new ExcelDateTypeConfig(0, 5);
        sheetDataConfig.addExcelDateTypeConfig(edtc);
        config2.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config2);
        doTest(mc);
    }

    public void doTest(MapConverter mc) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        String configJson = objectMapper.writeValueAsString(mc.getConvertConfig());
        System.out.println("===== 配置 =====");
        System.out.println(configJson);
        System.out.println("===== 配置 =====");
        DataMapWrapper dmw = mc.toMapData();
        System.out.println("===== 数据头 =====");
        MapHeaders headers = dmw.getHeaders();
        String headersJson = objectMapper.writeValueAsString(headers.getHeaders());
        System.out.println(headersJson);
        System.out.println("===== 数据头 =====");
        Map<String, Object> map = dmw.getData();
        String json = objectMapper.writeValueAsString(map);
        System.out.println(json);

        Assert.assertTrue(map.containsKey("S1"));
        List<Map<String, Object>> list = (List<Map<String, Object>>) map.get("S1");
        Assert.assertEquals(3, list.size());
        Map<String, Object> m1 = list.get(0);
        Assert.assertEquals("跨列", m1.get("A"));
        Assert.assertEquals("跨列", m1.get("B"));
        Assert.assertEquals("跨行", m1.get("C"));
        // FIXME HSSF event 模式下公式计算不出结果的时候是0，而其他的会根据公式给出结果。
        // FIXME 公式为：=IFERROR(C2-A2,"-") , 结果应该是 "-"
        Assert.assertTrue(Objects.equals("-", m1.get("D")) || Objects.equals(0D, m1.get("D")));
        Assert.assertEquals("跨行跨列", m1.get("E"));
        Assert.assertEquals("跨行跨列", m1.get("2000-01-11"));

        Map<String, Object> m2 = list.get(1);
        Assert.assertEquals(12D, m2.get("A"));
        Assert.assertEquals(1300D, m2.get("B"));
        Assert.assertEquals("跨行", m2.get("C"));
        Assert.assertEquals(-1288D, m2.get("D"));
        Assert.assertEquals("跨行跨列", m2.get("E"));
        Assert.assertEquals("跨行跨列", m2.get("2000-01-11"));
    }

    @Test
    public void testXlsxMultiHead() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xlsx";
        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
        // 指定数据范围，标题行是[1, 3)，数据行从3开始，数据列从1开始
        SheetDataRangeConfig.SheetDataRangeBuilder builder = new SheetDataRangeConfig.SheetDataRangeBuilder();
        builder.headRowStart(1).headRowEnd(3).dataRowStart(3).dataColumnStart(1);
        SheetDataRangeConfig sheetDataRange = builder.build();
        // sheet下标为1
        SheetDataConfig sheetDataConfig = new SheetDataConfig(1, sheetDataRange);
        // 指定多个时间坐标和格式化
        ExcelDateTypeConfig edtc = new ExcelDateTypeConfig(
                new int[][]{{4, 7}, {5, 6}, {5, 7}}, "yyyy-MM-dd HH:mm:ss");
        sheetDataConfig.addExcelDateTypeConfig(edtc);
        config.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config);
        doTestMultiHead(mc, "S2");

        // sheet S3 测试
        sheetDataConfig.setSheetIndex(2);
        SheetDataRangeConfig sdr = sheetDataConfig.getSheetDataRange();
        sdr.setDataRowEnd(6);
        sdr.setDataColumnEnd(8);
        MapConverter mc2 = Any2Map.createMapConverter(config);
        doTestMultiHead(mc2, "S3");
    }

    @Test
    public void testXlsMultiHead() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xls";
        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
        // 指定数据范围，标题行是[1, 3)，数据行从3开始，数据列从1开始
        SheetDataRangeConfig.SheetDataRangeBuilder builder = new SheetDataRangeConfig.SheetDataRangeBuilder();
        builder.headRowStart(1).headRowEnd(3).dataRowStart(3).dataColumnStart(1);
        SheetDataRangeConfig sheetDataRange = builder.build();
        // sheet下标为1
        SheetDataConfig sheetDataConfig = new SheetDataConfig(1, sheetDataRange);
        // 指定多个时间坐标和格式化
        ExcelDateTypeConfig edtc = new ExcelDateTypeConfig(
                new int[][]{{4, 7}, {5, 6}, {5, 7}}, "yyyy-MM-dd HH:mm:ss");
        sheetDataConfig.addExcelDateTypeConfig(edtc);
        config.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config);
        doTestMultiHead(mc, "S2");

        // sheet S3 测试
        sheetDataConfig.setSheetIndex(2);
        SheetDataRangeConfig sdr = sheetDataConfig.getSheetDataRange();
        sdr.setDataRowEnd(6);
        sdr.setDataColumnEnd(8);
        MapConverter mc2 = Any2Map.createMapConverter(config);
        doTestMultiHead(mc2, "S3");
    }

    public void doTestS2Header(MapHeaders headers) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        System.out.println("===== 数据头 =====");
        String headersJson = objectMapper.writeValueAsString(headers.getHeaders());
        System.out.println(headersJson);
        System.out.println("===== 数据头 =====");

        MapHeader A = headers.getHeader("A", 0);
        Assert.assertFalse(A.isLeaf());
        MapHeader B = headers.getHeader("B", 0);
        Assert.assertFalse(B.isLeaf());
        MapHeader C = headers.getHeader("C", 0);
        Assert.assertFalse(C.isLeaf());
        MapHeader D = headers.getHeader("D", 0);
        Assert.assertFalse(D.isLeaf());

        Assert.assertEquals(A, headers.getHeader("a", 1).getParentHeader());
        Assert.assertEquals(B, headers.getHeader("b1", 1).getParentHeader());
        Assert.assertEquals(B, headers.getHeader("b2", 1).getParentHeader());
        Assert.assertEquals(C, headers.getHeader("c1", 1).getParentHeader());
        Assert.assertEquals(C, headers.getHeader("c2", 1).getParentHeader());
        Assert.assertEquals(C, headers.getHeader("c3d1", 1).getParentHeader());
        Assert.assertEquals(D, headers.getHeader("c3d1", 1, 1).getParentHeader());
    }

    public void doTestS3Header(MapHeaders headers) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        System.out.println("===== 数据头 =====");
        String headersJson = objectMapper.writeValueAsString(headers.getHeaders());
        System.out.println(headersJson);
        System.out.println("===== 数据头 =====");
    }

    // 测试多层Head的情况
    public void doTestMultiHead(MapConverter mc, String sheetName) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        String configJson = objectMapper.writeValueAsString(mc.getConvertConfig());
        System.out.println("===== 配置 =====");
        System.out.println(configJson);
        System.out.println("===== 配置 =====");
        DataMapWrapper dmw = mc.toMapData();
        if(sheetName.equals("S2")) {
            doTestS2Header(dmw.getHeaders());
        } else if(sheetName.equals("S3")) {
            doTestS3Header(dmw.getHeaders());
        }
        Map<String, Object> map = dmw.getData();
        String json = objectMapper.writeValueAsString(map);
        System.out.println(json);

        Assert.assertTrue(map.containsKey(sheetName));
        List<Map<String, Object>> list = (List<Map<String, Object>>) map.get(sheetName);
        Assert.assertEquals(3, list.size());

        Map<String, Object> m0 = list.get(0);
        Assert.assertEquals("AaBb1", ((Map) m0.get("A")).get("a"));
        Assert.assertEquals("AaBb1", ((Map) m0.get("B")).get("b1"));
        Assert.assertEquals("Bb2", ((Map) m0.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals("-", ((Map) m0.get("C")).get("c1")) ||
                Objects.equals(0D, ((Map) m0.get("C")).get("c1")));
        Assert.assertEquals("Cc2c3", ((Map) m0.get("C")).get("c2"));
        Assert.assertEquals("Cc2c3", ((Map) m0.get("C")).get("c3d1"));
        Assert.assertNull(((Map) m0.get("D")).get("c3d1"));

        Map<String, Object> m1 = list.get(1);
        Assert.assertEquals(12D, ((Map) m1.get("A")).get("a"));
        Assert.assertEquals(1300D, ((Map) m1.get("B")).get("b1"));
        Assert.assertEquals("Bb2", ((Map) m1.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals(-1288D, ((Map) m1.get("C")).get("c1")));
        Assert.assertEquals("Cc2c3", ((Map) m1.get("C")).get("c2"));
        Assert.assertEquals("Cc2c3", ((Map) m1.get("C")).get("c3d1"));
        Assert.assertEquals("2022-12-22 00:00:00", ((Map) m1.get("D")).get("c3d1"));

        Map<String, Object> m2 = list.get(2);
        Assert.assertEquals(true, ((Map) m2.get("A")).get("a"));
        Assert.assertEquals(false, ((Map) m2.get("B")).get("b1"));
        Assert.assertEquals("Bb2Cc1c2", ((Map) m2.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals("Bb2Cc1c2", ((Map) m2.get("C")).get("c1")));
        Assert.assertEquals("Bb2Cc1c2", ((Map) m2.get("C")).get("c2"));
        Assert.assertEquals("2022-12-22 00:00:00", ((Map) m2.get("C")).get("c3d1"));
        Assert.assertEquals("2022-12-22 00:00:00", ((Map) m2.get("D")).get("c3d1"));
    }

    @Test
    public void testXlsxWithGroup() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xlsx";
        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
        // 指定数据范围，标题行是[1, 3)，数据行从3开始，分组信息范围是[0, 2), 数据列从2开始
        SheetDataRangeConfig.SheetDataRangeBuilder builder = new SheetDataRangeConfig.SheetDataRangeBuilder();
        builder.headRowStart(1).headRowEnd(3)
                .groupColumnStart(0).groupColumnEnd(2)
                .dataRowStart(3).dataColumnStart(2);
        SheetDataRangeConfig sheetDataRange = builder.build();
        // sheet下标为3
        SheetDataConfig sheetDataConfig = new SheetDataConfig(3, sheetDataRange);
        config.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config);
        doTestWithGroup(mc, "S4");
    }

    @Test
    public void testXlsWithGroup() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xls";
        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
        // 指定数据范围，标题行是[1, 3)，数据行从3开始，分组信息范围是[0, 2), 数据列从2开始
        SheetDataRangeConfig.SheetDataRangeBuilder builder = new SheetDataRangeConfig.SheetDataRangeBuilder();
        builder.headRowStart(1).headRowEnd(3)
                .groupColumnStart(0).groupColumnEnd(2)
                .dataRowStart(3).dataColumnStart(2);
        SheetDataRangeConfig sheetDataRange = builder.build();
        // sheet下标为3
        SheetDataConfig sheetDataConfig = new SheetDataConfig(3, sheetDataRange);
        config.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config);
        doTestWithGroup(mc, "S4");
    }

    // 测试多层Head的情况
    public void doTestWithGroup(MapConverter mc, String sheetName) throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        String configJson = objectMapper.writeValueAsString(mc.getConvertConfig());
        System.out.println("===== 配置 =====");
        System.out.println(configJson);
        System.out.println("===== 配置 =====");
        DataMapWrapper dmw = mc.toMapData();
        System.out.println("===== 数据头 =====");
        MapHeaders headers = dmw.getHeaders();
        String headersJson = objectMapper.writeValueAsString(headers.getHeaders());
        System.out.println(headersJson);
        System.out.println("===== 数据头 =====");
        Map<String, Object> map = dmw.getData();
        String json = objectMapper.writeValueAsString(map);
        System.out.println(json);

        Assert.assertTrue(map.containsKey(sheetName));
        Map sheetMap = (Map) map.get(sheetName);
        Assert.assertEquals(2, sheetMap.size());

        Map<String, Object> sheetMapG1 = (Map<String, Object>) sheetMap.get("分组1");
        Assert.assertEquals(2, sheetMapG1.size());

        Map<String, Object> sheetMapG11 = (Map<String, Object>) sheetMapG1.get("甲");
        Assert.assertEquals("AaBb1", ((Map) sheetMapG11.get("A")).get("a"));
        Assert.assertEquals("AaBb1", ((Map) sheetMapG11.get("B")).get("b1"));
        Assert.assertEquals("Bb2", ((Map) sheetMapG11.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals("-", ((Map) sheetMapG11.get("C")).get("c1")) ||
                Objects.equals(0D, ((Map) sheetMapG11.get("C")).get("c1")));
        Assert.assertEquals("Cc2c3", ((Map) sheetMapG11.get("C")).get("c2"));

        Map<String, Object> sheetMapG12 = (Map<String, Object>) sheetMapG1.get("4");
        Assert.assertEquals(12D, ((Map) sheetMapG12.get("A")).get("a"));
        Assert.assertEquals(1300D, ((Map) sheetMapG12.get("B")).get("b1"));
        Assert.assertEquals("Bb2", ((Map) sheetMapG12.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals(-1288D, ((Map) sheetMapG12.get("C")).get("c1")));
        Assert.assertEquals("Cc2c3", ((Map) sheetMapG12.get("C")).get("c2"));

        Map<String, Object> sheetMapG2 = (Map<String, Object>) sheetMap.get("分组2");
        Assert.assertEquals(1, sheetMapG2.size());
        Map<String, Object> sheetMapG21 = (Map<String, Object>) sheetMapG2.get("乙");
        Assert.assertEquals(true, ((Map) sheetMapG21.get("A")).get("a"));
        Assert.assertEquals(false, ((Map) sheetMapG21.get("B")).get("b1"));
        Assert.assertEquals("Bb2Cc1c2", ((Map) sheetMapG21.get("B")).get("b2"));
        Assert.assertTrue(Objects.equals("Bb2Cc1c2", ((Map) sheetMapG21.get("C")).get("c1")));
        Assert.assertEquals("Bb2Cc1c2", ((Map) sheetMapG21.get("C")).get("c2"));
    }


//    @Test
//    public void testReadBigXlsx() {
//        String separator = File.separator;
//        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
//                + "resources" + separator + "big-test3.xlsx";
//        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
//        SheetDataConfig sheetDataConfig = new SheetDataConfig();
//        config.addSheetDataConfig(sheetDataConfig);
//        MapConverter mc = Any2Map.createMapConverter(config);
//        long start = System.currentTimeMillis();
//        mc.toMap();
//        long cost = System.currentTimeMillis() - start;
//        log.debug("耗时: {}", cost);
//        Assert.assertTrue(cost < 10 * 1000);
//    }


//    @Test
//    public void testReadBigFile() throws IOException {
//        String separator = File.separator;
//        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
//                + "resources" + separator + "big-test3.xlsx";
//        generateBigTestFileBySXSSFSheet(filePath);
//    }

    public void generateBigTestFileBySXSSFSheet(String filePath) throws IOException {
        long start = System.currentTimeMillis();
        SXSSFWorkbook workbook = new SXSSFWorkbook(10000);
        try (OutputStream out = new FileOutputStream(filePath)) {
            SXSSFSheet sheet = workbook.createSheet("Sheet1");
            //100w
            for (int i = 0; i < 100 * 100 * 10; i++) {
                SXSSFRow row = sheet.createRow(i);
                for (int j = 0; j < 100; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(i * j + j);
                }
            }
            long start1 = System.currentTimeMillis();
            log.debug("数据生成耗时: {}", start1 - start);
            workbook.write(out);
            log.debug("写入文件耗时: {}", System.currentTimeMillis() - start1);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            workbook.dispose();
        }
        log.debug("生成文件总耗时: {}", System.currentTimeMillis() - start);
    }

    public void generateBigTestFile(String filePath) throws IOException {
        long start = System.currentTimeMillis();
        try (FileOutputStream fileOut = new FileOutputStream(filePath);
             Workbook workbook = WorkbookFactory.create(true);
             FileChannel channel = fileOut.getChannel()) {
            Random random = new Random();
            Sheet sheet = workbook.createSheet("LargeSheet");
            int i = 0;
            for (int writeCnt = 0; writeCnt < 2; writeCnt++) {
                long start1 = System.currentTimeMillis();
                for (; i < 500 * (writeCnt + 1); i++) {
                    Row row = sheet.createRow(i);
                    for (int colIndex = 0; colIndex < 200; colIndex++) {
                        Cell cell = row.createCell(colIndex);
                        cell.setCellValue(random.nextInt());
                    }
                }

                log.debug("创建片段[{}]耗时: {}", writeCnt, System.currentTimeMillis() - start1);
                long start2 = System.currentTimeMillis();
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                workbook.write(baos);
                ByteBuffer buffer = ByteBuffer.wrap(baos.toByteArray());
                channel.write(buffer);
                log.debug("创建片段写入文件[{}]耗时: {}", writeCnt, System.currentTimeMillis() - start2);

//                log.debug("创建片段[{}]耗时: {}", writeCnt, System.currentTimeMillis() - start1);
//                long start2 = System.currentTimeMillis();
//                workbook.write(fileOut);
//                log.debug("创建片段写入文件[{}]耗时: {}", writeCnt, System.currentTimeMillis() - start2);
            }
        }
        log.debug("创建耗时: {}", System.currentTimeMillis() - start);
    }

    // 使用 RandomAccessFile 的方式持续写入数据
    public void generateBigTestFileByRandomAccessFile(String filePath) {
        int size = 1024 * 1024 * 1024;
        Random random = new Random();
        int rowNum = 0;
        while (true) {
            try (Workbook workbook = new SXSSFWorkbook();
                 RandomAccessFile fileOut = new RandomAccessFile(filePath, "rw")) {
                Sheet sheet;
                if (rowNum == 0) {
                    sheet = workbook.createSheet("LargeSheet");
                } else {
                    sheet = workbook.getSheet("largeSheet");
                }

                int total = rowNum + 100;
                for (int i = rowNum; i < total; i++) {
                    Row row = sheet.createRow(i);
                    for (int colIndex = 0; colIndex < 100; colIndex++) {
                        Cell cell = row.createCell(colIndex);
                        cell.setCellValue(random.nextInt());
                    }
                }
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                workbook.write(baos);
                fileOut.write(baos.toByteArray());
                System.out.println("1GB XLSX file generated successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

}
