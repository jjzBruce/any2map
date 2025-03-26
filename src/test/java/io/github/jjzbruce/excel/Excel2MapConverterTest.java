package io.github.jjzbruce.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.jjzbruce.Any2Map;
import io.github.jjzbruce.MapConverter;
import org.apache.poi.ss.formula.functions.Log;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.net.URL;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.UUID;
import java.util.concurrent.ThreadLocalRandom;

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
        Map<String, Object> map = mc.toMap();
        String json = objectMapper.writeValueAsString(map);
        System.out.println(json);

        Assert.assertTrue(map.containsKey("S1"));
        List<Map<String, Object>> list = (List<Map<String, Object>>) map.get("S1");
        Assert.assertEquals(3, list.size());
        Map<String, Object> m1 = list.get(0);
        Assert.assertEquals("跨列", m1.get("A"));
        Assert.assertEquals("跨列", m1.get("B"));
        Assert.assertEquals("跨行", m1.get("C"));
        // TODO HSSF event 模式下 值是 0
        Assert.assertEquals("-", m1.get("D"));
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
    public void testReadBigXlsx() {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "big-test3.xlsx";
        ExcelConvertConfig config = new ExcelConvertConfig(filePath);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        config.addSheetDataConfig(sheetDataConfig);
        MapConverter mc = Any2Map.createMapConverter(config);
        long start = System.currentTimeMillis();
        mc.toMap();
        long cost = System.currentTimeMillis() - start;
        log.debug("耗时: {}", cost);
        Assert.assertTrue(cost < 10 * 1000);
    }


    @Test
    public void testReadBigFile() throws IOException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "big-test3.xlsx";
        generateBigTestFileBySXSSFSheet(filePath);
    }

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
