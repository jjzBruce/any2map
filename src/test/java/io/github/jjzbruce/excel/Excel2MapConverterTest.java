package io.github.jjzbruce.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.jjzbruce.Any2Map;
import io.github.jjzbruce.MapConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * Xlsx2MapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class Excel2MapConverterTest {

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
    public void testReadBigFile() {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "big-test.xlsx";
        // 产生 1GB 的文件
        generateBigTestFile(filePath);
    }

    public void generateBigTestFile(String filePath) {
        int size = 1024 * 1024 * 1024;
        Random random = new Random();
        int rowNum = 0;
        while (true) {
            try (Workbook workbook = new SXSSFWorkbook();
                 RandomAccessFile fileOut = new RandomAccessFile(filePath, "rw")) {
                Sheet sheet;
                if(rowNum == 0) {
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
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

}
