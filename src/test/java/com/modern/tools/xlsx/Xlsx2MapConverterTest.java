package com.modern.tools.xlsx;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Xlsx2MapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverterTest {

    @Test
    public void test() {
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xlsx")) {
            Xlsx2MapConverter x2m = new Xlsx2MapConverter();
            Map<String, Object> listMap = x2m.toMap(is);
            Assert.assertTrue(listMap.isEmpty());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    public void test2() {
        ObjectMapper objectMapper = new ObjectMapper();
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xlsx")) {
            Xlsx2MapConverter x2m2 = new Xlsx2MapConverter();
            XlsxConvertConfig config2 = new XlsxConvertConfig();
            SheetDataConfig sheetDataConfig = new SheetDataConfig();
            config2.addSheetDataConfig(sheetDataConfig);
            x2m2.setConvertConfig(config2);
            String configJson = objectMapper.writeValueAsString(x2m2.getConfig());
            System.out.println("===== 配置 =====");
            System.out.println(configJson);
            System.out.println("===== 配置 =====");
            Map<String, Object> map = x2m2.toMap(is);
            String json = objectMapper.writeValueAsString(map);
            System.out.println(json);

            Assert.assertTrue(map.containsKey("S1"));
            List<Map<String, Object>> list = (List<Map<String, Object>>) map.get("S1");
            Assert.assertEquals(2, list.size());
            Map<String, Object> m1 = list.get(0);
            Assert.assertEquals("跨列", m1.get("A"));
            Assert.assertEquals("跨列", m1.get("B"));
            Assert.assertEquals("跨行", m1.get("C"));
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

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    public void test2Big() {
        ObjectMapper objectMapper = new ObjectMapper();
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("test-big.xlsx")) {
            Xlsx2MapConverter x2m2 = new Xlsx2MapConverter();
            XlsxConvertConfig config2 = new XlsxConvertConfig();
            SheetDataConfig sheetDataConfig = new SheetDataConfig();
            config2.addSheetDataConfig(sheetDataConfig);
            x2m2.setConvertConfig(config2);
            String configJson = objectMapper.writeValueAsString(x2m2.getConfig());
            System.out.println("===== 配置 =====");
            System.out.println(configJson);
            System.out.println("===== 配置 =====");
            Map<String, Object> map = x2m2.toMap(is);
            String json = objectMapper.writeValueAsString(map);
            System.out.println(json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
