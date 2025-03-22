package com.modern.tools.xlsx;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Xlsx2MapConverterBySaxTest
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverterBySax2Test {

    @Test
    public void test2() {
        ObjectMapper objectMapper = new ObjectMapper();
        try {
            Xlsx2MapConverterBySax2 x2ms = new Xlsx2MapConverterBySax2();
            XlsxConvertConfig config2 = new XlsxConvertConfig();
            SheetDataConfig sheetDataConfig = new SheetDataConfig();
            config2.addSheetDataConfig(sheetDataConfig);
            x2ms.setConvertConfig(config2);
            String configJson = objectMapper.writeValueAsString(x2ms.getConfig());
            System.out.println("===== 配置 =====");
            System.out.println(configJson);
            System.out.println("===== 配置 =====");
            Map<String, Object> map = x2ms.toMap("D:\\code\\open\\any2map\\src\\test\\resources\\test.xlsx");
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
    public void test2Big2() {
        ObjectMapper objectMapper = new ObjectMapper();
        try {
            Xlsx2MapConverterBySax2 x2ms = new Xlsx2MapConverterBySax2();
            XlsxConvertConfig config2 = new XlsxConvertConfig();
            SheetDataConfig sheetDataConfig = new SheetDataConfig();
            config2.addSheetDataConfig(sheetDataConfig);
            x2ms.setConvertConfig(config2);
            String configJson = objectMapper.writeValueAsString(x2ms.getConfig());
            System.out.println("===== 配置 =====");
            System.out.println(configJson);
            System.out.println("===== 配置 =====");
            Map<String, Object> map = x2ms.toMap("D:\\code\\open\\any2map\\src\\test\\resources\\test-big.xlsx");
            String json = objectMapper.writeValueAsString(map);
            System.out.println(json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
