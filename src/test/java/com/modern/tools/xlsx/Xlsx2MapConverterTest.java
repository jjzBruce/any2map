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
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("test2.xlsx")) {
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
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("test2.xlsx")) {
            Xlsx2MapConverter x2m2 = new Xlsx2MapConverter();
            XlsxConvertConfig config2 = new XlsxConvertConfig();
            SheetDataConfig sheetDataConfig = new SheetDataConfig();
            config2.addSheetDataConfig(sheetDataConfig);
            x2m2.setConvertConfig(config2);
            String configJson = objectMapper.writeValueAsString(x2m2.getConfig());
            System.out.println(configJson);
            Map<String, Object> map = x2m2.toMap(is);
            String json = objectMapper.writeValueAsString(map);
            System.out.println(json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
