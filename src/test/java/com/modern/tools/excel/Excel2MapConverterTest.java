package com.modern.tools.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.modern.tools.Any2Map;
import com.modern.tools.MapConverter;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Xlsx2MapConverter
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class Excel2MapConverterTest {

    @Test
    public void test2() throws JsonProcessingException {
        String separator = File.separator;
        String filePath = System.getProperty("user.dir") + separator + "src" + separator + "test" + separator
                + "resources" + separator + "test.xlsx";
        ObjectMapper objectMapper = new ObjectMapper();
        ExcelConvertConfig config2 = new ExcelConvertConfig(filePath,
                Excel2MapConverter.class);
        SheetDataConfig sheetDataConfig = new SheetDataConfig();
        config2.addSheetDataConfig(sheetDataConfig);
        MapConverter x2m2 = Any2Map.createMapConverter(config2);

        String configJson = objectMapper.writeValueAsString(x2m2.getConvertConfig());
        System.out.println("===== 配置 =====");
        System.out.println(configJson);
        System.out.println("===== 配置 =====");
        Map<String, Object> map = x2m2.toMap();
        String json = objectMapper.writeValueAsString(map);
        System.out.println(json);

        Assert.assertTrue(map.containsKey("S1"));
        List<Map<String, Object>> list = (List<Map<String, Object>>) map.get("S1");
        Assert.assertEquals(3, list.size());
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
    }

}
