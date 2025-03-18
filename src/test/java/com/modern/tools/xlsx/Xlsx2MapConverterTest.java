package com.modern.tools.xlsx;

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
        try (InputStream is = Thread.currentThread().getClass().getResourceAsStream("excel.xlsx")) {
            Xlsx2MapConverter x2m = new Xlsx2MapConverter();
            List<Map<String, Object>> listMap = x2m.toListMap(is);
            System.out.println(listMap);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
