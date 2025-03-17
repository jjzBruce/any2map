package com.modern.tools.xlsx;

import com.modern.tools.ConvertConfig;
import com.modern.tools.MapConverter;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Xlsx To Map
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Xlsx2MapConverter implements MapConverter<XlsxConvertConfig> {

    private XlsxConvertConfig config = new XlsxConvertConfig();

    /**
     * 转换设置
     *
     * @param config 配置
     */
    @Override
    public void setConvertConfig(XlsxConvertConfig config) {
        this.config = config;
    }

    /**
     * 输出目标 Map
     *
     * @return Map
     */
    @Override
    public Map<String, Object> toMap(Object source) {
        InputStream is = null;
        if (source instanceof InputStream) {
            is = (InputStream) source;
        }

        try (Workbook workbook = StreamingReader.builder()
                .rowCacheSize(10 * 10)
                .bufferSize(1024 * 4)
                //打开资源，可以是InputStream或者是File，注意：只能打开.xlsx格式的文件
                .open(is)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {

            }


            Sheet sheet1 = workbook.getSheetAt(0);
            handleSheet1(sheet1, productSummaryMonth);
            Sheet sheet2 = workbook.getSheetAt(1);
            handleSheet2(sheet2, productSummaryMonth);
            Sheet sheet3 = workbook.getSheetAt(2);
            handleSheet3(sheet3, productSummaryMonth);
        } catch (IOException e) {
            // TODO 处理
            throw new RuntimeException(e);
        }
    }
}
