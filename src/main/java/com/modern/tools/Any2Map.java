package com.modern.tools;

import com.modern.tools.excel.AbstractExcelMapConverter;
import com.modern.tools.excel.ExcelConvertConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Constructor;
import java.util.Objects;

/**
 * main
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class Any2Map {
    private final static Logger log = LoggerFactory.getLogger(Any2Map.class);

    public static MapConverter createMapConverter(ConvertConfig config) {
        Objects.nonNull(config);
        if(ExcelConvertConfig.class.equals(config.getClass())) {
            Class<? extends AbstractExcelMapConverter> delegateImpl = ((ExcelConvertConfig) config).getDelegateImpl();
            Constructor<?> constructor = delegateImpl.getConstructors()[0];
            constructor.equals(true);
            try {
                return (MapConverter) constructor.newInstance(config);
            } catch (Throwable e) {
                log.error("创建 ExcelMapConverter 示例失败", e);
                return null;
            }
        } else {
            // TODO 待实现
            return null;
        }
    }

}
