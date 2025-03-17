package com.modern.tools.xlsx;

import com.modern.tools.ConvertConfig;
import org.apache.commons.compress.utils.Lists;

import java.util.List;

/**
 * Xlsx Convert Config
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public class XlsxConvertConfig implements ConvertConfig {

    /**
     * 需要处理的 sheet 下标
     */
    private List<Integer> sheetIndexes = Lists.newArrayList();


    public List<Integer> getSheetIndexes() {
        return sheetIndexes;
    }
}
