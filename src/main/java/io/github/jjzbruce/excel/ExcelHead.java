package io.github.jjzbruce.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel标题
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class ExcelHead {

    /**
     * 标题开始行号
     */
    private int beginRowNum;

    /**
     * 标题结束行号
     */
    private int endRowNum;

    /**
     * 标题信息维护
     */
    private List<Map<Integer, String>> headValueCaches = new ArrayList<>();

    public ExcelHead(int beginRowNum, int endRowNum) {
        this.beginRowNum = beginRowNum;
        this.endRowNum = endRowNum;
    }

    /**
     * 获取标题数组。当存在多层标题时，返回标题组
     *
     * @param colNum 列下标
     */
    public String[] getHeads(int colNum) {
        String[] heads = new String[endRowNum - beginRowNum];
        int index = 0;
        for (int i = beginRowNum; i < endRowNum; i++) {
            Map<Integer, String> m = headValueCaches.get(i);
            String head = null;
            if (m != null) {
                head = m.get(colNum);
            }
            heads[index++] = head;
        }
        return heads;
    }

    /**
     * 设置标题数组。当存在多层标题时，根据行号往上依次查询标题组合成队列返回
     *
     * @param rowNum 行下标
     * @param colNum 列下标
     */
    public void setHeads(int rowNum, int colNum, String headValue) {
        Map<Integer, String> head;
        while (rowNum + 1 > headValueCaches.size()) {
            headValueCaches.add(new HashMap<>());
        }
        head = headValueCaches.get(rowNum);
        head.put(colNum, headValue);
    }

}
