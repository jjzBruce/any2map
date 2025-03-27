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
public class ExcelGroup {

    /**
     * 数据开始行下标
     */
    private int beginRowNum;

    /**
     * 标题开始行号
     */
    private int beginColumnNum;

    /**
     * 标题结束行号
     */
    private int endColumnNum;

    /**
     * 标题信息维护
     */
    private List<Map<Integer, String>> groupCaches = new ArrayList<>();

    public ExcelGroup(int beginRowNum, int beginColumnNum, int endColumnNum) {
        this.beginRowNum = beginRowNum;
        this.beginColumnNum = beginColumnNum;
        this.endColumnNum = endColumnNum;
    }

    public int getBeginRowNum() {
        return beginRowNum;
    }

    /**
     * 获取标题数组。当存在多层标题时，返回标题组
     *
     * @param rowNum 行下标
     */
    public String[] getGroups(int rowNum) {
        String[] groups = new String[endColumnNum - beginColumnNum];
        int index = 0;
        for (int i = beginColumnNum; i < endColumnNum; i++) {
            Map<Integer, String> m = groupCaches.get(i);
            String head = null;
            if (m != null) {
                head = m.get(rowNum);
            }
            groups[index++] = head;
        }
        return groups;
    }

    /**
     * 设置标题数组。当存在多层标题时，根据行号往上依次查询标题组合成队列返回
     *
     * @param rowNum 行下标
     * @param colNum 列下标
     */
    public void setGroups(int rowNum, int colNum, String groupValue) {
        Map<Integer, String> group;
        while (colNum + 1 > groupCaches.size()) {
            groupCaches.add(new HashMap<>());
        }
        group = groupCaches.get(colNum);
        if(groupValue == null || "null".equals(groupValue)) {
            // 开启分组的时候，组名默认为
            groupValue = rowNum + "";
        }
        group.put(rowNum, groupValue);
    }

}
