package com.modern.tools.xlsx;

/**
 * Xlsx数据范围
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class SheetDataRange {

    private Integer headRowStart = 0;

    private Integer dataRowStart = 1;
    private Integer headRowEnd = dataRowStart;

    private Integer dataRowEnd = Integer.MAX_VALUE;
    private Integer dataColumnStart = 0;
    private Integer dataColumnEnd = Integer.MAX_VALUE;

    public Integer getDataRowStart() {
        return dataRowStart;
    }

    public void setDataRowStart(Integer dataRowStart) {
        this.dataRowStart = dataRowStart;
    }

    public Integer getDataRowEnd() {
        return dataRowEnd;
    }

    public void setDataRowEnd(Integer dataRowEnd) {
        this.dataRowEnd = dataRowEnd;
    }

    public Integer getDataColumnStart() {
        return dataColumnStart;
    }

    public void setDataColumnStart(Integer dataColumnStart) {
        this.dataColumnStart = dataColumnStart;
    }

    public Integer getDataColumnEnd() {
        return dataColumnEnd;
    }

    public void setDataColumnEnd(Integer dataColumnEnd) {
        this.dataColumnEnd = dataColumnEnd;
    }

    public Integer getHeadRowStart() {
        return headRowStart;
    }

    public Integer getHeadRowEnd() {
        return headRowEnd;
    }

    public void setHeadRowEnd(Integer headRowEnd) {
        this.headRowEnd = headRowEnd;
    }

    @Override
    public String toString() {
        return "SheetDataRange{" +
                "headRowStart=" + headRowStart +
                ", headRowEnd=" + headRowEnd +
                ", dataRowStart=" + dataRowStart +
                ", dataRowEnd=" + dataRowEnd +
                ", dataColumnStart=" + dataColumnStart +
                ", dataColumnEnd=" + dataColumnEnd +
                '}';
    }
}
