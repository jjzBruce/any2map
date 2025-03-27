package io.github.jjzbruce.excel;

/**
 * Xlsx数据范围。包含开始下标，不包含结束下标
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangjun</a>
 * @since 1.0.0
 */
public class SheetDataRangeConfig {

    private Integer headRowStart = 0;
    private Integer headRowEnd = headRowStart + 1;

    private Integer dataRowStart = 1;
    private Integer dataRowEnd = Integer.MAX_VALUE;
    private Integer dataColumnStart = 0;
    private Integer dataColumnEnd = Integer.MAX_VALUE;

    public SheetDataRangeConfig() {
    }

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

    public static class SheetDataRangeBuilder {
        private SheetDataRangeConfig sheetDataRange;

        public SheetDataRangeBuilder() {
            this.sheetDataRange = new SheetDataRangeConfig();
        }

        public SheetDataRangeBuilder headRowStart(Integer headRowStart) {
            sheetDataRange.headRowStart = headRowStart;
            return this;
        }

        public SheetDataRangeBuilder headRowEnd(Integer headRowEnd) {
            sheetDataRange.headRowEnd = headRowEnd;
            return this;
        }

        public SheetDataRangeBuilder dataRowStart(Integer dataRowStart) {
            sheetDataRange.dataRowStart = dataRowStart;
            return this;
        }

        public SheetDataRangeBuilder dataRowEnd(Integer dataRowEnd) {
            sheetDataRange.dataRowEnd = dataRowEnd;
            return this;
        }

        public SheetDataRangeBuilder dataColumnStart(Integer dataColumnStart) {
            sheetDataRange.dataColumnStart = dataColumnStart;
            return this;
        }

        public SheetDataRangeBuilder dataColumnEnd(Integer dataColumnEnd) {
            sheetDataRange.dataColumnEnd = dataColumnEnd;
            return this;
        }

        public SheetDataRangeConfig build() {
            return sheetDataRange;
        }
    }
}
