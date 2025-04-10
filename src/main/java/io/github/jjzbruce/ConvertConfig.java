package io.github.jjzbruce;

/**
 * 转换配置
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public interface ConvertConfig {

    Object getSource();

    /**
     * 标题是日期时间的格式化表示
     */
    default String getHeaderFormat() {
        return "yyyy-MM-dd";
    }

}
