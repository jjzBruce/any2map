package io.github.jjzbruce;

/**
 * 转换配置
 *
 * @author <a href="mailto:brucezhang_jjz@163.com">zhangj</a>
 * @since 1.0.0
 */
public interface ConvertConfig {

    /**
     * 数据源
     */
    Object getSource();

    /**
     * 每行数据结构是否固定。 <br/>
     * 相同则会按照<b>首行</b>作为标准来适配其他行数据，否则数据的类型是混合的。
     */
    boolean isFixedSchema();

}
