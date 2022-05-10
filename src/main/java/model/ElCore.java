package model;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-09 17:26
 * @Description: 表达式支持
 */
public class ElCore {

    /**
     * 文本填充开始标签
     */
    public static final String START_LABEL = "{{";
    
    /**
     * 文本填充结束标签
     */
    public static final String END_LABEL = "}}";
    
    public static final CharSequence FOREACH = "each:";

    public static final String LEFT_BRACKET       = "(";
    public static final String RIGHT_BRACKET      = ")";
    
    public static final CharSequence EMPTY = "";

    private ElCore() {
    }
}
