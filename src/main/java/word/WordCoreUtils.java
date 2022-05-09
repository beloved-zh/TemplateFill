package word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.Map;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-09 17:08
 * @Description: word 模板填充
 */
public class WordCoreUtils {

    public static void fillWord(XWPFDocument document, Map<String, Object> params) {
        new ParseWord07().fillWord(document, params);
    }
    
}
