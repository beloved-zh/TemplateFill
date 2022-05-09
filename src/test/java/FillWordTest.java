import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.BeforeClass;
import org.junit.Test;
import word.WordCoreUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.HashMap;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-09 14:12
 * @Description: word填充测试
 */
public class FillWordTest {

    private static final HashMap<String, Object> params = new HashMap<>();
    
    @BeforeClass
    public static void beforeClass() throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    }
    
    @Test
    public void test01() throws Exception {
        File templateFile = new File("D:\\project\\TemplateFill\\src\\test\\resources\\template\\测试模板.docx");
        
        String targetPath = "D:\\download\\"+System.currentTimeMillis()+".docx";

        InputStream inputStream = new FileInputStream(templateFile);
        
        XWPFDocument document = new XWPFDocument(inputStream);

        WordCoreUtils.fillWord(document, params);
        
    }
    
}
