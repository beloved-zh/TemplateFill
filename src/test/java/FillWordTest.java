import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.BeforeClass;
import org.junit.Test;
import word.WordCoreUtils;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

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
        params.put("bar_code", 1234567856);
        params.put("username", "张三");
        params.put("age", 20);
        params.put("sex", "男");
        params.put("birthday", "2020-04-04");
        params.put("image_photo", "image_photo");
        params.put("qr_code", "qr_code");


        ArrayList<Map<String, Object>> table = new ArrayList<>();
        for (int i = 0; i < 2; i++) {
            HashMap<String, Object> map = new HashMap<>();
            map.put("username", "张三");
            map.put("age", 20 + i);
            map.put("sex", "男");
            table.add(map);
        }
        
        params.put("list", table);
        
        params.put("print_date", sdf.format(new Date()));
        
    }
    
    @Test
    public void test01() throws Exception {
        File templateFile = new File("D:\\project\\TemplateFill\\src\\test\\resources\\template\\测试模板.docx");
        
        String targetPath = "D:\\download\\"+System.currentTimeMillis()+".docx";

        InputStream inputStream = new FileInputStream(templateFile);
        
        XWPFDocument document = new XWPFDocument(inputStream);

        WordCoreUtils.fillWord(document, params);

        FileOutputStream os = new FileOutputStream(targetPath);
        document.write(os);
        
    }
    
}
