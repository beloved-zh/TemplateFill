import com.beloved.core.ImageEntity;
import com.beloved.core.WordCore;
import com.beloved.utils.BarCodeUtil;
import com.beloved.utils.QRCodeUtil;
import com.beloved.utils.WordUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
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

    private static final String templatePath = "D:\\project\\TemplateFill\\src\\main\\resources\\template\\TestTemplate.docx";

    private static final String targetPath = "D:\\download\\"+System.currentTimeMillis()+".docx";
    
    private static final String targetPdfPath = "D:\\download\\"+System.currentTimeMillis()+".pdf";
    
    private static final String templateName = "TestTemplate.docx";
    
    @BeforeClass
    public static void beforeClass() throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        params.put("bar_code", new ImageEntity(BarCodeUtil.getBarCodeByteArray("1234567"), 150, 50));
        params.put("username", "张三");
        params.put("age", 20);
        params.put("sex", "男");
        params.put("birthday", "2020-04-04");          //https://img0.baidu.com/it/u=4168010673,707269819
        params.put("image_photo", new ImageEntity("https://joeschmoe.io/api/v1/random", 100, 100));
        params.put("qr_code", new ImageEntity(QRCodeUtil.getQRCodeByteArray("https://www.bilibili.com/"), 100, 100));
        params.put("header", "页头");
        params.put("foot", "页脚");


        ArrayList<Map<String, Object>> table = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            HashMap<String, Object> map = new HashMap<>();
            map.put("username", "张三");
            map.put("age", 20 + i);
            map.put("sex", "男");
            map.put("image_photo", new ImageEntity("https://joeschmoe.io/api/v1/random", 100, 100));
            map.put("qr_code", new ImageEntity(QRCodeUtil.getQRCodeByteArray("https://www.bilibili.com/"), 100, 100));
            table.add(map);
        }
        
        params.put("list", table);
        
        params.put("print_date", sdf.format(new Date()));
        
    }
    
    @Test
    public void test01() throws Exception {
        String templatePath = "D:\\project\\TemplateFill\\src\\test\\resources\\template\\TestTemplate.docx";
        
        String targetPath = "D:\\download\\"+System.currentTimeMillis()+".docx";
        

        XWPFDocument document = WordCore.fillWord("D:\\project\\TemplateFill\\src\\test\\resources\\template\\TestTemplate.docx", params);

        FileOutputStream os = new FileOutputStream(targetPath);
        document.write(os);
        
    }

    @Test
    public void test02() throws Exception {
        File templateFile = new File("D:\\project\\TemplateFill\\src\\test\\resources\\template\\TestTemplate.docx");

        String targetPath = "D:\\download\\"+System.currentTimeMillis()+".docx";
        
        InputStream inputStream = new FileInputStream(templateFile);

        XWPFDocument document = new XWPFDocument(inputStream);

        WordCore.fillWord(document, params);
        
        FileOutputStream os = new FileOutputStream(targetPath);
        document.write(os);

    }
    
    @Test
    public void test03() {
        String str = "{{abc}}";

        System.out.println(str.indexOf("{{"));

        System.out.println(StringUtils.substringBetween(str, "{{", "}}"));
    }
    
    @Test
    public void test04() throws Exception {
        WordUtil.templateFill(templatePath, targetPath, params);
    }

    @Test
    public void test05() throws Exception {
        byte[] bytes = WordUtil.templateFill(templateName, params);
        
        FileOutputStream os = new FileOutputStream(targetPath);
        os.write(bytes);
    }

    @Test
    public void test06() throws Exception {
        byte[] bytes = WordUtil.templateFillToPdfByteArray(templateName, params);

        FileOutputStream os = new FileOutputStream(targetPdfPath);
        os.write(bytes);
    }
    
}
