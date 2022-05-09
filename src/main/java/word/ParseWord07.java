package word;

import model.ElCore;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-09 17:06
 * @Description: 2007 版本 word.docx 文档填充  
 */
public class ParseWord07 {
    
    public void fillWord(XWPFDocument document, Map<String, Object> params) {
        if(document == null){
            throw new NullPointerException("XWPFDocument is null");
        }
        //填充段落
        fillAllParagraph(document.getParagraphs(), params);
    }
    
    private void fillAllParagraph(List<XWPFParagraph> paragraphs, Map<String, Object> params) {
        for (XWPFParagraph paragraph : paragraphs) {
            if (paragraph.getText().contains(ElCore.START_LABEL)) {
                this.fillThisParagraph(paragraph, params);
            }
        }
    }
    
    private void fillThisParagraph(XWPFParagraph paragraph, Map<String, Object> params) {
        
        // 拿到的第一个run,用来set值,可以保存格式
        XWPFRun currentRun = null;
        // 存放当前的text
        StringBuilder currentText = new StringBuilder();
        // 判断是不是已经遇到开始标签
        boolean ifFind = false;
        // 存储遇到的run,把他们置空
        List<Integer> runIndex = new ArrayList<Integer>();
        
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            // 如果为空或者 "" 跳过当前循环
            if (text == null || text == "") {
                continue;
            }
            if (ifFind) {
                currentText.append(text);
                if (!currentText.toString().contains(ElCore.START_LABEL)) {
                    ifFind = false;
                    runIndex.clear();
                } else {
//                    runIndex.add(i);
                }
                if (currentText.toString().contains(ElCore.END_LABEL)) {
                    System.out.println("1111111111");
                    System.out.println(currentText);
//                    changeValues(paragraph, currentRun, currentText, runIndex, map);
                    currentText = new StringBuilder();
                    ifFind = false;
                }
            } else if(text.contains(ElCore.START_LABEL) || ElCore.START_LABEL.contains(text)) {
                currentText = new StringBuilder(text);
                ifFind = true;
                currentRun = run;
            } else {
                currentText = new StringBuilder();
            }
            if (currentText.toString().contains(ElCore.END_LABEL)) {
                ifFind = false;
                System.out.println("22222222222");
                System.out.println(currentText);
            }
        }
    }
}
