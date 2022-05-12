package com.beloved.utils;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import com.beloved.core.WordCore;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-05 15:41
 * @Description: Word  操作工具类
 */
public class WordUtil {

    private static final String TEMPLATE_PATH = "/template/";
    private static final String CREDENTIALS_PATH = "/credentials/license.xml";
    
    /**
     * 生成填充模板后的文件
     * @param templatePath      模板路径
     * @param targetPath        输出目录  .docx
     * @param params            参数
     * @throws Exception
     */
    public static void templateFill(String templatePath, String targetPath,  Map<String, Object> params) throws Exception {
        
        File targetFile = new File(targetPath);
        // 是否有父级目录没有则创建
        if(!targetFile.getParentFile().exists()){
            targetFile.getParentFile().mkdirs();
        }

        FileOutputStream fos = null;

        try {
            XWPFDocument docx = WordCore.fillWord(templatePath, params);
            
            fos = new FileOutputStream(targetFile);
            
            docx.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (fos != null) {
                fos.close();
            }
        }

    }

    /**
     * 填充 word 模板
     * @param templateName  模板名称
     * @param params        填充参数
     * @return              填充后的byte数组
     * @throws Exception
     */
    public static byte[] templateFill(String templateName, Map<String, Object> params) throws Exception {
        
        String filePath = WordCore.class.getResource(TEMPLATE_PATH + templateName).getPath();

        ByteArrayOutputStream os = null;
        
        try {
            os = new ByteArrayOutputStream();
            
            XWPFDocument docx = WordCore.fillWord(filePath, params);

            docx.write(os);
            
            return os.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (os != null) {
                os.close();
            }
        }
    }

    /**
     * word 转 pdf
     * @param wordPath      word 模板路径
     * @param pdfPath       生成 pdf 路径
     * @throws Exception
     */
    public static void wordToPdf(String wordPath, String pdfPath) throws Exception {
        wordToPdf(wordPath, pdfPath, false);
    }

    /**
     * word 转 pdf
     * @param wordPath      word 模板路径
     * @param pdfPath       生成 pdf 路径
     * @param deleteWord    是否删除 word 模板
     * @throws Exception
     */
    public static void wordToPdf(String wordPath, String pdfPath, boolean deleteWord) throws Exception {
        
        InputStream is = null;
        FileOutputStream os = null;

        try {
            // 获取凭证，校验凭证
            is = WordCore.class.getResourceAsStream(CREDENTIALS_PATH);
            new License().setLicense(is);

            File pdf = new File(pdfPath);

            // 是否有父级目录没有则创建
            if(!pdf.getParentFile().exists()){
                pdf.getParentFile().mkdirs();
            }

            os = new FileOutputStream(pdf);

            // 要转换的word文件
            Document word = new Document(wordPath);
            word.save(os, SaveFormat.PDF);

            if (deleteWord) {
                Files.deleteIfExists(Paths.get(wordPath));
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (is != null) {
                is.close();
            }
            if (os != null) {
                os.close();
            }
        }
    }

    /**
     * word 转 pdf
     * @param wordByteArray     word byte[]
     * @return                  pdf byte[]
     * @throws Exception
     */
    public static byte[] wordToPdf(byte[] wordByteArray) throws Exception {
        InputStream is = null;
        InputStream wordInputStream = null;
        ByteArrayOutputStream os = null;
        
        try {
            // 获取凭证，校验凭证
            is = WordCore.class.getResourceAsStream(CREDENTIALS_PATH);
            new License().setLicense(is);

            wordInputStream = new ByteArrayInputStream(wordByteArray);
            
            // 要转换的word文件
            Document word = new Document(wordInputStream);

            os = new ByteArrayOutputStream();
            word.save(os, SaveFormat.PDF);
            
            return os.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (is != null) {
                is.close();
            }
            if (wordInputStream != null) {
                wordInputStream.close();
            }
            if (os != null) {
                os.close();
            }
        }
    }

    /**
     * 填充 word 模板后转换成 pdf
     * @param templateName  word 模板名称
     * @param params        填充参数
     * @return              转换 pdf 后的 byte[]
     * @throws Exception
     */
    public static byte[] templateFillToPdfByteArray(String templateName, Map<String, Object> params) throws Exception {

        byte[] wordByteArray = templateFill(templateName, params);

        byte[] pdfByteArray = wordToPdf(wordByteArray);

        return pdfByteArray;
    }

}
