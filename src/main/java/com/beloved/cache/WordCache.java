package com.beloved.cache;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.InputStream;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-11 17:58
 * @Description: word 缓存中心
 */
public class WordCache {

    public static XWPFDocument getXWPFDocument(String url) {
        InputStream is = null;
        try {
            is = CacheManager.loaderFileStream(url);
            XWPFDocument document = new XWPFDocument(is);
            return document;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return null;
    }
    
}
