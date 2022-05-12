package com.beloved.cache;

import org.apache.poi.util.IOUtils;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-11 17:30
 * @Description: 文件加载
 */
public class FileLoader {

    private static final String HTTP = "http";
   
    /**
     * 连接网络文件超时时间
     */
    private static final int CONNECT_TIMEOUT = 30 * 1000;
   
    /**
     * 读取网络文件超时时间
     */
    private static final int READ_TIMEOUT = 60 * 1000;

    /**
     * 文件大小限制
     */
    private static final int MAX_BYTE_SIZE = 1024 * 1;

    /**
     * 加载文件
     * @param urlOrPath 网络地址 或 文件路径（相对路径 | 绝对路径）
     * @return
     */
    public static byte[] loaderFile(String urlOrPath) {
        InputStream is = null;
        ByteArrayOutputStream baos = null;
        try {
            if (urlOrPath.toLowerCase().startsWith(HTTP)) {
                URL url = new URL(urlOrPath);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setConnectTimeout(CONNECT_TIMEOUT);
                connection.setReadTimeout(READ_TIMEOUT);
                connection.setDoInput(true);
                is = connection.getInputStream();
            } else {
                try {
                    // 绝对路径读取文件
                    is = new FileInputStream(urlOrPath);
                } catch (FileNotFoundException e){
                    // 相对路径读取文件
                    is = FileLoader.class.getClassLoader().getResourceAsStream(urlOrPath);
                }
            }

            baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[MAX_BYTE_SIZE];
            int len;
            while ((len = is.read(buffer)) > -1) {
                baos.write(buffer, 0, len);
            }
            baos.flush();
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(is);
            IOUtils.closeQuietly(baos);
        }
        return null;
    }
}
