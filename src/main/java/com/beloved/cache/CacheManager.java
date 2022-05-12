package com.beloved.cache;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-11 17:42
 * @Description: 缓存管理
 */
public class CacheManager {

    /**
     * 线程模板加载参数
     */
    public static final ThreadLocal<FileLoader> LOCAL_TEMPLATE_LOADER;

    static {
        LOCAL_TEMPLATE_LOADER = new ThreadLocal<>();
    }
    
    /**
     * 获取文件
     * @param url 网络地址 或 文件路径（相对路径 | 绝对路径）
     * @return
     * @throws IOException
     */
    public static InputStream loaderFileStream(String url) {
        return new ByteArrayInputStream(loaderFileByteArray(url));
    }

    /**
     * 获取文件
     * @param url 网络地址 或 文件路径（相对路径 | 绝对路径）
     * @return
     * @throws IOException
     */
    public static byte[] loaderFileByteArray(String url) {
        byte[] result;
        //复杂数据,防止操作原数据
        if (LOCAL_TEMPLATE_LOADER.get() != null) {
            result = LOCAL_TEMPLATE_LOADER.get().loaderFile(url);
        } else {
            FileLoader fileLoader = new FileLoader();
            LOCAL_TEMPLATE_LOADER.set(fileLoader);
            result = fileLoader.loaderFile(url);
        }
        result = Arrays.copyOf(result, result.length);
        return result;
    }
}
