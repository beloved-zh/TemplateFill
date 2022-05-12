package com.beloved.utils;

import com.google.zxing.*;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.common.HybridBinarizer;
import com.google.zxing.oned.Code128Writer;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-05 10:30
 * @Description: 条形码工具类
 */
public class BarCodeUtil {

    // 默认图片宽度
    private static final int DEFAULT_PICTURE_WIDTH = 200;

    // 默认图片高度
    private static final int DEFAULT_PICTURE_HEIGHT = 150;

    // 默认条形码宽度
    private static final int DEFAULT_WIDTH = 200;

    // 默认条形码高度
    private static final int DEFAULT_HEIGHT = 100;

    // 默认字体大小
    private static final int DEFAULT_FONT_SIZE = 15;

    private static final String FORMAT = "png";
    private static final String ENCODE = "UTF-8";

    /**
     * 解析条形码
     * @param imgPath  图片地址路径
     *            注意：无法解析特殊字符
     * @return         解析内容
     * @throws IOException
     * @throws NotFoundException
     */
    public static String decodeBarCode(String imgPath) throws IOException, NotFoundException {

        BufferedImage image = ImageIO.read(new File(imgPath));

        LuminanceSource source = new BufferedImageLuminanceSource(image);
        BinaryBitmap bitmap = new BinaryBitmap(new HybridBinarizer(source));

        Map<DecodeHintType, Object> hints = new Hashtable<>();
        hints.put(DecodeHintType.CHARACTER_SET, ENCODE);

        Result result = new MultiFormatReader().decode(bitmap, hints);

        return result.getText();
    }

    /**
     * 生成 默认宽高 Base64 条形码
     * @param content 内容
     * @return  Base64
     * @throws WriterException
     * @throws IOException
     */
    public static String getBarCodeBase64(String content) throws WriterException, IOException {
        return getBarCodeBase64(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 生成 Base64 条形码
     * @param content 内容
     * @param width   宽
     * @param height    高
     * @return    Base64
     * @throws WriterException
     * @throws IOException
     */
    public static String getBarCodeBase64(String content, int width, int height) throws WriterException, IOException {
        return new String(Base64.decodeBase64(getBarCodeByteArray(content, width, height)));
    }

    /**
     * 生成 byte[] 条形码
     * @param content 内容
     * @return    byte[]
     * @throws WriterException
     * @throws IOException
     */
    public static byte[] getBarCodeByteArray(String content) throws WriterException, IOException {
        return getBarCodeByteArray(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 生成 byte[] 条形码
     * @param content 内容
     * @param width   宽
     * @param height    高
     * @return    byte[]
     * @throws WriterException
     * @throws IOException
     */
    public static byte[] getBarCodeByteArray(String content, int width, int height) throws WriterException, IOException {

        BufferedImage bufferedImage = getBarCode(content, width, height);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bufferedImage, FORMAT, baos);

        return baos.toByteArray();
    }


    /**
     * 生成默认宽高条形码
     * @param content   内容
     * @param imgPath   保存地址
     * @throws WriterException
     * @throws IOException
     */
    public static void getBarCode(String content, String imgPath) throws WriterException, IOException {
        getBarCode(content, DEFAULT_WIDTH, DEFAULT_HEIGHT, imgPath);
    }

    /**
     * 生成条形码
     * @param content   内容
     * @param width     宽
     * @param height    高
     * @param imgPath   保存地址
     * @throws WriterException
     * @throws IOException
     */
    public static void getBarCode(String content, int width, int height, String imgPath) throws WriterException, IOException {
        BufferedImage bufferedImage = getBarCode(content, width, height);

        File saveFile = new File(imgPath);

        // 是否有父级目录没有则创建
        if(!saveFile.getParentFile().exists()){
            saveFile.getParentFile().mkdirs();
        }

        ImageIO.write(bufferedImage, FORMAT, saveFile);
    }

    /**
     * 生成默认宽高条形码
     * @param content   内容
     */
    private static BufferedImage getBarCode(String content) throws WriterException {
        return getBarCode(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 生成条形码
     * @param content   内容
     * @param width     宽
     * @param height    高
     */
    private static BufferedImage getBarCode(String content, int width, int height) throws WriterException {
        if (StringUtils.isEmpty(content)) {
            return null;
        }

        // 设置参数
        HashMap<EncodeHintType, Comparable> hints = new HashMap<>();
        hints.put(EncodeHintType.CHARACTER_SET, ENCODE);
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
        hints.put(EncodeHintType.MARGIN, 0);

        Code128Writer writer = new Code128Writer();

        // 编码内容, 编码类型, 宽度, 高度, 设置参数
        BitMatrix bitMatrix = writer.encode(content, BarcodeFormat.CODE_128, width, height, hints);

        BufferedImage bufferedImage = MatrixToImageWriter.toBufferedImage(bitMatrix);

        return bufferedImage;
    }

    /**
     * 生成带文本内容 Base64  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @return Base64
     */
    public static String getBarCodeWordsBase64(String codeValue, String bottomStr) throws IOException, WriterException {
        return getBarCodeWordsBase64(codeValue, bottomStr, "", "");
    }

    /**
     * 生成带文本内容 Base64  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @return Base64
     */
    public static String getBarCodeWordsBase64(String codeValue,
                                               String bottomStr,
                                               String topLeftStr,
                                               String topRightStr) throws IOException, WriterException {
        return getBarCodeWordsBase64(
                codeValue,
                bottomStr,
                topLeftStr,
                topRightStr,
                DEFAULT_PICTURE_WIDTH,
                DEFAULT_PICTURE_HEIGHT,
                0,
                0,
                10,
                0,
                -10,
                0,
                DEFAULT_FONT_SIZE);
    }

    /**
     * 生成带文本内容 Base64  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @param pictureWidth    图片宽度
     * @param pictureHeight   图片高度
     * @param codeOffsetX     条形码宽度
     * @param codeOffsetY     条形码高度
     * @param topLeftOffsetX  左上角文字X轴偏移量
     * @param topLeftOffsetY  左上角文字Y轴偏移量
     * @param topRightOffsetX 右上角文字X轴偏移量
     * @param topRightOffsetY 右上角文字Y轴偏移量
     * @param fontSize        字体大小
     * @return Base64
     */
    public static String getBarCodeWordsBase64(String codeValue,
                                       String bottomStr,
                                       String topLeftStr,
                                       String topRightStr,
                                       int pictureWidth,
                                       int pictureHeight,
                                       int codeOffsetX,
                                       int codeOffsetY,
                                       int topLeftOffsetX,
                                       int topLeftOffsetY,
                                       int topRightOffsetX,
                                       int topRightOffsetY,
                                       int fontSize) throws IOException, WriterException {
        return new String(Base64.decodeBase64(getBarCodeWordsByteArray(codeValue, bottomStr, topLeftStr, topRightStr, pictureWidth, pictureHeight, codeOffsetX, codeOffsetY, topLeftOffsetX, topLeftOffsetY, topRightOffsetX, topRightOffsetY, fontSize)));
    }

    /**
     * 生成带文本内容 byte[]  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @return byte[]
     */
    public static byte[] getBarCodeWordsByteArray(String codeValue, String bottomStr) throws IOException, WriterException {
        return getBarCodeWordsByteArray(codeValue, bottomStr, "", "");
    }

    /**
     * 生成带文本内容 byte[]  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @return byte[]
     */
    public static byte[] getBarCodeWordsByteArray(String codeValue,
                                                  String bottomStr,
                                                  String topLeftStr,
                                                  String topRightStr) throws IOException, WriterException {
        return getBarCodeWordsByteArray(
                codeValue,
                bottomStr,
                topLeftStr,
                topRightStr,
                DEFAULT_PICTURE_WIDTH,
                DEFAULT_PICTURE_HEIGHT,
                0,
                0,
                10,
                0,
                -10,
                0,
                DEFAULT_FONT_SIZE);
    }

    /**
     * 生成带文本内容 byte[]  条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @param pictureWidth    图片宽度
     * @param pictureHeight   图片高度
     * @param codeOffsetX     条形码宽度
     * @param codeOffsetY     条形码高度
     * @param topLeftOffsetX  左上角文字X轴偏移量
     * @param topLeftOffsetY  左上角文字Y轴偏移量
     * @param topRightOffsetX 右上角文字X轴偏移量
     * @param topRightOffsetY 右上角文字Y轴偏移量
     * @param fontSize        字体大小
     * @return byte[]
     */
    public static byte[] getBarCodeWordsByteArray(String codeValue,
                                                  String bottomStr,
                                                  String topLeftStr,
                                                  String topRightStr,
                                                  int pictureWidth,
                                                  int pictureHeight,
                                                  int codeOffsetX,
                                                  int codeOffsetY,
                                                  int topLeftOffsetX,
                                                  int topLeftOffsetY,
                                                  int topRightOffsetX,
                                                  int topRightOffsetY,
                                                  int fontSize) throws IOException, WriterException {
        BufferedImage bufferedImage = getBarCodeWords(codeValue, bottomStr, topLeftStr, topRightStr, pictureWidth, pictureHeight, codeOffsetX, codeOffsetY, topLeftOffsetX, topLeftOffsetY, topRightOffsetX, topRightOffsetY, fontSize);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bufferedImage, FORMAT, baos);

        return baos.toByteArray();
    }

    /**
     * 生成带文本内容条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param imgPath         保存图片地址
     */
    public static void getBarCodeWords(String codeValue, String bottomStr, String imgPath) throws IOException, WriterException {
        getBarCodeWords(codeValue, bottomStr, "", "", imgPath);
    }

    /**
     * 生成带文本内容条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @param imgPath         保存图片地址
     */
    public static void getBarCodeWords(String codeValue,
                                        String bottomStr,
                                        String topLeftStr,
                                        String topRightStr,
                                        String imgPath) throws IOException, WriterException {
        getBarCodeWords(
                codeValue,
                bottomStr,
                topLeftStr,
                topRightStr,
                DEFAULT_PICTURE_WIDTH,
                DEFAULT_PICTURE_HEIGHT,
                0,
                0,
                10,
                0,
                -10,
                0,
                DEFAULT_FONT_SIZE,
                imgPath);
    }

    /**
     * 生成带文本内容条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @param pictureWidth    图片宽度
     * @param pictureHeight   图片高度
     * @param codeOffsetX     条形码宽度
     * @param codeOffsetY     条形码高度
     * @param topLeftOffsetX  左上角文字X轴偏移量
     * @param topLeftOffsetY  左上角文字Y轴偏移量
     * @param topRightOffsetX 右上角文字X轴偏移量
     * @param topRightOffsetY 右上角文字Y轴偏移量
     * @param fontSize        字体大小
     * @param imgPath         保存图片地址
     */
    public static void getBarCodeWords(String codeValue,
                                        String bottomStr,
                                        String topLeftStr,
                                        String topRightStr,
                                        int pictureWidth,
                                        int pictureHeight,
                                        int codeOffsetX,
                                        int codeOffsetY,
                                        int topLeftOffsetX,
                                        int topLeftOffsetY,
                                        int topRightOffsetX,
                                        int topRightOffsetY,
                                        int fontSize,
                                        String imgPath) throws IOException, WriterException {
        BufferedImage bufferedImage = getBarCodeWords(codeValue, bottomStr, topLeftStr, topRightStr, pictureWidth, pictureHeight, codeOffsetX, codeOffsetY, topLeftOffsetX, topLeftOffsetY, topRightOffsetX, topRightOffsetY, fontSize);

        File saveFile = new File(imgPath);

        // 是否有父级目录没有则创建
        if(!saveFile.getParentFile().exists()){
            saveFile.getParentFile().mkdirs();
        }

        ImageIO.write(bufferedImage, FORMAT, saveFile);
    }

    /**
     * 生成带文本内容条形码
     *
     * @param codeValue       条形码内容
     * @param bottomStr       底部文字
     * @param topLeftStr      左上角文字
     * @param topRightStr     右上角文字
     * @param pictureWidth    图片宽度
     * @param pictureHeight   图片高度
     * @param codeOffsetX     条形码宽度
     * @param codeOffsetY     条形码高度
     * @param topLeftOffsetX  左上角文字X轴偏移量
     * @param topLeftOffsetY  左上角文字Y轴偏移量
     * @param topRightOffsetX 右上角文字X轴偏移量
     * @param topRightOffsetY 右上角文字Y轴偏移量
     * @param fontSize        字体大小
     * @return 条形码图片
     */
    private static BufferedImage getBarCodeWords(String codeValue,
                                                 String bottomStr,
                                                 String topLeftStr,
                                                 String topRightStr,
                                                 int pictureWidth,
                                                 int pictureHeight,
                                                 int codeOffsetX,
                                                 int codeOffsetY,
                                                 int topLeftOffsetX,
                                                 int topLeftOffsetY,
                                                 int topRightOffsetX,
                                                 int topRightOffsetY,
                                                 int fontSize) throws WriterException {
        BufferedImage codeImage = getBarCode(codeValue);
        BufferedImage picImage = new BufferedImage(pictureWidth, pictureHeight, BufferedImage.TYPE_INT_RGB);

        Graphics2D g2d = picImage.createGraphics();
        // 抗锯齿
        setGraphics2D(g2d);
        // 设置白色
        setColorWhite(g2d, picImage.getWidth(), picImage.getHeight());

        // 条形码默认居中显示
        int codeStartX = (pictureWidth - codeImage.getWidth()) / 2 + codeOffsetX;
        int codeStartY = (pictureHeight - codeImage.getHeight()) / 2 + codeOffsetY;
        // 画条形码到新的面板
        g2d.drawImage(codeImage, codeStartX, codeStartY, codeImage.getWidth(), codeImage.getHeight(), null);

        // 画文字到新的面板
        g2d.setColor(Color.BLACK);
        // 字体、字型、字号
        g2d.setFont(new Font("微软雅黑", Font.PLAIN, fontSize));
        // 文字与条形码之间的间隔
        int wordAndCodeSpacing = 3;

        if (StringUtils.isNotEmpty(bottomStr)) {
            // 文字长度
            int strWidth = g2d.getFontMetrics().stringWidth(bottomStr);
            // 文字X轴开始坐标，这里是居中
            int strStartX = codeStartX + (codeImage.getWidth() - strWidth) / 2;
            // 文字Y轴开始坐标
            int strStartY = codeStartY + codeImage.getHeight() + fontSize + wordAndCodeSpacing;
            // 画文字
            g2d.drawString(bottomStr, strStartX, strStartY);
        }

        if (StringUtils.isNotEmpty(topLeftStr)) {
            // 文字长度
            int strWidth = g2d.getFontMetrics().stringWidth(topLeftStr);
            // 文字X轴开始坐标
            int strStartX = codeStartX + topLeftOffsetX;
            // 文字Y轴开始坐标
            int strStartY = codeStartY + topLeftOffsetY - wordAndCodeSpacing;
            // 画文字
            g2d.drawString(topLeftStr, strStartX, strStartY);
        }

        if (StringUtils.isNotEmpty(topRightStr)) {
            // 文字长度
            int strWidth = g2d.getFontMetrics().stringWidth(topRightStr);
            // 文字X轴开始坐标，这里是居中
            int strStartX = codeStartX + codeImage.getWidth() - strWidth + topRightOffsetX;
            // 文字Y轴开始坐标
            int strStartY = codeStartY + topRightOffsetY - wordAndCodeSpacing;
            // 画文字
            g2d.drawString(topRightStr, strStartX, strStartY);
        }

        g2d.dispose();
        picImage.flush();

        return picImage;
    }

    /**
     * 设置 Graphics2D 属性  （抗锯齿）
     *
     * @param g2d Graphics2D提供对几何形状、坐标转换、颜色管理和文本布局更为复杂的控制
     */
    private static void setGraphics2D(Graphics2D g2d) {
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_DEFAULT);
        Stroke s = new BasicStroke(1, BasicStroke.CAP_ROUND, BasicStroke.JOIN_MITER);
        g2d.setStroke(s);
    }

    /**
     * 设置背景为白色
     *
     * @param g2d Graphics2D提供对几何形状、坐标转换、颜色管理和文本布局更为复杂的控制
     */
    private static void setColorWhite(Graphics2D g2d, int width, int height) {
        g2d.setColor(Color.WHITE);
        //填充整个屏幕
        g2d.fillRect(0, 0, width, height);
        //设置笔刷
        g2d.setColor(Color.BLACK);
    }
}

