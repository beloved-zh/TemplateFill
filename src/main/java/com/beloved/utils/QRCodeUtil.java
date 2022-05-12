package com.beloved.utils;

import com.google.zxing.*;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.common.HybridBinarizer;
import com.google.zxing.qrcode.QRCodeWriter;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-05 10:29
 * @Description: 二维码工具类
 */
public class QRCodeUtil {

    // 二维码默认宽度，单位像素
    private static final int DEFAULT_WIDTH = 400;
    // 二维码默认高度，单位像素
    private static final int DEFAULT_HEIGHT = 400;

    private static final String FORMAT = "png";
    private static final String ENCODE = "UTF-8";

    /**
     * 解析二维码
     *
     * @param imgPath   二维码图片全路径
     * @return          解析内容
     */
    public static String decodeQRCode(String imgPath) throws IOException, NotFoundException {

        BufferedImage image = ImageIO.read(new File(imgPath));
        LuminanceSource source = new BufferedImageLuminanceSource(image);
        BinaryBitmap bitmap = new BinaryBitmap(new HybridBinarizer(source));

        Map<DecodeHintType, Object> hints = new Hashtable<>();
        hints.put(DecodeHintType.CHARACTER_SET, ENCODE);

        Result result = new MultiFormatReader().decode(bitmap, hints);

        return result.getText();
    }

    /**
     * 获取默认宽高二维码
     * @param content 内容
     * @return        Base64编码图片
     * @throws WriterException
     * @throws IOException
     */
    public static String getQRCodeBase64(String content) throws WriterException, IOException {
        return getQRCodeBase64(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 获取二维码
     * @param content 内容
     * @param width   宽度
     * @param height  高度
     * @return        Base64编码图片
     * @throws WriterException
     */
    public static String getQRCodeBase64(String content, int width, int height) throws WriterException, IOException {
        return new String(Base64.decodeBase64(getQRCodeByteArray(content, width, height)));
    }

    /**
     * 获取默认宽高二维码
     * @param content 内容
     * @return        byte[] 编码图片
     * @throws WriterException
     * @throws IOException
     */
    public static byte[] getQRCodeByteArray(String content) throws WriterException, IOException {
        return getQRCodeByteArray(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 获取二维码
     * @param content 内容
     * @param width   宽度
     * @param height  高度
     * @return        byte[] 编码图片
     * @throws WriterException
     */
    public static byte[] getQRCodeByteArray(String content, int width, int height) throws WriterException, IOException {

        BufferedImage bufferedImage = getQRCode(content, width, height);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bufferedImage, FORMAT, baos);

        return baos.toByteArray();
    }

    /**
     * 获取默认宽高二维码
     * @param content   内容
     * @param imgPath   生成图片全路径
     * @throws WriterException
     * @throws IOException
     */
    public static void getQRCode(String content, String imgPath) throws WriterException, IOException {
        getQRCode(content, DEFAULT_WIDTH, DEFAULT_HEIGHT, imgPath);
    }

    /**
     * 获取二维码
     * @param content   内容
     * @param width     宽度
     * @param height    高度
     * @param imgPath   生成图片全路径
     * @throws WriterException
     * @throws IOException
     */
    public static void getQRCode(String content, int width, int height, String imgPath) throws WriterException, IOException {
        BufferedImage bufferedImage = getQRCode(content, width, height);

        File saveFile = new File(imgPath);

        // 是否有父级目录没有则创建
        if(!saveFile.getParentFile().exists()){
            saveFile.getParentFile().mkdirs();
        }

        ImageIO.write(bufferedImage, FORMAT, saveFile);
    }

    /**
     * 获取默认宽高二维码
     * @param content 内容
     * @return
     * @throws WriterException
     */
    public static BufferedImage getQRCode(String content) throws WriterException {
        return getQRCode(content, DEFAULT_WIDTH, DEFAULT_HEIGHT);
    }

    /**
     * 获取二维码
     * @param content 内容
     * @param width   宽度
     * @param height  高度
     * @return
     * @throws WriterException
     */
    public static BufferedImage getQRCode(String content, int width, int height) throws WriterException {
        if (StringUtils.isEmpty(content)) {
            return null;
        }

        // 设置二维码参数
        HashMap<EncodeHintType, Comparable> hints = new HashMap<>();

        // 设置字符编码类型
        hints.put(EncodeHintType.CHARACTER_SET, ENCODE);
        // 设置误差校正
        // ErrorCorrectionLevel：误差校正等级，L = ~7% correction、M = ~15% correction、Q = ~25% correction、H = ~30% correction
        // 不设置时，默认为 L 等级，等级不一样，生成的图案不同，但扫描的结果是一样的
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);

        // 设置二维码边距，单位像素，值越小，二维码距离四周越近
        hints.put(EncodeHintType.MARGIN, 0);

        QRCodeWriter writer = new QRCodeWriter();
        BitMatrix bitMatrix = writer.encode(content, BarcodeFormat.QR_CODE, width, height, hints);

        BufferedImage bufferedImage = MatrixToImageWriter.toBufferedImage(bitMatrix);

        return bufferedImage;
    }
}
