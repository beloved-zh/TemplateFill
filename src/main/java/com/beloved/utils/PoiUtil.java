package com.beloved.utils;

import com.beloved.cache.CacheManager;
import com.beloved.core.ImageEntity;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-11 16:11
 * @Description: poi工具类
 */
public class PoiUtil {

    private static String PICXML = ""
            + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
            + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "         <pic:nvPicPr>"
            + "            <pic:cNvPr id=\"%s\" name=\"Generated\"/>"
            + "            <pic:cNvPicPr/>" + "         </pic:nvPicPr>"
            + "         <pic:blipFill>"
            + "            <a:blip r:embed=\"%s\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
            + "            <a:stretch>"
            + "               <a:fillRect/>"
            + "            </a:stretch>" + "         </pic:blipFill>"
            + "         <pic:spPr>" + "            <a:xfrm>"
            + "               <a:off x=\"0\" y=\"0\"/>"
            + "               <a:ext cx=\"%s\" cy=\"%s\"/>"
            + "            </a:xfrm>"
            + "            <a:prstGeom prst=\"rect\">"
            + "               <a:avLst/>" + "            </a:prstGeom>"
            + "         </pic:spPr>" + "      </pic:pic>"
            + "   </a:graphicData>" + "</a:graphic>";
    
    /**
     * 返回流和图片类型
     *
     * @param entity
     * @return (byte[]) isAndType[0],(Integer)isAndType[1]
     * @throws Exception
     * @author JueYue
     * 2013-11-20
     */
    public static Object[] getIsAndType(ImageEntity entity) {
        Object[] result = new Object[2];
        String   type;
        if (entity.getType().equals(ImageEntity.URL)) {
            result[0] = CacheManager.loaderFileByteArray(entity.getUrl());
        } else {
            result[0] = entity.getData();
        }
        type = getFileExtendName((byte[])result[0]);
        result[1] = getImageType(type);
        return result;
    }

    /**
     * 获取文件扩展名
     * @param photoByte
     * @return
     */
    public static String getFileExtendName(byte[] photoByte) {
        String strFileExtendName = "JPG";
        if ((photoByte[0] == 71) && (photoByte[1] == 73) && (photoByte[2] == 70)
                && (photoByte[3] == 56) && ((photoByte[4] == 55) || (photoByte[4] == 57))
                && (photoByte[5] == 97)) {
            strFileExtendName = "GIF";
        } else if ((photoByte[6] == 74) && (photoByte[7] == 70) && (photoByte[8] == 73)
                && (photoByte[9] == 70)) {
            strFileExtendName = "JPG";
        } else if ((photoByte[0] == 66) && (photoByte[1] == 77)) {
            strFileExtendName = "BMP";
        } else if ((photoByte[1] == 80) && (photoByte[2] == 78) && (photoByte[3] == 71)) {
            strFileExtendName = "PNG";
        }
        return strFileExtendName;
    }

    /**
     * 获取图片类型
     * @param type
     * @return
     */
    private static Integer getImageType(String type) {
        if ("JPG".equalsIgnoreCase(type) || "JPEG".equalsIgnoreCase(type)) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        }
        if ("GIF".equalsIgnoreCase(type)) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        if ("BMP".equalsIgnoreCase(type)) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        if ("PNG".equalsIgnoreCase(type)) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        }
        return XWPFDocument.PICTURE_TYPE_JPEG;
    }

    public static void createPicture(XWPFRun run, String blipId, int id, int width, int height) {
        final int emu = 9525;
        width *= emu;
        height *= emu;
        CTInline inline   = run.getCTR().addNewDrawing().addNewInline();
        String   picXml   = String.format(PICXML, id, blipId, width, height);
        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }

    public static void createPicture(XWPFRun run, String blipId, int id, int width, int height, boolean isAbove) {
        createPicture(run, blipId, id, width, height);
        CTDrawing drawing         = run.getCTR().getDrawingArray(0);
        CTGraphicalObject graphicalObject = drawing.getInlineArray(0).getGraphic();
        //拿到新插入的图片替换添加CTAnchor 设置浮动属性 删除inline属性
        CTAnchor anchor = getAnchorWithGraphic(graphicalObject, RandomStringUtils.randomAlphanumeric(10),
                Units.toEMU(width), Units.toEMU(height),//图片大小
                Units.toEMU(50), Units.toEMU(0), isAbove);
        drawing.setAnchorArray(new CTAnchor[]{anchor});
        drawing.removeInline(0);
    }

    /**
     * @param graphicalObject 图片数据
     * @param deskFileName    图片描述
     * @param width           宽
     * @param height          高
     * @param leftOffset      水平偏移 left
     * @param topOffset       垂直偏移 top
     * @param behind          文字上方，文字下方
     * @return
     * @throws Exception
     */
    public static CTAnchor getAnchorWithGraphic(CTGraphicalObject graphicalObject,
                                                String deskFileName, int width, int height,
                                                int leftOffset, int topOffset, boolean behind) {
        String anchorXML =
                "<wp:anchor xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        + "simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"" + ((behind) ? 1 : 0) + "\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
                        + "<wp:simplePos x=\"0\" y=\"0\"/>"
                        + "<wp:positionH relativeFrom=\"column\">"
                        + "<wp:posOffset>" + leftOffset + "</wp:posOffset>"
                        + "</wp:positionH>"
                        + "<wp:positionV relativeFrom=\"paragraph\">"
                        + "<wp:posOffset>" + topOffset + "</wp:posOffset>" +
                        "</wp:positionV>"
                        + "<wp:extent cx=\"" + width + "\" cy=\"" + height + "\"/>"
                        + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
                        + "<wp:wrapNone/>"
                        + "<wp:docPr id=\"1\" name=\"Drawing 0\" descr=\"" + deskFileName + "\"/><wp:cNvGraphicFramePr/>"
                        + "</wp:anchor>";

        CTDrawing drawing = null;
        try {
            drawing = CTDrawing.Factory.parse(anchorXML);
        } catch (XmlException e) {
            e.printStackTrace();
        }
        CTAnchor anchor = drawing.getAnchorArray(0);
        anchor.setGraphic(graphicalObject);
        return anchor;
    }
}
