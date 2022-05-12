package com.beloved.core;

import com.beloved.cache.WordCache;
import com.beloved.utils.PoiUtil;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.lang.reflect.Field;
import java.util.*;

/**
 * @Author: Beloved
 * @CreateTime: 2022-05-11 17:53
 * @Description: word操作核心
 */
public class WordCore {

    public static XWPFDocument fillWord(String url, Map<String, Object> params) throws Exception {
        XWPFDocument document = WordCache.getXWPFDocument(url);
        parseWordSetValue(document, params);
        return document;
    }

    
    public static void fillWord(XWPFDocument document, Map<String, Object> params) throws Exception {
        parseWordSetValue(document, params);
    }

    /**
     * 解析 word 填充值
     * @param document
     * @param params
     * @throws Exception
     */
    private static void parseWordSetValue(XWPFDocument document, Map<String, Object> params) throws Exception {
        //填充段落
        fillAllParagraph(document.getParagraphs(), params);
        //填充页眉
        fillHeaderAndFoot(document, params);
        //填充表格
        fillAllTable(document.getTablesIterator(), params);
    }

    /**
     * 填充页眉页脚
     * @param document
     * @param params
     */
    private static void fillHeaderAndFoot(XWPFDocument document, Map<String, Object> params) {
        List<XWPFHeader> headerList = document.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            for (XWPFParagraph paragraph : xwpfHeader.getListParagraph()) {
                fillThisParagraph(paragraph, params);
            }
        }
        List<XWPFFooter> footerList = document.getFooterList();
        for (XWPFFooter xwpfFooter : footerList) {
            for (XWPFParagraph paragraph : xwpfFooter.getListParagraph()) {
                fillThisParagraph(paragraph, params);
            }
        }
    }

    /**
     * 填充所有表格
     * @param itTable
     * @param params
     */
    private static void fillAllTable(Iterator<XWPFTable> itTable, Map<String, Object> params) {
        while (itTable.hasNext()) {
            XWPFTable table = itTable.next();
            if (table.getText().contains(ElCore.START_LABEL)) {
                parseThisTable(table, params);
            }
        }
    }

    /**
     * 解析表格
     * @param table
     * @param params
     */
    private static void parseThisTable(XWPFTable table, Map<String, Object> params) {
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            XWPFTableRow row = table.getRow(i);
            Object object = checkThisTableIsNeedIterator(row.getCell(0), params);
            
            if (ObjectUtils.isEmpty(object)) {
                parseThisRow(row, params);
            }else {
                fillNextRowAndAddRow(table, i, (List<Object>)object);
                i = i + ((List<Object>)object).size() - 1;
            }
        }
    }

    /**
     * 遍历填充下一行
     * @param table
     * @param index
     * @param list
     */
    private static void fillNextRowAndAddRow(XWPFTable table, int index, List<Object> list) {
        List<XWPFTableCell> tempCellList = table.getRow(index).getTableCells();
        XWPFTableRow currentRow = table.getRow(index);

        List<String> newRowLabel = parseCurrentRowGetParams(currentRow);

        String listName = newRowLabel.get(0).replace(ElCore.FOREACH, ElCore.EMPTY);
        String[] keys = listName.replaceAll("\\s{1,}", " ").trim().split(" ");
        newRowLabel.set(0, keys[1]);

        for (Object obj : list) {
            currentRow = table.insertNewTableRow(index++);
            
            for (int cellIndex = 0; cellIndex < newRowLabel.size(); cellIndex++) {
                Object fillValue = getCellMapValue(newRowLabel.get(cellIndex), obj);

                if (fillValue instanceof ImageEntity) {
                    cellAddImage(currentRow.createCell(), (ImageEntity)fillValue);
                } else {
                    copyCellAndSetValue(tempCellList.get(cellIndex), currentRow.createCell(), fillValue.toString());
                }
            }
        }

        table.removeRow(index);
    }

    /**
     * 单元格设置图片
     * @param cell
     * @param imageEntity
     */
    private static void cellAddImage(XWPFTableCell cell, ImageEntity imageEntity) {
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        XWPFRun run = paragraph.createRun();
        addImage(run, imageEntity);
    }

    /**
     * 复制单元格并设置文本信息
     * @param tmpCell
     * @param cell
     * @param text
     */
    private static void copyCellAndSetValue(XWPFTableCell tmpCell, XWPFTableCell cell, String text) {
        CTTc cttc2 = tmpCell.getCTTc();
        CTTcPr ctPr2 = cttc2.getTcPr();
        cell.getTableRow().setHeight(tmpCell.getTableRow().getHeight());
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        if (tmpCell.getColor() != null) {
            cell.setColor(tmpCell.getColor());
        }
        if (tmpCell.getVerticalAlignment() != null) {
            cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
        }
        if (ctPr2.getTcW() != null) {
            ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
        }
        if (ctPr2.getVAlign() != null) {
            ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
        }
        if (cttc2.getPList().size() > 0) {
            CTP ctp = cttc2.getPList().get(0);
            if (ctp.getPPr() != null) {
                if (ctp.getPPr().getJc() != null) {
                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(ctp.getPPr().getJc().getVal());
                }
            }
        }

        if (ctPr2.getTcBorders() != null) {
            ctPr.setTcBorders(ctPr2.getTcBorders());
        }

        XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
        XWPFParagraph cellP = cell.getParagraphs().get(0);
        XWPFRun tmpR = null;
        if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
            tmpR = tmpP.getRuns().get(0);
        }

        XWPFRun cellR = cellP.createRun();
        cellR.setText(text);
        //复制字体信息
        if (tmpR != null) {
            cellR.setBold(tmpR.isBold());
            cellR.setItalic(tmpR.isItalic());
            cellR.setStrike(tmpR.isStrike());
            cellR.setUnderline(tmpR.getUnderline());
            cellR.setColor(tmpR.getColor());
            cellR.setTextPosition(tmpR.getTextPosition());
            if (tmpR.getFontSize() != -1) {
                cellR.setFontSize(tmpR.getFontSize());
            }
            if (tmpR.getFontFamily() != null) {
                cellR.setFontFamily(tmpR.getFontFamily());
            }
            if (tmpR.getCTR() != null) {
                if (tmpR.getCTR().isSetRPr()) {
                    CTRPr tmpRPr = tmpR.getCTR().getRPr();
                    if (tmpRPr.isSetRFonts()) {
                        CTFonts tmpFonts = tmpRPr.getRFonts();
                        CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR.getCTR().getRPr() : cellR.getCTR().addNewRPr();
                        CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr.getRFonts() : cellRPr.addNewRFonts();
                        cellFonts.setAscii(tmpFonts.getAscii());
                        cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
                        cellFonts.setCs(tmpFonts.getCs());
                        cellFonts.setCstheme(tmpFonts.getCstheme());
                        cellFonts.setEastAsia(tmpFonts.getEastAsia());
                        cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
                        cellFonts.setHAnsi(tmpFonts.getHAnsi());
                        cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
                    }
                }
            }
        }
        //复制段落信息
        if (tmpP.getAlignment() != null) {
            cellP.setAlignment(tmpP.getAlignment());
        }
        if (tmpP.getVerticalAlignment() != null) {
            cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
        }
        if (tmpP.getBorderBetween() != null) {
            cellP.setBorderBetween(tmpP.getBorderBetween());
        }
        if (tmpP.getBorderBottom() != null){
            cellP.setBorderBottom(tmpP.getBorderBottom());
        }
        if (tmpP.getBorderLeft() != null){
            cellP.setBorderLeft(tmpP.getBorderLeft());
        }
        if (tmpP.getBorderRight() != null){
            cellP.setBorderRight(tmpP.getBorderRight());
        }
        if (tmpP.getBorderTop() != null){
            cellP.setBorderTop(tmpP.getBorderTop());
        }
        cellP.setPageBreak(tmpP.isPageBreak());
        if (tmpP.getCTP() != null) {
            if (tmpP.getCTP().getPPr() != null) {
                CTPPr tmpPPr = tmpP.getCTP().getPPr();
                CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP.getCTP().getPPr() : cellP.getCTP().addNewPPr();
                //复制段落间距信息
                CTSpacing tmpSpacing = tmpPPr.getSpacing();
                if (tmpSpacing != null) {
                    CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr.getSpacing() : cellPPr.addNewSpacing();
                    if (tmpSpacing.getAfter() != null) {
                        cellSpacing.setAfter(tmpSpacing.getAfter());
                    }
                    if (tmpSpacing.getAfterAutospacing() != null) {
                        cellSpacing.setAfterAutospacing(tmpSpacing.getAfterAutospacing());
                    }
                    if (tmpSpacing.getAfterLines() != null) {
                        cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
                    }
                    if (tmpSpacing.getBefore() != null) {
                        cellSpacing.setBefore(tmpSpacing.getBefore());
                    }
                    if (tmpSpacing.getBeforeAutospacing() != null) {
                        cellSpacing.setBeforeAutospacing(tmpSpacing.getBeforeAutospacing());
                    }
                    if (tmpSpacing.getBeforeLines() != null) {
                        cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
                    }
                    if (tmpSpacing.getLine() != null) {
                        cellSpacing.setLine(tmpSpacing.getLine());
                    }
                    if (tmpSpacing.getLineRule() != null) {
                        cellSpacing.setLineRule(tmpSpacing.getLineRule());
                    }
                }
                //复制段落缩进信息
                CTInd tmpInd = tmpPPr.getInd();
                if (tmpInd != null) {
                    CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd() : cellPPr.addNewInd();
                    if (tmpInd.getFirstLine() != null) {
                        cellInd.setFirstLine(tmpInd.getFirstLine());
                    }
                    if (tmpInd.getFirstLineChars() != null) {
                        cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
                    }
                    if (tmpInd.getHanging() != null) {
                        cellInd.setHanging(tmpInd.getHanging());
                    }
                    if (tmpInd.getHangingChars() != null) {
                        cellInd.setHangingChars(tmpInd.getHangingChars());
                    }
                    if (tmpInd.getLeft() != null) {
                        cellInd.setLeft(tmpInd.getLeft());
                    }
                    if (tmpInd.getLeftChars() != null) {
                        cellInd.setLeftChars(tmpInd.getLeftChars());
                    }
                    if (tmpInd.getRight() != null) {
                        cellInd.setRight(tmpInd.getRight());
                    }
                    if (tmpInd.getRightChars() != null) {
                        cellInd.setRightChars(tmpInd.getRightChars());
                    }
                }
            }
        }
    }
    
    /**
     * 获取表格单元格字段映射值
     * @param fieldName 字段参数
     * @param object 填充值
     * @return 映射值
     */
    private static Object getCellMapValue(String fieldName, Object object) {
        try {
            if (object instanceof Map) {
                return ((Map<String, Object>) object).get(fieldName) == null ? "" : ((Map<String, Object>) object).get(fieldName);
            }
            Class<?> clazz = object.getClass();
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            return field.get(object);
        }catch (Exception e){
            return "";
        }
    }

    /**
     * 解析行参数
     * @param currentRow 需要解析的行
     * @return 行参数
     */
    private static List<String> parseCurrentRowGetParams(XWPFTableRow currentRow) {
        List<String> newRow = new ArrayList<>();

        List<XWPFTableCell> cells = currentRow.getTableCells();

        for (XWPFTableCell cell : cells) {
            String text = cell.getText();
            
            if (StringUtils.isEmpty(text)) {
                newRow.add(ElCore.EMPTY);
            } else {
                text = text.replace(ElCore.START_LABEL, ElCore.EMPTY).replace(ElCore.END_LABEL, ElCore.EMPTY).trim();
                newRow.add(text);
            }
        }
        
        return newRow;
    }

    /**
     * 填充行
     * @param row
     * @param params
     */
    private static void parseThisRow(XWPFTableRow row, Map<String, Object> params) {
        for (XWPFTableCell cell : row.getTableCells()) {
            fillAllParagraph(cell.getParagraphs(), params);
        }
    }

    /**
     * 判断是不是迭代输出
     * @param cell
     * @param params
     * @return
     * @throws Exception
     */
    private static Object checkThisTableIsNeedIterator(XWPFTableCell cell, Map<String, Object> params) {
        String text = cell.getText().trim();
        // 判断是不是迭代输出
        if (StringUtils.isNotEmpty(text) && text.contains(ElCore.FOREACH) && text.startsWith(ElCore.START_LABEL)) {
            text = text.replace(ElCore.FOREACH, ElCore.EMPTY).replace(ElCore.START_LABEL, ElCore.EMPTY);
            String[] keys = text.replaceAll("\\s{1,}", " ").trim().split(" ");
            Object result = params.get(keys[0]);
            return ObjectUtils.isNotEmpty(result) ? result : new ArrayList<Map<String, Object>>(0);
        }
        return null;
    }

    /**
     * 填充所有段落
     * @param paragraphs
     * @param params
     */
    private static void fillAllParagraph(List<XWPFParagraph> paragraphs, Map<String, Object> params) {
        for (XWPFParagraph paragraph : paragraphs) {
            if (paragraph.getText().contains(ElCore.START_LABEL)) {
                fillThisParagraph(paragraph, params);
            }
        }
    }

    /**
     * 填充段落
     * @param paragraph
     * @param params
     */
    private static void fillThisParagraph(XWPFParagraph paragraph, Map<String, Object> params) {

        // 拿到的第一个run, 用来set值, 填充的格式已此格式为准
        XWPFRun currentRun = null;
        
        // 存储除第一个遇到的run, 把他们置空
        List<XWPFRun> runList = new ArrayList<XWPFRun>();
        
        // 存放当前的text
        String currentText = "";
        
        // 判断是不是已经遇到开始标签
        boolean ifFind = false;
        

        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            // 如果为空或者 "" 跳过当前循环
            if (StringUtils.isEmpty(text)) {
                continue;
            }
            if (ifFind) {
                currentText += text;
                if (!currentText.contains(ElCore.START_LABEL)) {
                    ifFind = false;
                    runList.clear();
                } else {
                    runList.add(run);
                }
                if (currentText.contains(ElCore.END_LABEL)) {
                    fillValue(currentRun, currentText, runList, params);
                    runList.clear();
                    currentText = "";
                    ifFind = false;
                }
            } else if(text.contains(ElCore.START_LABEL) || ElCore.START_LABEL.contains(text)) {
                currentText = text;
                ifFind = true;
                currentRun = run;
            } else {
                currentText = "";
            }
            if (currentText.contains(ElCore.END_LABEL)) {
                fillValue(currentRun, currentText, runList, params);
                runList.clear();
                currentText = "";
                ifFind = false;
            }
        }
    }

    /**
     * 填充数据
     * @param currentRun
     * @param currentText
     * @param runList
     * @param params
     */
    private static void fillValue(XWPFRun currentRun, String currentText, List<XWPFRun> runList, Map<String, Object> params) {
        Object obj = getRealValue(currentText, params);
        if (obj instanceof ImageEntity) {
            currentRun.setText("", 0);
            addImage(currentRun, (ImageEntity) obj);
        } else {
            currentText = obj.toString();
            setText(currentRun, currentText);
        }
        for (XWPFRun run : runList) {
            run.setText("", 0);
        }
    }

    /**
     * 添加图片
     * @param currentRun
     * @param obj
     */
    private static void addImage(XWPFRun currentRun, ImageEntity obj) {
        try {
            Object[] isAndType = PoiUtil.getIsAndType(obj);
            String picId = currentRun.getDocument().addPictureData((byte[]) isAndType[0], (Integer) isAndType[1]);
            if (obj.getLocationType() == ImageEntity.EMBED) {
                PoiUtil.createPicture(currentRun, picId, currentRun.getDocument().getNextPicNameNumber((Integer) isAndType[1]), obj.getWidth(), obj.getHeight());
            } else if (obj.getLocationType() == ImageEntity.ABOVE) {
                PoiUtil.createPicture(currentRun, picId, currentRun.getDocument().getNextPicNameNumber((Integer) isAndType[1]), obj.getWidth(), obj.getHeight(), false);
            }  else if (obj.getLocationType() == ImageEntity.BEHIND) {
                PoiUtil.createPicture(currentRun, picId, currentRun.getDocument().getNextPicNameNumber((Integer) isAndType[1]), obj.getWidth(), obj.getHeight(), true);
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        
    }

    /**
     * 设置文本
     * @param currentRun
     * @param currentText
     */
    private static void setText(XWPFRun currentRun, String currentText) {
        if (StringUtils.isNotEmpty(currentText)) {
            String[] tempArr = currentText.split("\r\n");
            for (int i = 0, le = tempArr.length - 1; i < le; i++) {
                currentRun.setText(tempArr[i], i);
                currentRun.addBreak();
            }
            currentRun.setText(tempArr[tempArr.length - 1], tempArr.length - 1);
        } else {
            //对blank字符串做处理，避免显示"{{"
            currentRun.setText("", 0);
        }
    }

    /**
     * 解析当前文本
     * @param currentText
     * @param params
     * @return
     */
    private static Object getRealValue(String currentText, Map<String, Object> params) {
        String label = StringUtils.substringBetween(currentText, ElCore.START_LABEL, ElCore.END_LABEL);

        Object obj = params.get(label.trim());

        if (obj instanceof ImageEntity || obj instanceof List) {
            return obj;
        }
        if (ObjectUtils.isNotEmpty(obj)) {
            currentText = obj.toString();
        } else {
            currentText = ElCore.EMPTY;
        }
        
        return currentText;
    }

}
