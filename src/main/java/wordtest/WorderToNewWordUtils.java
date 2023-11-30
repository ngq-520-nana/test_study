package wordtest;


import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WorderToNewWordUtils {

    private static final Log logger = LogFactory.getLog(WorderToNewWordUtils.class);
    /**
     * 根据模板生成word文档
     *
     * @param textMap 需要替换的文本内容
     * @param mapList 需要动态生成的内容
     *
     */
//    public static CustomXWPFDocument changWord(InputStream inputStream, Map<String, Object> textMap, Map<String,Object> mapList, List<Map<String,Object>> richText) {
//        CustomXWPFDocument document = null;
//        try {
//            document = new CustomXWPFDocument(inputStream);
//
//            //处理富文本内容,没有可以忽略
//            if(CollectionUtils.isNotEmpty(richText)){
//                WorderToNewWordUtils.wordInsertRitchText(document, richText);
//            }
//
//            //解析替换文本段落对象
//            WorderToNewWordUtils.processParagraphs1(document, textMap);
//
//            //解析替换表格对象
//            WorderToNewWordUtils.changeTable(document, textMap, mapList);
//        } catch (Exception e) {
//           // e.printStackTrace();
//            logger.error("解析失败"+e.getMessage(),e);
//        }
//        return document;
//    }
    public static CustomXWPFDocument changWord(String url, Map<String, Object> textMap, Map<String,Object> mapList, List<Map<String,Object>> richText) {
        CustomXWPFDocument document = null;
        try {
            document = new CustomXWPFDocument(POIXMLDocument.openPackage(url));

            //处理富文本内容,没有可以忽略
            if(CollectionUtils.isNotEmpty(richText)){
                WorderToNewWordUtils.wordInsertRitchText(document, richText);
            }

            //解析替换文本段落对象
            WorderToNewWordUtils.processParagraphs1(document, textMap);

            //解析替换表格对象
            WorderToNewWordUtils.changeTable(document, textMap, mapList);
        } catch (Exception e) {
           // e.printStackTrace();
            logger.error("解析失败"+e.getMessage(),e);
        }
        return document;
    }
    protected static InputStream getInputStreamByUrl(String url) {
        try {
            URL conUrl = new URL(url);
            HttpURLConnection conn = (HttpURLConnection)conUrl.openConnection();
            conn.setReadTimeout(5000);
            conn.setConnectTimeout(5000);
            if (conn.getResponseCode() != 200) {

            } else {
                InputStream inputStream = conn.getInputStream();
                return inputStream;
            }
        } catch (Exception e) {
            logger.error("获取URL失败"+e.getMessage(),e);
        }
        return null;
    }

    /**
     * 替换段落文本
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     */
    public static void changeText(CustomXWPFDocument document, Map<String, Object> textMap){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if(checkText(text)){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    // 段落换行
                    String value = ob.toString();
                    if (value != null){
                        run.setText((String)ob,0);
                    }
                }
            }
        }
    }

    public static void processParagraphs1(CustomXWPFDocument document, Map<String, Object> textMap){
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        if(paragraphs == null || paragraphs.size() <= 0){
            return;
        }
        if(paragraphs != null && paragraphs.size() > 0){
            for(XWPFParagraph paragraph:paragraphs){
                //poi转换过来的行间距过大，需要手动调整
                if(paragraph.getSpacingBefore() >= 1000 || paragraph.getSpacingAfter() > 1000) {
                    paragraph.setSpacingBefore(0);
                    paragraph.setSpacingAfter(0);
                }
                //设置word中左右间距
                paragraph.setIndentationLeft(0);
                paragraph.setIndentationRight(0);
                List<XWPFRun> runs = paragraph.getRuns();
                //加了图片，修改了paragraph的runs的size，所以循环不能使用runs
                List<XWPFRun> allRuns = new ArrayList<>(runs);
                for (XWPFRun run : allRuns) {
                    String text = run.getText(0);
                    if(text == null){
                        continue;
                    }
                    boolean isSetText = false;
                    for (Entry<String, Object> entry : textMap.entrySet()) {
                        String key = entry.getKey();
                        if(text.contains(key)){
                            isSetText = true;
                            Object value = entry.getValue();
                            if(value == null){
                                text = text.replace(key,"");
                            }
                            if (value instanceof Map) {//图片替换
                                text = text.replace(key, "");
                                Map pic = (Map)value;
                                int width = Integer.parseInt(pic.get("width").toString());
                                int height = Integer.parseInt(pic.get("height").toString());
                                int picType = getPictureType(pic.get("type").toString());
                                byte[] byteArray = (byte[]) pic.get("content");
                                ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                                try {
                                    String blipId = document.addPictureData(byteInputStream,picType);
                                    document.createPicture(blipId,document.getNextPicNameNumber(picType), width, height,paragraph);
                                } catch (Exception e) {
                                    logger.error("图片转换失败"+e.getMessage(),e);
                                }
                            } else {//文本替换
                                assert value != null;
                                text = text.replace(key,StringUtils.isNotEmpty(value.toString())?value.toString():"-");
                            }
                        }
                    }
                    if(isSetText){
                        run.setText(text,0);
                    }
                }
            }
        }
    }



    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param mapList 需要动态生成的内容
     */


    public static void changeTable(CustomXWPFDocument document, Map<String, Object> textMap, Map<String,Object> mapList) throws InvalidFormatException {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();

        //循环所有需要进行替换的文本，进行替换
        for (XWPFTable table : tables) {
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                //遍历表格,并替换模板
                eachTable(document, rows, textMap);
            }
        }

        List<String[]> list01 = (List<String[]>) mapList.get("list01");
        List<String[]> list02 = (List<String[]>) mapList.get("list02");
        List<String[]> list03 = (List<String[]>) mapList.get("list03");
        List<String[]> list04 = (List<String[]>) mapList.get("list04");
        List<String[]> list05 = (List<String[]>) mapList.get("list05");
        List<String[]> list06 = (List<String[]>) mapList.get("list06");
        //操作word中的表格
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            //表格从 0 开始 i == 0 对应第一个表格
            XWPFTable table = tables.get(i);
            if (null != list01 && 0 < list01.size() && i == 0){
                insertTable(table, list01,3,1);
            }
            if (null != list02 && 0 < list02.size() && i == 1){
                insertTable(table, list02,3,1);
            }
            if (null != list03 && 0 < list03.size() && i == 2){
                 insertTable(table, list03,3,1);
            }
            if (null != list04 && 0 < list04.size() && i == 3){
                insertTable(table, list04,10,2);
            }
            if (null != list05 && 0 < list05.size() && i == 4){
                insertTable(table, list05,8,2);
            }
            if (null != list06 && 0 < list06.size() && i == 5){
                insertTable(table, list06,7,2);
            }
        }
    }



    /**
     * 遍历表格
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(CustomXWPFDocument document, List<XWPFTableRow> rows , Map<String, Object> textMap) throws InvalidFormatException {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            Object ob = changeValue(run.toString(), textMap);
                            if (ob instanceof String){
                                run.setText((String)ob,0);
                            }else if (ob instanceof Map){
                                run.setText("",0);
                                Map pic = (Map)ob;
                                int width = Integer.parseInt(pic.get("width").toString());
                                int height = Integer.parseInt(pic.get("height").toString());
                                int picType = getPictureType(pic.get("type").toString());
                                byte[] byteArray = (byte[]) pic.get("content");
                                ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                                try {
                                    String blipId = document.addPictureData(byteInputStream,picType);
                                    document.createPicture(blipId,document.getNextPicNameNumber(picType), width, height,paragraph);
                                } catch (Exception e) {
                                    logger.error("图片转换失败"+e.getMessage(),e);
                                    throw e;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public static void insertTable(XWPFTable table, List<String[]> daList, Integer num, Integer heardrows){
        //创建行和创建需要的列
        if(daList.size()>=2){
            for(int i = 1; i <= daList.size(); i++){
                XWPFTableRow row = table.insertNewTableRow(heardrows+i-1);//添加一个新行
                for (int k = 0; k < num; k++) {
                    row.createCell();//添加第K+1个列
                }
            }
        }else if(daList.size()==1){
            XWPFTableRow row = table.insertNewTableRow(heardrows);//添加一个新行
            for (int k = 0; k < num; k++) {
                row.createCell();//添加第K+1个列
            }
        }
        //创建行,根据需要插入的数据添加新行，不处理表头
        for(int i = 0; i < daList.size(); i++){
            List<XWPFTableCell> cells = table.getRow(i+heardrows).getTableCells();
            for(int j = 0; j < cells.size(); j++){
                XWPFTableCell cell02 = cells.get(j);
                cell02.setText(daList.get(i)[j]);
            }
        }
    }

    public static void insertRow(XWPFTable table, int copyrowIndex, int newrowIndex) {
        // 在表格中指定的位置新增一行
        XWPFTableRow targetRow = table.insertNewTableRow(newrowIndex);
        // 获取需要复制行对象
        XWPFTableRow copyRow = table.getRow(copyrowIndex);
        //复制行对象
        targetRow.getCtRow().setTrPr(copyRow.getCtRow().getTrPr());
        //或许需要复制的行的列
        List<XWPFTableCell> copyCells = copyRow.getTableCells();
        //复制列对象
        XWPFTableCell targetCell;
        for (XWPFTableCell copyCell : copyCells) {
            targetCell = targetRow.addNewTableCell();
            targetCell.getCTTc().setTcPr(copyCell.getCTTc().getTcPr());
            if (copyCell.getParagraphs() != null && copyCell.getParagraphs().size() > 0) {
                targetCell.getParagraphs().get(0).getCTP().setPPr(copyCell.getParagraphs().get(0).getCTP().getPPr());
                if (copyCell.getParagraphs().get(0).getRuns() != null
                        && copyCell.getParagraphs().get(0).getRuns().size() > 0) {
                    XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
                    cellR.setBold(copyCell.getParagraphs().get(0).getRuns().get(0).isBold());
                }
            }
        }
    }

    /**
     * 判断文本中时候包含$
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(text.indexOf("$")!= -1){
            check = true;
        }
        return check;
    }

    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static Object changeValue(String value, Map<String, Object> textMap){
        Set<Entry<String, Object>> textSets = textMap.entrySet();
        Object valu = "";
        for (Entry<String, Object> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if(value.indexOf(key)!= -1){
                valu = textSet.getValue();
            }
        }
        return valu;
    }

    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @return int
     */
    private static int getPictureType(String picType){
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if(picType != null){
            if(picType.equalsIgnoreCase("png")){
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }


    public static Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)}", Pattern.CASE_INSENSITIVE);
        return pattern.matcher(str);
    }

    /**
     * 往doc的标记位置插入富文本内容 注意：目前支持富文本里面带url的图片，不支持base64编码的图片
     *
     * @param doc          需要插入内容的Word
     * @param ritchtextMap 标记位置对应的富文本内容
     * @param
     */
    public static void wordInsertRitchText(CustomXWPFDocument doc, List<Map<String, Object>> ritchtextMap) {
        try {
            // 如果需要替换多份富文本，通过Map来操作，key:要替换的标记，value：要替换的富文本内容
            for (Map<String, Object> mapList : ritchtextMap) {
                for (Entry<String, Object> entry : mapList.entrySet()) {
                    for (XWPFParagraph paragraph : doc.getParagraphs()) {
                        if (entry.getKey().equals(paragraph.getText().trim())) {
                            // 在标记处插入指定富文本内容
                            HtmlUtils.resolveHtml(entry.getValue().toString(), doc, paragraph);
                            doc.removeBodyElement(doc.getPosOfParagraph(paragraph));
                            break;
                        }
                    }

                }
            }
        } catch (Exception e) {
            logger.error("富文本内容替换失败"+e.getMessage(),e);
        }
    }


}
