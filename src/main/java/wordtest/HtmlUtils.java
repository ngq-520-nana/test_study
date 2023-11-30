package wordtest;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;


/**
 * @author niuguoqiang
 * @date 2021年08月17日 17:50
 */
public class HtmlUtils {
    /**
     * 给document添加指定元素
     *
     *
     */
    private static final Log logger = LogFactory.getLog(HtmlUtils.class);

    public static void addElement(Document document) {
        if (ObjectUtils.isEmpty(document)) {
            throw new NullPointerException("不允许为空的对象添加元素");
        }
        Elements elements = document.getAllElements();
        for (Element e : elements) {
            String attrName = ElementEnum.getValueByCode(e.tag().getName());
            if (!StringUtils.isEmpty(attrName)) {
                e.attr(CommonConStant.COMMONATTR, attrName);
            }
        }
    }


    /**
     * 将富文本内容写入到Word
     * 因富文本样式种类繁多，不能一一枚举，目前实现了H1、H2、H3、段落、图片、表格枚举
     *
     * @param ritchText 富文本内容
     * @param doc       需要写入富文本内容的Word 写入图片和表格需要用到
     *
     *
     */
    public static void resolveHtml(String ritchText, XWPFDocument doc, XWPFParagraph paragraph) {
        Document document = Jsoup.parseBodyFragment(ritchText, "UTF-8");
        try {
            // 添加固定元素
            HtmlUtils.addElement(document);
            Elements elements = document.select("[" + CommonConStant.COMMONATTR + "]");
            for (Element em : elements) {
                XmlCursor xmlCursor = paragraph.getCTP().newCursor();
                switch (em.attr(CommonConStant.COMMONATTR)) {
                    case "title":
                        break;
                    case "subtitle":
                        break;
                    case "imgurl":
                        String src = em.attr("src");
//                        URL url = new URL(src);
//                        URLConnection uc = url.openConnection();
//                        InputStream inputStream = uc.getInputStream();
//                        InputStream inputStream = new FileInputStream(src);
//                        XWPFParagraph imgurlparagraph = doc.insertNewParagraph(xmlCursor);
//                        ParagraphStyleUtil.setImageCenter(imgurlparagraph);
//                        imgurlparagraph.createRun().addPicture(inputStream, XWPFDocument.PICTURE_TYPE_PNG, "图片.jpeg", Units.toEMU(150), Units.toEMU(150));
//                        closeStream(inputStream);
//                        File file = new File("picture.jpg");
//                        boolean exists = file.exists();
//                        if (exists) {
//                            Boolean boo = file.delete();
//                            if(!boo){
//                             logger.info("文件删除失败");
//                            }
//                        }
                        String code = src.replace("data:image/png;base64,", "");
                        BASE64Decoder decoder = new BASE64Decoder();
                        byte[] b = decoder.decodeBuffer(code);
                        for (int i = 0; i < b.length; ++i) {
                            if (b[i] < 0) {// 调整异常数据
                                b[i] += 256;
                            }
                        }
                        String SuiJi = (int) (1000 + Math.random() * (9999 - 1000 + 1)) + "";
                        String imgPath = "D:\\发票\\"+SuiJi+".jpg";
                        File tempFile = new File(imgPath);
                        if (!tempFile.getParentFile().exists()) {
                            tempFile.getParentFile().mkdirs();
                        }
                        OutputStream out = new FileOutputStream(imgPath);
                        out.write(b);
                        out.flush();
                        out.close();
                        InputStream inputStream = new FileInputStream(imgPath);
                        XWPFParagraph imgurlparagraph = doc.insertNewParagraph(xmlCursor);
                        //居中
                        ParagraphStyleUtil.setImageCenter(imgurlparagraph);
                        imgurlparagraph.createRun().addPicture(inputStream,XWPFDocument.PICTURE_TYPE_PNG,"图片.jpeg", Units.toEMU(200),Units.toEMU(200));
                        closeStream(inputStream);
                        break;
                    case "imgbase64":
                        break;
                    case "table":
                        XWPFTable xwpfTable = doc.insertNewTbl(xmlCursor);
                        addTable(xwpfTable, em);
                        // 设置表格居中
                        ParagraphStyleUtil.setTableLocation(xwpfTable, "center");
                        // 设置内容居中
                        ParagraphStyleUtil.setCellLocation(xwpfTable, "CENTER", "center");
                        break;
                    case "h1":
                        XWPFParagraph h1paragraph = doc.insertNewParagraph(xmlCursor);
                        XWPFRun xwpfRun_1 = h1paragraph.createRun();
                        xwpfRun_1.setText(em.text());
                        //居中
                        ParagraphStyleUtil.setImageCenter(h1paragraph);
                        // 设置字体
                        ParagraphStyleUtil.setTitle(xwpfRun_1, TitleFontEnum.H1.getTitle());
                        break;
                    case "h2":
                        XWPFParagraph h2paragraph = doc.insertNewParagraph(xmlCursor);
                        XWPFRun xwpfRun_2 = h2paragraph.createRun();
                        xwpfRun_2.setText(em.text());
                        //居中
                        ParagraphStyleUtil.setImageCenter(h2paragraph);
                        // 设置字体
                        ParagraphStyleUtil.setTitle(xwpfRun_2, TitleFontEnum.H2.getTitle());
                        break;
                    case "h3":
                        XWPFParagraph h3paragraph = doc.insertNewParagraph(xmlCursor);
                        XWPFRun xwpfRun_3 = h3paragraph.createRun();
                        xwpfRun_3.setText(em.text());
                        // 设置字体
                        ParagraphStyleUtil.setTitle(xwpfRun_3, TitleFontEnum.H3.getTitle());
                        break;
                    case "paragraph":
                        XWPFParagraph paragraphd = doc.insertNewParagraph(xmlCursor);
                        // 设置段落缩进 4个空格
                        paragraphd.createRun().setText("    " + em.text());
                        break;
                    case "br":
                        XWPFParagraph br = doc.insertNewParagraph(xmlCursor);
                        XWPFRun run = br.createRun();
                        run.addBreak(BreakType.TEXT_WRAPPING);
                        break;
                    default:
                        break;
                }
            }

        } catch (Exception e) {
            logger.error("转换元素失败",e);
        }
    }

    /**
     * 关闭输入流
     *
     *
     */
    public static void closeStream(Closeable... closeables) {
        for (Closeable c : closeables) {
            if (c != null) {
                try {
                    c.close();
                } catch (IOException e) {
                    logger.error("流关闭失败",e);
                }
            }
        }

    }

    /**
     * 将富文本的表格转换为Word里面的表格
     */
    private static void addTable(XWPFTable xwpfTable, Element table) {
        Elements trs = table.getElementsByTag("tr");
        // XWPFTableRow 第0行特殊处理
        int rownum = 0;
        for (Element tr : trs) {
            addTableTr(xwpfTable, tr, rownum);
            rownum++;
        }
    }


    /**
     * 将元素里面的tr 提取到 xwpfTabel
     */
    private static void addTableTr(XWPFTable xwpfTable, Element tr, int rownum) {
        Elements tds = tr.getElementsByTag("th").isEmpty() ? tr.getElementsByTag("td") : tr.getElementsByTag("th");
        XWPFTableRow row_1 = null;
        for (int i = 0, j = tds.size(); i < j; i++) {
            if (0 == rownum) {
                // XWPFTableRow 第0行特殊处理,
                XWPFTableRow row_0 = xwpfTable.getRow(0);
                if (i == 0) {
                    row_0.getCell(0).setText(tds.get(i).text());
                } else {
                    row_0.addNewTableCell().setText(tds.get(i).text());
                }
            } else {
                if (i == 0) {
                    // 换行需要创建一个新行
                    row_1 = xwpfTable.createRow();
                    row_1.getCell(i).setText(tds.get(i).text());
                } else {
                    row_1.getCell(i).setText(tds.get(i).text());
                }
            }
        }

    }

    public static String Image2base64(String imgurl) throws IOException {
        URL url = null;
        InputStream is = null;
        ByteArrayOutputStream outputStream = null;
        HttpURLConnection httpURLConnection = null;
        try {
            url = new URL(imgurl);
            httpURLConnection = (HttpURLConnection) url.openConnection();
            httpURLConnection.connect();
            httpURLConnection.getInputStream();
            outputStream = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int len = 0;
            while ((len = is.read(buffer)) != -1){
                outputStream.write(buffer,0,len);
            }
            return new BASE64Encoder().encode(outputStream.toByteArray());
        } catch (MalformedURLException e) {
            e.printStackTrace();
        } finally {
            if(is != null){
                is.close();
            }
            if(outputStream != null){
                outputStream.close();
            }
            if(httpURLConnection != null){
                httpURLConnection.disconnect();
            }
        }
        return imgurl;
    }

}
