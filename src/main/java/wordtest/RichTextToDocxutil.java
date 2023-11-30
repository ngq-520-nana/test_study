package wordtest;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author niuguoqiang
 * @date 2021年08月17日 17:41
 */
public class RichTextToDocxutil {
    /**
     * 导出富本框到docx
     */
    private static final Log logger = LogFactory.getLog(RichTextToDocxutil.class);
    public static void outRichTextToDocx(String contents ,String outFilePath) {
        String content = txt2String(contents);
        InputStream inputStream=null;
        OutputStream out = null;
        try {
            // 输入富文本内容，返回字节数组
            byte[] result = HtmlToWord.resolveHtml(content);
            //输出文件
            out = new FileOutputStream(outFilePath);
            out.write(result);
        } catch (IOException e) {
//            e.printStackTrace();
            logger.error("富文本内容输出失败",e);
        } finally {
            try {
                if(out != null) {
                    out.close();
                }
            } catch (IOException e) {
//                e.printStackTrace();
                logger.error("文件流关闭失败",e);
            }

        }
    }

    /**
     * 读取html文件的内容
     *
     * @param content 读取富文本
     * @return 返回文件内容
     */
    public static String txt2String(String content) {
        StringBuilder result = new StringBuilder();
        try {
            // 构造一个BufferedReader类来读取富文本
            result.append(System.lineSeparator()+content);
        } catch (Exception e) {
//            e.printStackTrace();
            logger.error("读取富文本内容失败",e);
        }
        return result.toString();
    }

}
