import model01.Main1;
import model01.ReadWord;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;

public class Test1 {
    public static final Path A27 = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\3-需要提取的第一条信息\\A (27).docx");
    public static final Path NEWA27 = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\test\\A (27).docx");

    @Test
    public void charsetDemo() {
        //获取制定编码实例
        Charset urt8 = Charset.forName("utf-8");
        System.out.println(urt8);
        //all 获取系统支持的所有编码形式
        SortedMap<String, Charset> all = Charset.availableCharsets();
        all.forEach((s, charset) -> {
            System.out.println(s + "->" + charset);
        });
        //获取默认编码
        Charset def = Charset.defaultCharset();
        System.out.println(def);
    }

    @Test
    public void readA27() {
        try {
            FileInputStream fis = new FileInputStream(A27.toFile());
            XWPFDocument xdoc = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//            extractor.appendBodyElementText(new StringBuffer("12312312312312"), new XWPFParagraph());
            String doc1 = extractor.getText();
            System.out.println(doc1);
            fis.close();
            OutputStream os = new FileOutputStream(NEWA27.toFile());
            xdoc.write(os);

            os.close();
//            List<String> list = Files.readAllLines(a27, Charset.forName("GB2312"));
//            Path file = Files.createFile(newa27);
//            Files.write(file, list, StandardOpenOption.APPEND);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readDocxs() {
        Path a27 = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\3-需要提取的第一条信息");
        try {
            Map<String, XWPFDocument> map = ReadWord.readDocxs(a27);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readDocxs2() throws IOException {
        Path p = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\3-需要提取的第一条信息\\A (2).docx");
        Path p2 = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\1-文档\\1.doc");
        FileInputStream fis = new FileInputStream(p.toFile());
        XWPFDocument xdoc = new XWPFDocument(fis);
        XWPFParagraph para = xdoc.createParagraph();
        XWPFRun run = para.createRun();

//        xdoc.insertNewParagraph(paragraph);
        System.out.println(xdoc.getParagraphs().get(5).getText());
        ;
        XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//                String doc1 = extractor.getText();
//                System.out.println(doc1);
        fis.close();
    }

    @Test
    public void readDoc() throws IOException {
        Path a27 = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\1-文档\\2.doc");
        InputStream fis = new FileInputStream(new File("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\1-文档\\2.doc"));
        WordExtractor ex = new WordExtractor(fis);//is是WORD文件的InputStream
        String text2003 = ex.getText();
        System.out.println(text2003);
//        POIFSFileSystem fs = new POIFSFileSystem(fis);
//        System.out.println(fs.toString());
//        HWPFDocument doc = new HWPFDocument(fs);
//        String doc1 = doc.getDocumentText();
//        System.out.println(doc1);
//        StringBuilder doc2 = doc.getText();
//        System.out.println(doc2);
//        Range rang = doc.getRange();
//        String doc3 = rang.text();
//        System.out.println(doc3);
//        fis.close();
    }

    @Test
    public void Test00() {
        String s = "A123";
        String a = s.substring(1);
//        System.out.println(a);
        System.out.println(s.substring(1, s.length()));
    }

    @Test
    public void DocxRemoveParagraphs() {
        try {
            //原文件
            Path path = Paths.get(Main1.OUT + "\\100.docx");
            XWPFDocument docx = ReadWord.readDocx(path);
            List<XWPFParagraph> list = docx.getParagraphs();
            for (XWPFParagraph paragraph : list) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = "k";
                    String replace = "abcd";
                    String str = run.getText(0);
                    if (text != null && text.contains(text)){
                     str = str.replaceAll(text, replace);
                     run.setText(str,0);
                    }
                }
            }
            FileOutputStream fos = new FileOutputStream(path.toFile());
            docx.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Test
    public void replaceAlltest() {
        try {
            ReadWord.replaceAll(Paths.get(Main1.OUT + "\\5.docx"),"<J>Steel building<J>","你是山东的");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
