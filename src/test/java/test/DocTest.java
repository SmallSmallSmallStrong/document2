package test;

import model01.ReadWord;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public class DocTest {

    @Test
    public void inster() {

        Path p = Paths.get("C:\\Users\\sz\\Desktop\\文档内容批量处理程序\\1-文档\\1.doc");
        try {

            List<String> alllines = Files.readAllLines(p, Charset.forName("GB2312"));
//            InputStream is = new FileInputStream(p.toFile());
//            HWPFDocument hwpfDocument = new HWPFDocument(is);
//            StringBuilder text2003 = hwpfDocument.getText();
//            WordExtractor ex = new WordExtractor(is);
//            String text2003 = ex.getText();
//            System.out.println(text2003.toString());

            XWPFDocument doc_2 = new XWPFDocument();
            alllines.forEach(s -> {
                XWPFParagraph newpar = doc_2.createParagraph();
                XWPFRun run = newpar.createRun();
                run.setText(s, 0);
            });
            FileOutputStream fout = new FileOutputStream(Paths.get("C:\\Users\\sz\\Desktop\\文档内容批量处理程序\\out\\2.docx").toFile());
            doc_2.write(fout);
            fout.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
