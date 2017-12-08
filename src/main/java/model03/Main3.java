package model03;

import model01.Main1;
import model01.ReadWord;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

public class Main3 {
    public static String REGEX = "\\d?[.][\\s\\w\\W]+:";

    public static void main(String[] args) throws IOException {
        Map<String, XWPFDocument> list = ReadWord.readDocxs(Paths.get(Main1.OUT));
        list.forEach((s, xwpfDocument) -> {
            System.out.println(s);
            List<XWPFParagraph> parlist = xwpfDocument.getParagraphs();
            parlist.forEach(paragraph -> {
                String str = paragraph.getText();
                List<XWPFRun> runslist = paragraph.getRuns();
                if (str.matches(REGEX)) {
                    runslist.forEach(xwpfRun -> {
                        xwpfRun.setBold(true);
                    });
                }
                runslist.forEach(xwpfRun -> {
                    xwpfRun.setFontFamily("arial");
                    xwpfRun.setFontSize(11);
                });
            });

            FileOutputStream fos = null;
            try {
                fos = new FileOutputStream(Paths.get(Main1.OUT + "\\" + s).toFile());
                xwpfDocument.write(fos);
                fos.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
    }
}
