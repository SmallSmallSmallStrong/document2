package model01;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadWord {
    /**
     * 传入路径 进行读取
     *
     * @param path
     * @return
     */
    public static Map<String, XWPFDocument> readDocxs(Path path) throws IOException {
        Map<String, XWPFDocument> res = new HashMap<>();
        DirectoryStream.Filter<Path> filter = file -> {
            boolean isok = true;
            //判断前缀（是否是隐藏文件）
            if (file.getFileName().toString().startsWith("~$")) isok = false;
            //判断后缀
            String glob2 = "*.docx";
            FileSystem fs = file.getFileSystem();
            final PathMatcher matcher2 = fs.getPathMatcher("glob:" + glob2);
            if (matcher2.matches(file.getFileName()) == false)
                isok = false;
            return isok;
        };
        DirectoryStream<Path> paths = Files.newDirectoryStream(path, filter);
        paths.forEach(p -> {
            String filename = p.getFileName().toString();
            try {
                res.put(filename, readDocx(p));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        return res;
    }

    public static XWPFDocument readDocx(Path path) throws IOException {
        FileInputStream fis = new FileInputStream(path.toFile());
        XWPFDocument xdoc = new XWPFDocument(fis);
        XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//                String doc1 = extractor.getText();
//                System.out.println(doc1);
        fis.close();
        return xdoc;
    }

    public static Map<String, List<String>> readDoc(Path path) throws IOException {
        Map<String, List<String>> map = new HashMap<>();
        DirectoryStream.Filter<Path> filter = file -> {
            boolean isok = true;
            //判断前缀（是否是隐藏文件）
            if (file.getFileName().toString().startsWith("~$")) isok = false;
            //判断后缀
            String glob2 = "*.doc";
            FileSystem fs = file.getFileSystem();
            final PathMatcher matcher2 = fs.getPathMatcher("glob:" + glob2);
            if (matcher2.matches(file.getFileName()) == false)
                isok = false;
            return isok;
        };
        try {
            DirectoryStream<Path> paths = Files.newDirectoryStream(path, filter);
            paths.forEach(p -> {
                List<String> alllines = null;
                try {
//                    忽略隐藏文件
                    if (!p.getFileName().toString().startsWith("~$"))
                        alllines = Files.readAllLines(p, Charset.forName("GB2312"));
                    map.put(p.getFileName().toString(), alllines);
                } catch (IOException e) {
                    System.err.println(p.getFileName());
                    e.printStackTrace();
                }

            });
        } catch (IOException e) {
            System.out.println("获取paths，文件列表失败");
            e.printStackTrace();
        }
        return map;
    }


    public static Map<String, List<String>> readDoc2(Path path) throws IOException {
        Map<String, List<String>> map = new HashMap<>();
        DirectoryStream.Filter<Path> filter = file -> {
            boolean isok = true;
            //判断前缀（是否是隐藏文件）
            if (file.getFileName().toString().startsWith("~$")) isok = false;
            //判断后缀
            String glob2 = "*.doc";
            FileSystem fs = file.getFileSystem();
            final PathMatcher matcher2 = fs.getPathMatcher("glob:" + glob2);
            if (matcher2.matches(file.getFileName()) == false)
                isok = false;
            return isok;
        };
        try {
            DirectoryStream<Path> paths = Files.newDirectoryStream(path, filter);
            paths.forEach(p -> {
                List<String> alllines = null;
                try {
                    File file = new File("C:\\Users\\tuzongxun123\\Desktop\\aa.doc");
                    String str = "";
                    try {
                        FileInputStream fis = new FileInputStream(file);
                        HWPFDocument doc = new HWPFDocument(fis);
                        String doc1 = doc.getDocumentText();
                        System.out.println(doc1);
                        StringBuilder doc2 = doc.getText();
                        System.out.println(doc2);
                        Range rang = doc.getRange();
                        String doc3 = rang.text();
                        System.out.println(doc3);
                        fis.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
//                    忽略隐藏文件
                    if (!p.getFileName().toString().startsWith("~$"))
                        alllines = Files.readAllLines(p, Charset.forName("GB2312"));
                    map.put(p.getFileName().toString(), alllines);
                } catch (IOException e) {
                    System.err.println(p.getFileName());
                    e.printStackTrace();
                }

            });
        } catch (IOException e) {
            System.out.println("获取paths，文件列表失败");
            e.printStackTrace();
        }
        return map;
    }

    /**
     * 全部替换 .docx 的word文档
     *
     * @param path    docx 文档 路径
     * @param text    需要替换的内容
     * @param repalce 替换为repalce
     * @return
     */
    public static boolean replaceAll(Path path, String text, String repalce) throws IOException {
        XWPFDocument docx = readDocx(path);
        List<XWPFParagraph> ps = docx.getParagraphs();
        ps.forEach(paragraph -> {
            List<XWPFRun> runs = paragraph.getRuns();
            runs.forEach(xwpfRun -> {
                String str = xwpfRun.getText(0);
                if (str != null && str.contains(text)) {
                    str = str.replaceAll(text, repalce);
                    xwpfRun.setText(str, 0);
                }
            });
        });
        for (XWPFTable tbl : docx.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String str = r.getText(0);
                            if (str != null && text.contains(text)) {
                                str = str.replaceAll(text, repalce);
                                r.setText(str, 0);
                            }
                        }
                    }
                }
            }
        }

        FileOutputStream fos = new FileOutputStream(path.toFile());
        docx.write(fos);
        fos.close();
        return true;
    }


    public static boolean titkeBold() {
        return true;
    }
}
