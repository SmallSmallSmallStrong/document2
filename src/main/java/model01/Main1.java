package model01;

import model03.Main3;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main1 {

    public static final String PAK_PATH = "E:\\文档内容批量处理程序\\";
    //第一步替换查找的内容
    public static final String REGEX = "1. A";
    //附件一的文档文件夹系统路径
    public static final String WORD1 = PAK_PATH + "1-文档";
    //附件三的文档文件夹的系统路径
    public static final String WORD3 = PAK_PATH + "3-需要提取的第一条信息";
    //输出路径
    public static final String OUT = PAK_PATH + "out";
    //关系文件
    public static final String CONTENT = PAK_PATH + "2-文档中第一条信息选择.xlsx";
    //替换内容
    public static final String REPLACEEXCEL = PAK_PATH + "4-替换内容.xlsx";
    //替换文本A
    public static final String REPLACE_TEXT_A = " CCC ";
    public static final String REPLACE_TEXT_B = " JJJ ";


    public static void main(String[] args) {
        if (one()) {
            System.out.println("第一步执行成功");
            if (two()) {
                System.out.println("第二步执行成功");
                if (three()) {
                    System.out.println("第三步执行成功");
                } else System.err.println("第三步执行失败");
            } else System.err.println("第二步执行失败");
        } else System.err.println("第一步执行失败");
    }

    public static boolean one() {
        Map<String, List<String>> map1 = new HashMap<>();
        //读取附件 1的文档
        //文件筛选过滤器
        //创建过滤器 保证 改为文件不是
        Path path = Paths.get(WORD1);
        try {
            map1 = ReadWord.readDoc(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //读取关系 Excel
        final List<String> guanxilist = new ArrayList<>();
        Path excelpath = Paths.get(CONTENT);
        if (Files.exists(excelpath)) {
            guanxilist.clear();
            guanxilist.addAll(ReadExcel.readExcel(excelpath));
        } else {
            System.err.println("文件不存在");
            return false;
        }
        //读取附件3需要替换的内容
        //替换 生成新的文档
        Map<String, List<String>> finalMap1 = map1;
        for (int k = 0; k < guanxilist.size(); k++) {
            //获取附件1的对应的内容
            List<String> doc_1 = finalMap1.get((k + 1) + ".doc");
            try {
                XWPFDocument docx_3 = ReadWord.readDocx(Paths.get(WORD3 + "\\" + Util.docxname(guanxilist.get(k)) + ".docx"));
                XWPFDocument doc_2 = docx_3.getXWPFDocument();
                //判断第一条是不是 要替换的内容
                if (doc_1.get(0).equals(Main1.REGEX)) {
                    XWPFRun run = doc_2.getParagraphs().get(0).getRuns().get(0);
                    run.setText("1. " + run.getText(0), 0);
                    //先添加1的内容
                    for (int p = 0; p < doc_1.size(); p++) {
                        if (p != 0) {
                            XWPFParagraph newpar = doc_2.createParagraph();
                            XWPFRun run1 = newpar.createRun();
                            run1.setText(doc_1.get(p));
                        }
                    }
                }
                Path p = Paths.get(OUT + "\\" + (k + 1) + ".docx");
                Path file = Files.createFile(p);
                FileOutputStream os = new FileOutputStream(file.toFile());
                doc_2.write(os);
                os.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return true;
    }

    public static boolean two() {
        //读取附件4 - excel文档
        List<String> list = ReadExcel.readExcel(Paths.get(REPLACEEXCEL));
        if (list == null || list.size() == 0) return false;
        final int[] i = {1};
        list.forEach(s -> {
            String[] ss = s.split(ReadExcel.SPLITTER);
            if (ss.length == 2) {
                //进行替换操作
                if ("A".equals(ss[1].toUpperCase())) {
                    try {
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"), REPLACE_TEXT_A, ss[0]);
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"), REPLACE_TEXT_B, "Steel building");
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if ("B".equals(ss[1].toUpperCase())) {
                    try {
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"), REPLACE_TEXT_B, ss[0]);
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"), REPLACE_TEXT_A, "steel frame");
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                i[0]++;
            } else {
                return;
            }
        });

        if (i[0] < list.size()) System.out.println("替换文件数为:" + i[0]);

        return true;
    }

    public static boolean three() {
        Map<String, XWPFDocument> list = null;
        try {
            list = ReadWord.readDocxs(Paths.get(Main1.OUT));
        } catch (IOException e) {
            e.printStackTrace();
        }
        list.forEach((s, xwpfDocument) -> {
//            System.out.println(s);
            List<XWPFParagraph> parlist = xwpfDocument.getParagraphs();
            parlist.forEach(paragraph -> {
                String str = paragraph.getText();
                List<XWPFRun> runslist = paragraph.getRuns();
                if (str.matches(Main3.REGEX)) {
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
        return true;
    }
}
