package model01;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main1 {
    //第一步替换查找的内容
    public static final String REGEX = "1. A";
    //附件一的文档文件夹系统路径
    public static final String WORD1 = "C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\1-文档";
    //附件三的文档文件夹的系统路径
    public static final String WORD3 = "C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\3-需要提取的第一条信息";
    //输出路径
    public static final String OUT = "C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\1-out";
    //关系文件
    public static final String CONTENT = "C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\2-文档中第一条信息选择.xlsx";

    public static void main(String[] args) {
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
        }
        //读取附件3需要替换的内容
        Path path3 = Paths.get(WORD3);
        Map<String, XWPFDocument> docxmap = new HashMap<>();
        try {
            docxmap = ReadWord.readDocxs(path3);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //替换 生成新的文档
        final int[] i = {1};
        Map<String, List<String>> finalMap1 = map1;
        Map<String, XWPFDocument> finalDocxmap = docxmap;
        guanxilist.forEach(s -> {
            //获取附件1的对应的内容
            List<String> doc_1 = finalMap1.get(i[0] + ".doc");
            XWPFDocument docx_3 = finalDocxmap.get(Util.docxname(s) + ".docx");
            try {
                //判断第一条是不是 要替换的内容
                if (doc_1.get(0).equals(Main1.REGEX)) {
                    //如果是则替换
                    final boolean[] first = {true};
                    doc_1.forEach(s1 -> {
                        if (first[0]) {
                            first[0] = false;
                        } else {
                            XWPFParagraph para = docx_3.createParagraph();
                            XWPFRun run = para.createRun();
                            run.setText(s1);
                        }
                    });
                }
                Path p = Paths.get(OUT + "\\" + i[0] + ".docx");
                Path file = Files.createFile(p);
                FileOutputStream os = new FileOutputStream(file.toFile());
                docx_3.write(os);
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.out.println(i[0]++ + "" + s);
        });


    }


}
