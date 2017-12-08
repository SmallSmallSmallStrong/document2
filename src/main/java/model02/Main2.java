package model02;

import model01.Main1;
import model01.ReadExcel;
import model01.ReadWord;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;

/**
 * 第二步
 */
public class Main2 {
    public static void main(String[] args) {
        //读取附件4 - excel文档
        List<String> list = ReadExcel.readExcel(Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\4-替换内容.xlsx"));
        final int[] i = {1};
        list.forEach(s -> {
            String[] ss = s.split(ReadExcel.SPLITTER);
            if (ss.length == 2) {
//                System.out.println(ss[1]);
                //进行替换操作
                if ("A".equals(ss[1].toUpperCase())) {
                    try {
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"),"<C>steel frame<C>",ss[0] +"(A)");
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"),"<J>Steel building<J>","Steel building(A)");
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if ("B".equals(ss[1].toUpperCase())) {
                    try {
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"),"<J>Steel building<J>",ss[0] +"(B)");
                        ReadWord.replaceAll(Paths.get(Main1.OUT + "\\" + i[0] + ".docx"),"<C>steel frame<C>","steel frame(B)");
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                i[0]++;
            }
        });
    }
}
