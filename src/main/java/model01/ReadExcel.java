package model01;

import jodd.datetime.JDateTime;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ReadExcel {

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    //一行的分隔符
    public static final String SPLITTER = "#";

    /**
     * 判断文件是否是excel
     *
     * @throws Exception
     */
    public static void checkExcelVaild(File file) throws Exception {
        if (!file.exists()) {
            throw new Exception("文件不存在");
        }
        if (!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))) {
            throw new Exception("文件不是Excel");
        }
    }

    public static List<String> readExcel(Path path) {
        List<String> list = new ArrayList<>();
        try {
            // 同时支持Excel 2003、2007
            File excelFile = path.toFile();// 创建文件对象
            FileInputStream is = new FileInputStream(excelFile); // 文件流
            checkExcelVaild(excelFile);
//            Workbook workbook = getWorkbok(is, excelFile);
            Workbook workbook = WorkbookFactory.create(is); //这种方式 Excel2003/2007/2010都是可以处理的
            int sheetCount = workbook.getNumberOfSheets(); //Sheet的数量
            /**
             * 设置当前excel中sheet的下标：0开始
             */
            Sheet sheet = workbook.getSheetAt(0);   //遍历第一个Sheet
            // 为跳过第一行目录设置count
            for (Row row : sheet) {
                list.add(sheet(row));
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

//    public static List<String> readExcel(Path path, String rownum) throws Exception {
//        List<String> list = new ArrayList<>();
//        File excelFile = path.toFile();
//        FileInputStream is = new FileInputStream(excelFile);
//        checkExcelVaild(excelFile);
//        Workbook workbook = WorkbookFactory.create(is);
//        int sheetCount = workbook.getNumberOfSheets();
//        Sheet sheet = workbook.getSheetAt(0);
//        int count = 0;
////        if (sheetCount == 1)
//        for (Row row : sheet) {
//
//        }
//        return list;
//    }

    private static String sheet(Row row) {
        // 如果当前行没有数据，跳出循环
        if (row.getCell(0).toString().equals("")) {
            return "";
        }
        String rowValue = "";
        for (Cell cell : row) {
            if (cell.toString() == null) {
                continue;
            }
            CellType e = cell.getCellTypeEnum();
            String cellValue = "";
            switch (e) {
                case STRING:     // 文本
                    cellValue = cell.getRichStringCellValue().getString() + SPLITTER;
                    break;
                case NUMERIC:    // 数字、日期
                    if (DateUtil.isCellDateFormatted(cell)) {
                        JDateTime date = new JDateTime(cell.getDateCellValue());
                        cellValue = date.getFormat() + SPLITTER;
                    } else {
                        cell.setCellType(CellType.STRING);
                        cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + SPLITTER;
                    }
                    break;
                case BOOLEAN:    // 布尔型
                    cellValue = String.valueOf(cell.getBooleanCellValue()) + SPLITTER;
                    break;
                case BLANK: // 空白
                    cellValue = cell.getStringCellValue();
                    break;
                case ERROR:
                    cellValue = "错误" + SPLITTER;
                    break;
                case FORMULA:    // 公式
                    // 得到对应单元格的公式
                    //cellValue = cell.getCellFormula() + "#";
                    // 得到对应单元格的字符串
                    cell.setCellType(CellType.STRING);
                    cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + SPLITTER;
                    break;
                default:
                    cellValue = SPLITTER;
            }
            //System.out.print(cellValue);
            rowValue += cellValue;
        }
        return rowValue;
    }

    /**
     * 读取Excel测试，兼容 Excel 2003/2007/2010
     *
     * @throws Exception
     */

    public static void main(String[] args) throws Exception {
        BufferedWriter bw = new BufferedWriter(new FileWriter(new File("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\a.txt")));
        try {
            // 同时支持Excel 2003、2007
            Path excelpath = Paths.get("C:\\Users\\sz\\Desktop\\5-文档内容批量处理程序(1)\\2-文档中第一条信息选择.xlsx");
            File excelFile = excelpath.toFile();// 创建文件对象
            FileInputStream is = new FileInputStream(excelFile); // 文件流
            checkExcelVaild(excelFile);
            Workbook workbook = getWorkbok(is, excelFile);
            //Workbook workbook = WorkbookFactory.create(is); //这种方式 Excel2003/2007/2010都是可以处理的

            int sheetCount = workbook.getNumberOfSheets(); //Sheet的数量
            /**
             * 设置当前excel中sheet的下标：0开始
             */
            Sheet sheet = workbook.getSheetAt(0);   //遍历第一个Sheet

            // 为跳过第一行目录设置count
            int count = 0;

            for (Row row : sheet) {
                // 跳过第一行的目录
//                if(count == 0){
//                    count++;
//                    continue;
//                }
                // 如果当前行没有数据，跳出循环
                if (row.getCell(0).toString().equals("")) {
                    return;
                }
                String rowValue = "";
                for (Cell cell : row) {
                    if (cell.toString() == null) {
                        continue;
                    }
                    CellType e = cell.getCellTypeEnum();
                    String cellValue = "";
                    switch (e) {
                        case STRING:     // 文本
                            cellValue = cell.getRichStringCellValue().getString() + "#";
                            break;
                        case NUMERIC:    // 数字、日期
//                            if (DateUtil.isCellDateFormatted(cell)) {
//                                cellValue = fmt.format(cell.getDateCellValue()) + "#";
//                            } else {
//                                cell.setCellType(Cell.CELL_TYPE_STRING);
//                                cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#";
//                            }
                            break;
                        case BOOLEAN:    // 布尔型
                            cellValue = String.valueOf(cell.getBooleanCellValue()) + "#";
                            break;
                        case BLANK: // 空白
                            cellValue = cell.getStringCellValue() + " ";
                            break;
                        case ERROR: // 错误
                            cellValue = "错误#";
                            break;
                        case FORMULA:    // 公式
                            // 得到对应单元格的公式
                            //cellValue = cell.getCellFormula() + "#";
                            // 得到对应单元格的字符串
                            cell.setCellType(CellType.STRING);
                            cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#";
                            break;
                        default:
                            cellValue = "#";
                    }
                    //System.out.print(cellValue);
                    rowValue += cellValue;
                }
//                writeSql(rowValue,bw);
                System.out.println(rowValue);
            }
            bw.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            bw.close();
        }
    }

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param in
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(InputStream in, File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(EXCEL_XLS)) {  //Excel 2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {  // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }


}
