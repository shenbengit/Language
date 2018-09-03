import manager.ExcelReaderManager;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.List;

/**
 * 配置完excel地址、需要拼接的列、生成文件地址，直接执行main()方法
 */
public class Main {
    public static void main(String[] args) {
//        ExcelReaderManager manager = new ExcelReaderManager();
//        //id 所在列号
//        manager.setIdColumn(0);
//        //单独解析，解析第几列的内容
//        manager.setTargetColumn(3);
//        manager.setReaderConfig("E:\\strings.xls", "E:\\strings.txt");
//        manager.analyse();

        readWord("E:\\test.docx");
    }

    /**
     * 读取word文档
     *
     * @param filePath
     * @return
     */
    private static void readWord(String filePath) {
        if (filePath.endsWith(".doc")) {
            readDoc(filePath);
        } else if (filePath.endsWith(".docx")) {
            readDocx(filePath);
        } else {
            System.out.println("此文件不是word文件！");
        }
    }

    /**
     * 读取以doc结尾的word文档
     *
     * @param filePath
     */
    private static void readDoc(String filePath) {
        try {
            //载入文档
            FileInputStream fis = new FileInputStream(filePath);
            POIFSFileSystem pfs = new POIFSFileSystem(fis);
            HWPFDocument document = new HWPFDocument(pfs);
            //得到文档的读取范围
            Range range = document.getRange();
            TableIterator iterator = new TableIterator(range);
            while (iterator.hasNext()) {
                Table table = iterator.next();
                //迭代行，默认从0开始
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow row = table.getRow(i);
                    //迭代列，默认从0开始
                    for (int j = 0; j < row.numCells(); j++) {
                        //取得单元格
                        TableCell cell = row.getCell(j);
                        //循环取出单元格内容
                        for (int k = 0; k < cell.numParagraphs(); k++) {
                            Paragraph paragraph = cell.getParagraph(k);
                            String str = paragraph.text().trim();
                            System.out.println(str);
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 读取以docx结尾的word文档
     *
     * @param filePath
     */
    private static void readDocx(String filePath) {
        try {
            //载入文档
            FileInputStream fis = new FileInputStream(filePath);
            XWPFDocument document = new XWPFDocument(fis);
            // 获取所有表格
            List<XWPFTable> list = document.getTables();
            for (XWPFTable table : list) {
                // 获取表格的行
                List<XWPFTableRow> tableRows = table.getRows();
                for (XWPFTableRow tableRow : tableRows) {
                    // 获取表格的每个单元格
                    List<XWPFTableCell> tableCells = tableRow.getTableCells();
                    System.out.println("<string name=\"" + tableCells.get(0).getText() + "\">" + tableCells.get(1).getText() + "</string>");
//                    for (XWPFTableCell cell : tableCells) {
//                        // 获取单元格的内容
//                        String text = cell.getText();
//                        System.out.println(text);
//                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
