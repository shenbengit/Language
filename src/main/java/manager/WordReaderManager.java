package manager;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.util.List;

/**
 * word 表格读取
 */
public class WordReaderManager {
    /**
     * word文档路径
     */
    private String wordPath;
    /**
     * id所在列号
     * 默认:0
     */
    private int idColumn = 0;
    /**
     * 选择拼接语言的所在列号
     */
    private int targetColumn = 0;
    /**
     * 解析完生成结果路径
     */
    private String saveFilePath;


    public WordReaderManager(String wordPath, String saveFilePath) {
        this.wordPath = wordPath;
        this.saveFilePath = saveFilePath;
    }

    /**
     * 设置id所在列号
     *
     * @param idColumn id所在列号
     */
    public void setIdColumn(int idColumn) {
        this.idColumn = idColumn;
    }

    /**
     * 设置拼接语言的所在列号
     *
     * @param targetColumn 选择拼接语言的所在列号
     */
    public void setTargetColumn(int targetColumn) {
        this.targetColumn = targetColumn;
    }

    /**
     * 开始解析
     */
    public void analyse() {
        if (wordPath.endsWith(".doc")) {
            readDoc(wordPath);
        } else if (wordPath.endsWith(".docx")) {
            readDocx(wordPath);
        } else {
            System.out.println("此文件不是word文件！");
        }
    }

    /**
     * 读取以doc结尾的word文档
     *
     * @param wordPath
     */
    private void readDoc(String wordPath) {
        FileInputStream fis = null;
        try {
            //载入文档
            fis = new FileInputStream(wordPath);
            POIFSFileSystem pfs = new POIFSFileSystem(fis);
            HWPFDocument document = new HWPFDocument(pfs);
            //得到文档的读取范围
            Range range = document.getRange();
            TableIterator iterator = new TableIterator(range);
            StringBuilder builder = new StringBuilder();
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
                            System.out.println("单元格内容: " + str);
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                assert fis != null;
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 读取以docx结尾的word文档
     *
     * @param wordPath
     */
    private void readDocx(String wordPath) {
        FileInputStream fis = null;
        try {
            //载入文档
            fis = new FileInputStream(wordPath);
            XWPFDocument document = new XWPFDocument(fis);
            // 获取所有表格
            List<XWPFTable> list = document.getTables();
            StringBuilder builder = new StringBuilder();
            for (XWPFTable table : list) {
                // 获取表格的行
                List<XWPFTableRow> tableRows = table.getRows();
                for (XWPFTableRow tableRow : tableRows) {
                    // 获取表格的每个单元格
                    List<XWPFTableCell> tableCells = tableRow.getTableCells();
                    builder.append("<string name=\"").append(tableCells.get(idColumn).getText()).append("\">")
                            .append(tableCells.get(targetColumn).getText()).append("</string>\n");
                }
            }
            writeToFile(builder.toString());
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                assert fis != null;
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 写入保存成文件
     *
     * @param content
     */
    private void writeToFile(String content) {
        //list数据写入到文件里
        FileOutputStream fileOutputStream;
        BufferedWriter bufferedWriter;
        File file = new File(saveFilePath);
        try {
            if (!file.exists()) {
                file.createNewFile();
            }
            fileOutputStream = new FileOutputStream(file, false);
            bufferedWriter = new BufferedWriter(new OutputStreamWriter(fileOutputStream));
            bufferedWriter.write(content);
            bufferedWriter.flush();
            bufferedWriter.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
