import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 配置完excel地址、需要拼接的列、生成文件地址，直接执行main()方法
 */
public class Main {
    public static void main(String[] args) {
        final List<String> list = new ArrayList<String>();
        InputStream is = null;
        try {
            //本地国际化文件地址
            is = new FileInputStream("E:\\strings.xls");
            /*
             *第二个参数excelTypeEnum，是根据excel表后缀选的 分别为：ExcelTypeEnum.XLS | ExcelTypeEnum.XLSX
             * 最好选用ExcelTypeEnum.XLS，经测试，选用ExcelTypeEnum.XLSX会出现解析错误的现象。可以将.xlsx转为.xls后进行解析
             */
            ExcelReader reader = new ExcelReader(is, ExcelTypeEnum.XLS, null, new AnalysisEventListener<List<String>>() {

                /**
                 * 有多少行就会执行多少次
                 * @param object 列内容
                 * @param context
                 */
                @Override
                public void invoke(List<String> object, AnalysisContext context) {
//                    System.out.println("当前sheet:" + context.getCurrentSheet().getSheetNo() + " 当前行：" + context.getCurrentRowNum() + " data:" + object);

                    /*获取当前行的第几列，然后进行拼接
                     *object.get(0)：为string的id
                     *object.get(2)：中文简体、英文、中文繁体、日文，具体看所在列进行替换
                     */
                    list.add("<string name=\"" + object.get(0) + "\">" + object.get(2) + "</string>");
                }

                /**
                 * 解析完之后执行
                 * @param analysisContext
                 */
                @Override
                public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    //移除标题
                    list.remove(0);
                    FileOutputStream fileOutputStream;
                    BufferedWriter bufferedWriter;
                    File file;
                    try {
                        //解析完生成文件
                        file = new File("E:\\strings.txt");
                        if (!file.exists()) {
                            //noinspection ResultOfMethodCallIgnored
                            file.createNewFile();
                        }
                        fileOutputStream = new FileOutputStream(file, true);
                        bufferedWriter = new BufferedWriter(new OutputStreamWriter(fileOutputStream));
                        for (String str : list) {
                            bufferedWriter.write(str + "\n");
                        }
                        bufferedWriter.flush();
                        bufferedWriter.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    System.out.println("list: " + list.size());
                }
            });

            reader.read();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}
