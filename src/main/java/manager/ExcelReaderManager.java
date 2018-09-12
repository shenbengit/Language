package manager;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * excel表格读取
 */
public class ExcelReaderManager {
    /**
     * excel表路径
     */
    private String excelPath;
    /**
     * id所在列号
     * 默认:0
     */
    private int idColumn = 0;
    /**
     * 选择拼接语言的所在列号,在isAnalysedAll为false时有效
     */
    private int targetColumn = 0;
    /**
     * 解析完生成结果路径
     */
    private String saveFilePath;
    /**
     * 是否解析所有
     * 默认只解析一个
     */
    private boolean isAnalysedAll = false;

    /**
     * 设置读取信息配置，必须调用的方法
     *
     * @param excelPath    excel表路径
     * @param saveFilePath 解析完生成结果路径
     */
    public void setReaderConfig(String excelPath, String saveFilePath) {
        this.excelPath = excelPath;
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
     * 是否解析所有
     *
     * @param analysedAll true: 会使设置的targetColumn无效
     */
    public void setAnalysedAll(boolean analysedAll) {
        isAnalysedAll = analysedAll;
    }

    /**
     * 开始解析
     */
    public void analyse() {
        if (isEmpty(excelPath) || isEmpty(saveFilePath)) {
            System.out.println("请先设置setReaderConfig(String excelPath,String saveFilePath)");
            return;
        }
        final List<String> list = new ArrayList<String>();
        final List<String> list1 = new ArrayList<String>();
        final List<String> list2 = new ArrayList<String>();
        final List<String> list3 = new ArrayList<String>();
        final List<String> list4 = new ArrayList<String>();
        final List<List<String>> mList = new ArrayList<List<String>>();
        mList.add(list1);
        mList.add(list2);
        mList.add(list3);
        mList.add(list4);
        final List<Integer> integers = new ArrayList<Integer>();
        InputStream is = null;
        try {
            //本地国际化文件地址
            is = new FileInputStream(excelPath);
            /*
             *第二个参数excelTypeEnum，是根据excel表后缀选的 分别为：ExcelTypeEnum.XLS | ExcelTypeEnum.XLSX
             *最好选用ExcelTypeEnum.XLS，经测试，选用ExcelTypeEnum.XLSX会出现解析错误的现象。可以将.xlsx转为.xls后进行解析
             */
            ExcelReader reader = new ExcelReader(is, ExcelTypeEnum.XLS, null, new AnalysisEventListener<List<String>>() {

                public void invoke(List<String> strings, AnalysisContext analysisContext) {
                    if (!isAnalysedAll) {
                        list.add("<string name=\"" + strings.get(idColumn) + "\">" + strings.get(targetColumn) + "</string>");
                    } else {
                        integers.clear();
                        //除id所在的列号，其他列号全部添加到integers
                        for (int i = 0; i < strings.size(); i++) {
                            if (i != idColumn) {
                                integers.add(i);
                            }
                        }
                        //如果列号太多，可以自行添加List<String>至mList中
                        if (integers.size() > 4) {
                            System.out.println("列数已超过5，请单独解析");
                            return;
                        }
                        //根据不同的列拼接，拼接的字符串存放至不同的集合中
                        for (int i = 0; i < integers.size(); i++) {
                            mList.get(i).add("<string name=\"" + strings.get(idColumn) + "\">" + strings.get(integers.get(i)) + "</string>");
                        }

                    }

                }

                public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    if (!list.isEmpty()) {
                        //移除标题
                        list.remove(0);
                    } else {
                        list.clear();
                        //所有数据添加到list里
                        for (List<String> l : mList) {
                            if (!l.isEmpty()) {
                                //移除标题
                                l.remove(0);
                                list.addAll(l);
                                list.add("\n************************************************************************\n");
                            }
                        }
                    }
                    if (list.isEmpty()) {
                        return;
                    }
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

    private boolean isEmpty(String str) {
        return str == null || str.length() == 0;
    }
}
