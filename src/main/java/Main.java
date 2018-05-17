import manager.ExcelReaderManager;

/**
 * 配置完excel地址、需要拼接的列、生成文件地址，直接执行main()方法
 */
public class Main {
    public static void main(String[] args) {
        ExcelReaderManager manager=new ExcelReaderManager();
        //id 所在列号
        manager.setIdColumn(0);
        //单独解析，解析第几列的内容
        manager.setTargetColumn(2);
        //为true时，setTargetColumn(),无效，如果列数大于5，需要一个一个解析，或者到analyse()中自行添加List<String>至mList中
//        manager.setAnalysedAll(true);
        manager.setReaderConfig("E:\\strings.xls","E:\\strings.txt");
        manager.analyse();
    }
}
