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
        manager.setReaderConfig("E:\\strings.xls","E:\\strings.txt");
        manager.analyse();
    }
}
