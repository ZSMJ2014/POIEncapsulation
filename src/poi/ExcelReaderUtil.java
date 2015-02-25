package poi;

import java.io.IOException;


public class ExcelReaderUtil {

    public static final String EXCEL03_EXTENSION=".xls";
    public static final String EXCEL07_EXTENSION=".xlsx";
    
    public static void readExcel(IRowReader reader,String fileName) throws Exception{
        if(fileName.endsWith(EXCEL03_EXTENSION)){
            Excel2003Reader excel2003Reader=new Excel2003Reader();
            excel2003Reader.setRowReader(reader);
            excel2003Reader.Process(fileName);
        }else {
            throw new Exception("文件格式错误！");
        }
    }
}
