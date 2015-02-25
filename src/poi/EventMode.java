package poi;


public class EventMode {

    public static void main(String[] args) throws Exception {
        IRowReader reader=new RowReader();
        ExcelReaderUtil.readExcel(reader, "D:\\E1N3其它类型实体详细元数据表.xls");

    }

}
