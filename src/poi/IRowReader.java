package poi;

import java.util.List;


public interface IRowReader {

    public void getRows(String sheetName,int curRow,List<String> rowlist);
}
