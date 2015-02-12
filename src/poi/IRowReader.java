package poi;

import java.util.List;


public interface IRowReader {

    public void getRows(int sheetIndex,int curRow,List<String> rowlist);
}
