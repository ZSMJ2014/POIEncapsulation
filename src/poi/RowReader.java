package poi;

import java.util.List;


public class RowReader implements IRowReader {

    @Override
    public void getRows(String sheetName, int curRow, List<String> rowlist) {
        System.out.print(curRow+" ");
        for(int i=0;i<rowlist.size();i++){
            System.out.print(rowlist.get(i)+" ");
        }
        System.out.println();
    }

}
