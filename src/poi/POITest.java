package poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.NumberFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormat;



public class POITest {

    private static Connection connection;
    public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException {
//        Class.forName("com.mysql.jdbc.Driver");
//        connection=DriverManager.getConnection("jdbc:mysql://10.0.96.161:3306/nict", "root", "123456");
//        parseExcel("e:\\副本100台虚机信息表.xls");
        NumberFormat numberFormat=new DecimalFormat("#.00");
        System.out.println(numberFormat.format(160.411111));
    }

    private static void parseExcel(String excelFile)throws IOException, SQLException {  
        POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(excelFile));//打开Excel文件  
        HSSFWorkbook wbHssfWorkbook=new HSSFWorkbook(fs);//打开工作薄  
        HSSFSheet sheet=wbHssfWorkbook.getSheetAt(0);//打开工作表  
        java.sql.PreparedStatement statement=connection.prepareStatement("insert into machine(serverId,ip,cpuNum,memory,disk) values(?,?,?,?,?)");  
        HSSFRow row=null;  
        String data=null;  
        for (int i = 1; i <=sheet.getLastRowNum(); i++) {//循环读取每一行  
            row =sheet.getRow(i); 
            statement.setNString(1, row.getCell(3).getStringCellValue());
            statement.setNString(2, row.getCell(5).getStringCellValue());
            statement.setInt(3, 2);
            statement.setString(4, "4G");
            statement.setString(5, "40G");
            statement.executeUpdate();
            statement.clearParameters(); 
        }
        statement.close();
        connection.close();
    }  
}
