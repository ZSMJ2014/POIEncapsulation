package poi;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;


public class Excel2007Reader extends DefaultHandler{

    private SharedStringsTable sst;//共享字符串表
    private String lastContents;//上一次的内容
    private boolean nextIsString;
    private int sheetIndex=-1;
    private List<String> rowList=new ArrayList<String>();
    private int curRow=0;//当前行
    private int curCol=0;//当前列
    private boolean dateFlag;//日期标识
    private boolean numberFlag;//数字标志
    private boolean isTElement;
    private IRowReader rowReader;
    
    public void setRowReader(IRowReader rowReader){
        this.rowReader=rowReader;
    }
    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始
     * @param filename
     * @param sheetId
     * @throws Exception
     */
    public void processOneSheet(String filename,int sheetId) throws Exception{
        OPCPackage pkg=OPCPackage.open(filename);
        XSSFReader r=new XSSFReader(pkg);
        SharedStringsTable sst=r.getSharedStringsTable();
        XMLReader parser=fetchSheetParser(sst);
        //根据rId或rSheet查找sheet
        InputStream sheet=r.getSheet("rId"+sheetId);
        sheetIndex++;
        InputSource sheetSource=new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
    }
    /**
     * 遍历工作薄中所有的电子表格
     * @param filename
     * @throws Exception
     */
    public void process(String filename)throws Exception{
        OPCPackage pkg=OPCPackage.open(filename);
        XSSFReader r=new XSSFReader(pkg);
        SharedStringsTable sst=r.getSharedStringsTable();
        XMLReader parser=fetchSheetParser(sst);
        Iterator<InputStream> sheets=r.getSheetsData();
        while (sheets.hasNext()) {
            curRow=0;
            sheetIndex++;
            InputStream sheet = (InputStream) sheets.next();
            InputSource sheetSource=new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }
    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException{
        XMLReader parser=XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst=sst;
        parser.setContentHandler(this);
        return parser;
    }
    public void startElement(String uri,String localName,String name,Attributes attributes){
        //c代表单元格
        if("c".equals(name)){
            //如果下一个元素是SST的索引，则将nextIsString标记为true
            String cellType=attributes.getValue("t");
            if("s".equals(name)){
                nextIsString=true;
            }else {
                nextIsString=false;
            }
            //日期格式
            String cellDateType=attributes.getValue("s");
            if("1".equals(cellDateType)){
                dateFlag=true;
                numberFlag=false;
            }else if("2".equals(cellDateType)) {
                dateFlag=false;
                numberFlag=true;
            }else {
                dateFlag=false;
                numberFlag=false;
            }
        }
        //当元素为t时
        if("t".equals(name)){
            isTElement=true;
        }else {
            isTElement=false;
        }
        //置空
        lastContents="";
    }
    public void endElement(String uri,String localName,String name){
        //根据SST的索引值 取出单元格存储的字符串
        if(nextIsString){
            try {
                int index=Integer.parseInt(lastContents);
                lastContents=new XSSFRichTextString(sst.getEntryAt(index)).toString();
            } catch (Exception e) {
                // TODO: handle exception
            }
        }
        //t元素也包含字符串
        if(isTElement){
            String value=lastContents.trim();
            rowList.add(curCol,value);
            curCol++;
            isTElement=false;
        }else if ("v".equals(name)) {
            String value=lastContents.trim();
            value=value.equals("")?" ":value;
            //日期格式处理
            if(dateFlag){
                Date date=HSSFDateUtil.getJavaDate(Double.valueOf(value));
                SimpleDateFormat dateFormat=new SimpleDateFormat("dd/MM/yyyy");
                value=dateFormat.format(date);
            }
            //数字类型处理
            if(numberFlag){
                BigDecimal bd=new BigDecimal(value);
                value=bd.setScale(3,BigDecimal.ROUND_HALF_UP).toString();
            }
            rowList.add(curCol, value);
            curCol++;
        }else {
            //如果标签名称为row，这说明已到行尾，调用optRows()方法
            if(name.equals("row")){
                //rowReader.getRows(sheetIndex, curRow, rowList);
                rowList.clear();
                curRow++;
                curCol=0;
            }
        }
    }
    public void characters(char[] ch,int start,int length){
        lastContents+=new String(ch,start,length);
    }
}
