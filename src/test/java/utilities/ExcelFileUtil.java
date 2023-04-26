package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	Workbook wb;
	
public ExcelFileUtil(String Excelpath) throws Throwable
	{
		FileInputStream fi=new FileInputStream(Excelpath);
		wb=WorkbookFactory.create(fi);
}
     public int rowCount(String sheetname)
     {
    	 return wb.getSheet(sheetname).getLastRowNum();
     }
   

     public String getCellData(String sheetname,int row,int column)
     {
     	String data ="";
     	if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
     	{
     		int celldata =(int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
     		data =String.valueOf(celldata);
     	}
     	else
     	{
     		data =wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
     	}
     	return data;
     }
     public void setcelldata(String sheetname,int row,int column,String status,String WriteExcel) throws Throwable
     {
    	 Sheet ws=wb.getSheet(sheetname);
    	 Row rownum=ws.getRow(row);
    	 Cell cell=rownum.createCell(column);
    	 cell.setCellValue(status);
    	 if(status.equalsIgnoreCase("pass"))
    	 {
    		 CellStyle style=wb.createCellStyle();
    		 Font font=wb.createFont();
    		 font.setColor(IndexedColors.GREEN.getIndex());
    		 font.setBold(true);
    		 font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    		 style.setFont(font);
    		 rownum.getCell(column).setCellStyle(style);
    	 }
    	 else if(status.equalsIgnoreCase("fail"))
    	 {
    		 CellStyle style=wb.createCellStyle();
    		 Font font=wb.createFont();
    		 font.setColor(IndexedColors.RED.getIndex());
    		 font.setBold(true);
    		 font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    		 style.setFont(font);
    		 rownum.getCell(column).setCellStyle(style);
    	 }
    	 else if(status.equalsIgnoreCase("blocked"))
    	 {
    		 CellStyle style=wb.createCellStyle();
    		 Font font=wb.createFont();
    		 font.setColor(IndexedColors.BLUE.getIndex());
    		 font.setBold(true);
    		 font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    		 style.setFont(font);
    		 rownum.getCell(column).setCellStyle(style);
    	 }
    	 FileOutputStream fo=new FileOutputStream(WriteExcel);
    	 wb.write(fo);
     }
     public static void main(String[] args)throws Throwable {
    	 ExcelFileUtil xl=new ExcelFileUtil("E:\\DataBook.xlsx");
    	 		 int rc=xl.rowCount("EMPDATA");
    	 		 System.out.println(rc);
    	 		 for(int i=1;i<=rc;i++)
    	 		 {
    	 			 String fname=xl.getCellData("EMPDATA",i,0);
    	 			String mname=xl.getCellData("EMPDATA",i,1);
    	 			String lname=xl.getCellData("EMPDATA",i,2);
    	 			String empid=xl.getCellData("EMPDATA",i,3);
    	 			System.out.println(fname+" "+mname+" "+lname+" "+empid);
    	 			xl.setcelldata("EMPDATA",i,4,"pass","E://Results.xlsx");
    	 			//xl.setcelldata("EMPDATA",i,4,"fail","E://Results.xlsx");
    	 			//xl.setcelldata("EMPDATA",i,4,"blocked","E://Results.xlsx");
    	 		 }
    	
    	}
}