package packagefilereader.com;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileWriter_2 {

	public static void main(String[] args) throws IOException {
		
	
	
	XSSFWorkbook wb=new XSSFWorkbook();
	XSSFSheet sheet=wb.createSheet("Student_List");
	
	

// Data Passing Thru ArrayList with Object Array 
	
	ArrayList<Object[]> al= new ArrayList<Object[]>();
	al.add(new Object[]{"Sno","Name","Age"});
	al.add(new Object[]{1,"Karthi",26});
    al.add(new Object[] {2,"Ravi",27});
    al.add(new Object[] {3,"Arjun",26});
    al.add(new Object[] {4,"Ram",21});


    
	int rowvalue=0;
	
//	Row  iterator
	
	for (Object[] rdata : al) 
	{
	XSSFRow row=sheet.createRow(rowvalue++);
	
	
	int cellvalue=0;
	
//	Cell iterator
	
	for (Object cdata :rdata) 
	{
	XSSFCell cell=row.createCell(cellvalue++); 
	
//	I Use InstanceOf Keyword To Check Instance of Cell And TypeCasted
	
  if(cdata instanceof String)
	  cell.setCellValue((String) cdata);
	if(cdata instanceof Integer)
		cell.setCellValue((Integer)cdata);
	if(cdata instanceof Boolean)
		cell.setCellValue((Boolean)cdata);
	
}
}
	String path="\\C:\\Users\\Admin\\Desktop\\StudentList1.xlsx\\";
	FileOutputStream fout=new FileOutputStream(path,true);
	wb.write(fout);
	System.out.println("****************Written Sucessfully*******************");
fout.close();
	}
}

