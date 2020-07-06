package apachePOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadandWriteExcelFile {

	public static void main(String[] args) throws IOException {
		
		File scr= new File("D:\\SOFTWARE TESTING\\Testing Tools\\Selenium\\SeleniumDoc\\poi-4.1.2\\Read&Write.xlsx");
		
		//Load File
		 FileInputStream fis = new FileInputStream(scr);
		 
		 //Load Workbook
		 XSSFWorkbook wb = new XSSFWorkbook(fis);
		 
		 //Load Sheet- Here we are loading first sheet only
		 XSSFSheet sh1= wb.getSheetAt(0);
		 
		/* String data0=sh1.getRow(0).getCell(0).getStringCellValue();
		 System.out.println(data0);

		 String data1=sh1.getRow(0).getCell(1).getStringCellValue();
		 System.out.println(data1);*/
		 
		 int count = sh1.getLastRowNum();
		 System.out.println(count);
		  
		 for (int i=0;i<count;i++) {
			 String data0=sh1.getRow(0).getCell(0).getStringCellValue();
			 System.out.println(data0);
			 
		 }
		 for (int i=0;i<count;i++) {
			 String data1=sh1.getRow(0).getCell(1).getStringCellValue();
			 System.out.println(data1);
			 
		 }
           sh1.getRow(0).createCell(2).setCellValue("abc");
           sh1.getRow(1).createCell(2).setCellValue("def");
           sh1.getRow(2).createCell(2).setCellValue("xyz");
           
           FileOutputStream fout= new FileOutputStream(new File("D:\\SOFTWARE TESTING\\Testing Tools\\Selenium\\SeleniumDoc\\poi-4.1.2\\Read&Write.xlsx"));
           
           wb.write(fout);
           fout.close();
           
		

	}

}
