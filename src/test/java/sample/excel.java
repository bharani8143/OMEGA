//package sample;
//
//import java.io.File;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.ArrayList;
//import java.util.List;
//
//import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFRow;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class excel {
//	
//	public static void main(String[]  args) throws IOException {
//		File ab= new File("C:\\Users\\infinix\\eclipse-workspace\\sample\\target\\AugTest3.xlsx");
//		FileOutputStream u= new FileOutputStream(ab);
//		XSSFWorkbook book = new XSSFWorkbook();
//		  XSSFSheet sheet = book.createSheet("AugustTesting");
//		  XSSFRow row=sheet.createRow(0);
//		  XSSFCell cell=row.createCell(0);
//		  cell.setCellValue("bharaniDharan ");
////		  book.write(u);
////		  book.close();
//		  List<String> u1=new ArrayList<String>();
//		  u1.add("hey");
//		  u1.add("google");
//		  u1.add("openwhatsapp");
//		  u1.add("hey google");
//		  u1.add("movetonextrow");
//		  for(int i=0;i<u1.size();i++) {
//			  XSSFRow row1=sheet.createRow(i+1);
//			  XSSFCell cell1=row1.createCell(0);
//		  }
//	}
//
//}

package sample;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {
    public static void main(String[] args) throws IOException {
        // Creating the file
        File ab = new File("C:\\Users\\infinix\\eclipse-workspace\\sample\\target\\AugTest3.xlsx");
        FileOutputStream u = new FileOutputStream(ab);

        // Creating workbook and sheet
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("AugustTesting");

        // Creating header row and setting a value
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("bharaniDharan");

        // Data to write in next rows
        List<String> u1 = new ArrayList<String>();
        u1.add("hey");
        u1.add("google");
        u1.add("openwhatsapp");
        u1.add("hey google");
        u1.add("movetonextrow");

        // Loop to create rows and add list values
        for (int i = 0; i < u1.size(); i++) {
            XSSFRow row1 = sheet.createRow(i + 1);
            XSSFCell cell1 = row1.createCell(i);
            cell1.setCellValue(u1.get(i)); // Set cell value from list
        }

        // Write workbook to file and close resources
        book.write(u); // Write content to the file
        book.close();  // Close workbook
        u.close();     // Close FileOutputStream

        System.out.println("File saved successfully at " + ab.getAbsolutePath());
    }
}

