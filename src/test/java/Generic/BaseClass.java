package Generic;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.*;
public class BaseClass {
	public static void main(String arg[]) {
		BaseClass bc =new BaseClass();
		bc.writeExcel("abc.xlsx");
		bc.readExcel("abc.xlsx");
	}

	public void readExcel(String fileName) {
		File file=new File(this.getCurrentPath("\\src\\test\\java\\Files\\"+fileName)); 
	        FileInputStream fis;
			try {
				
				fis = new FileInputStream(file);
				XSSFWorkbook wb;
				wb = new XSSFWorkbook(fis);
				XSSFSheet sheet=wb.getSheetAt(0);
				XSSFRow row=sheet.getRow(1);
				XSSFCell cell=row.getCell(1);
				String Value=cell.getStringCellValue();
				System.out.println("Cell Value is" +Value);
				wb.close();
				
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	        
			
	 
	}	
	
	public void writeExcel(String fileName) {
		XSSFWorkbook wb= new XSSFWorkbook();
		XSSFSheet sheet= wb.createSheet();
		XSSFRow row=sheet.createRow(1);
		XSSFCell cell=row.createCell(1);
		cell.setCellValue("SUNIL");
		File file =new File(this.getCurrentPath("\\src\\test\\java\\Files\\"+fileName));
		try {
			FileOutputStream fos= new FileOutputStream(file);
			wb.write(fos);
			wb.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	
	}
	public String getCurrentPath(String extraPath) {
		String currentPath = System.getProperty("user.dir");
		currentPath=currentPath+extraPath;
		System.out.println(currentPath);
		return currentPath;
	}
	
	
	
	
}
