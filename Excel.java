package pipeline;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

public class Excel {

	@Test(priority=1)
	public void Search_Index() throws Exception {
	WebDriver driver = null;
	
	
	
	
	try {
//	FileInputStream inputStream = new FileInputStream("C:\\Users\\j.jeevitha.j.v\\Desktop\\Pay slips\\pipeline\\Resources\\glitter.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet("Quotes");
	XSSFRow row = sheet.createRow(15);
	XSSFCell cell = row.createCell(7);
	cell.setCellValue("pi");

	

	//inputStream.close();
    Date d =new Date();
    SimpleDateFormat s = new SimpleDateFormat("MM_dd_yyyy_HH_mm_ss");
    //System.out.println(s);
    
	FileOutputStream outputStream = new FileOutputStream("C:\\Users\\j.jeevitha.j.v\\Desktop\\Pay slips\\pipeline\\New\\TA-report-"+ s.format(d)+".xlsx");
			

	workbook.write(outputStream);
	workbook.close();
	outputStream.close();

	} catch (Exception e) {
	}
	}
	}

