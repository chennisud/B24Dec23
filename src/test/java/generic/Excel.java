package generic;

import java.io.FileInputStream;
import java.io.FileInputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class Excel {
	public static String getData(String path,String sheet,int r,int c) {
		String v="";
			try {
				Workbook wb = WorkbookFactory.create(new FileInputStream(path));
				v = wb.getSheet(sheet).getRow(r).getCell(c).toString();
				wb.close();
			}
			catch (Exception e) {
			}
		return v;
	}
	
	public static int getRowCount(String path,String sheet) {
		int v=0;
			try {
				Workbook wb = WorkbookFactory.create(new FileInputStream(path));
				v = wb.getSheet(sheet).getLastRowNum();
				wb.close();
			}
			catch (Exception e) {
			}
		return v;
	}

}
