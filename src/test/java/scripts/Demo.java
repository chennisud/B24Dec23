package scripts;

import org.testng.Reporter;
import org.testng.annotations.Test;

import generic.BaseTest;
import generic.Excel;

public class Demo extends BaseTest{
	@Test
public void testA() {
		String data = Excel.getData(Xl_Path, "Sheet1", 0, 0);
		Reporter.log("XL Data:"+data,true);
	Reporter.log("Demo test",true);
}
}
