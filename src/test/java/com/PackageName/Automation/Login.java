package PackageName;

import java.io.IOException;
import java.util.HashMap;

import com.excel.utility.DataUtil;
import com.excel.utility.MyXLSReader;

import POM.LoginPage;

import org.openqa.selenium.WebDriver;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Login extends Base {

	WebDriver driver;
	MyXLSReader excelReader;

	@AfterMethod
	public void tearDown() throws InterruptedException {
		
		
		if(driver!=null) {	
			driver.quit();
		}
		
	}

	@Test(dataProvider = "dataSupplier",priority = 1)
	public void testLogin(HashMap<String, String> hMap) throws IOException {

		if (!DataUtil.isRunnable(excelReader, "LoginTest", "Testcases") || hMap.get("Runmode").equals("N")) {

			throw new SkipException("Skipping the test as the runmode is set to N");

		}

		driver = openBrowser(hMap.get("Browser"));
		LoginPage login = new LoginPage(driver);

		login.login(hMap.get("Username"), hMap.get("Password"));

		String expectedResult = hMap.get("ExpectedResult");

		@SuppressWarnings("unused")
		boolean expectedConvertedResult = false;

		if (expectedResult.equalsIgnoreCase("Failure")) {

			expectedConvertedResult = false;

		} else if (expectedResult.equalsIgnoreCase("Success")) {

			expectedConvertedResult = true;
		}


	}

	@DataProvider
	public Object[][] dataSupplier() {

		Object[][] obj = null;

		try {

			excelReader = new MyXLSReader("src\\test\\java\\resources\\DataSheet.xlsx");
			obj = DataUtil.getTestData(excelReader, "LoginTest", "Data");

		} catch (Exception e) {

			e.printStackTrace();

		}

		return obj;

	}

}
