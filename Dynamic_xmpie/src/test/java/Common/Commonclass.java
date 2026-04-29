package Common;

import java.awt.AWTException;
import java.awt.Dimension;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.openqa.selenium.JavascriptExecutor;
import org.apache.commons.io.FileUtils;
import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.HtmlEmail;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.google.common.base.Function;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import Config.Config;
import Config.Property;
import ExcelFiles.Xls_Reader;
import Suite.OrderFlow;
import Utility.RW_File;
import io.github.bonigarcia.wdm.WebDriverManager;

// Main class starts here
public class Commonclass 
{
	// Custom Excel reader for test data input
	public static Xls_Reader datatable_suite1 = null;

	// WebDriver instance for browser automation (Remote)
	public static RemoteWebDriver d=null; 

	// Counter to track the number of errors during test execution
	public static int ErrorNumber = 0;

	// Strings used to store expected and actual financial values and other data
	public static String SheetName = null;
	public static String SheetNameErrorLog = null;
	public static String TotalFunds1 = null;
	public static String CoopFundused1 = null;
	public static String AvailableFunds1 = null;
	public static String ActualTotalFunds2 = null;
	public static String ActualCoopFundused2 = null;
	public static String ActualAvailableFunds2 = null;
	public static String ActualOverAllTaxValue = null;
	public static String VOupdatedshippingprice = null;

	// Format and store current date-time for unique file naming (e.g., for reports)
	public static DateFormat format =new java.text.SimpleDateFormat("_yyyy-MMM-dd_hh_mm_a");
	public static Date date=new Date();
	public static String Execution_Time =format.format(date);

	// Initialize ExtentReports and start a new test for logging test results
	public static ExtentReports er=new ExtentReports(System.getProperty("user.dir")+"\\ExtentReports\\Log_"+Execution_Time+".html");
	public static ExtentTest et=er.startTest("ACGen5 Dynamic XMPie Pricing Automation :: Test Reports");

	// Variables to manage Excel result output (Apache POI)
	static HSSFWorkbook workbook;
	static HSSFSheet sheet;
	protected static Map<String, Object[]> testresultdata;

	// Initialization of test environment and resources
	public static void initialize() throws Exception {

		// Create a unique name for the CSV result file using release number and current timestamp
		DateFormat dateFormat = new SimpleDateFormat("_yyyy-MMM-dd_h-mm-ss_a");
		Date date = new Date();
		SheetName = "\\"+Config.Script+"_DetailedOutput_"+ dateFormat.format(date)+".csv";
		System.out.println(dateFormat.format(date));

		// Create the result CSV file (utility)
		RW_File.CreateFile();

		// Load test data Excel file
		datatable_suite1 = new Xls_Reader(System.getProperty("user.dir")
				+ "//Testdata//ACGen5_Dynamic_XMPie_OrderFlow.xlsx");

		// Prepare test result sheet and logs before test suite begins
		setupBeforeSuite();
	}


	@SuppressWarnings("deprecation")
	public void stopDriver() throws IOException, InterruptedException 
	{
		d.quit();  // Close the browser
		Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe"); // Kill ChromeDriver process
		Common.Wait.wait5Second();
	}

	// Setup method called before any test is executed, to prepare the test result workbook
	public static void setupBeforeSuite()
			throws FileNotFoundException, IOException, AWTException 
	{
		MouseAdjFunction();  // Custom utility (likely for UI mouse handling or screen resolution)

		String fileName = System.getProperty("user.dir")+"\\Results\\TestResult.xls";
		POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(fileName));
		workbook = new HSSFWorkbook(fileSystem);  // Load existing Excel workbook

		// Create a new sheet for test results based on current timestamp
		DateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd_HH_mm_a");
		Date date = new Date();
		SheetNameErrorLog = ""+dateFormat.format(date);  // Timestamped name for error log
		String SheetName = "TestResults" + dateFormat.format(date);
		System.out.println(" SheetName : "+SheetName);

		// Create a new sheet for test results based on current timestamp
		int sheetcount = workbook.getNumberOfSheets();
		if (sheetcount == 3) {
			sheet = workbook.createSheet("Test Result");

		} else 
		{
			sheet = workbook.createSheet(SheetName);
		}

		// Initialize result data map with column headers
		testresultdata = new LinkedHashMap<String, Object[]>();
		testresultdata.put("1", new Object[] { "Test Step Id", "Action",
				"Expected Result", "Actual Result" });
	}

	public static void setupAfterSuite() throws AWTException {
		MouseAdjFunction();

		// Loop through testresultdata and write each entry as a row in the Excel sheet
		Set<String> keyset = testresultdata.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = testresultdata.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date)
					cell.setCellValue((Date) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}

		try 
		{
			// Save the written Excel workbook to the disk and close the stream
			FileOutputStream out = new FileOutputStream(System.getProperty("user.dir")+"\\Results\\TestResult.xls");
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");

		} 
		catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		}

	}

	/**
	 * Adjusts the mouse cursor position to the bottom-right corner of the screen
	 * if the configuration flag 'IsAdjustMOuse' is set to "Yes".
	 */
	public static void MouseAdjFunction() throws AWTException
	{
		try
		{
			Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
			int width = (int) screenSize.getWidth();
			int height = (int) screenSize.getHeight();
			Robot robot = new Robot();
			if(Config.IsAdjustMOuse.equalsIgnoreCase("Yes"))
			{
				robot.mouseMove(width-30, height);
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	/**
	 * Captures and saves a screenshot of the current browser window.
	 * <p>
	 * The screenshot is only taken if the configuration property {@code TakeScreenShot} is set to "Yes".
	 * The image is saved to a predefined folder path with a unique file name.
	 * </p>
	 */
	public static void captureScreenshot() throws InterruptedException {
		try{
			OrderFlow.ImageNumber = OrderFlow.ImageNumber +1;
			if(Config.TakeScreenShot.equalsIgnoreCase("Yes")){
				// capture screen shot
				//!System.out.println("Enter in to screen shot method");
				File scrFile = ((TakesScreenshot)d).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(scrFile, new File(RW_File.FolderPath+"\\"+"TestStep_"+OrderFlow.TestData1+"_"+OrderFlow.ImageNumber+".jpg"));
			}

		}catch (Exception e){
			// take screen shots
			e.printStackTrace();
		}		
	}
	/**
	 * Formats a numeric string to a specific number of decimal places without commas.
	 */
	public static String Decimalsetting(String Datavalue, String DecimalValue)
			throws InterruptedException, AWTException {
		MouseAdjFunction();
		String Datavalue1 = null;
		if(Datavalue.isEmpty())
		{
			return "";	
		}
		else
		{
			if(DecimalValue.equals("2.0"))
			{
				Datavalue1 = String.format("%.2f", new BigDecimal(Datavalue));
			}
			else if(DecimalValue.equals("3.0"))
			{
				Datavalue1 = String.format("%.3f", new BigDecimal(Datavalue));
			}
			else if(DecimalValue.equals("4.0"))
			{
				Datavalue1 = String.format("%.4f", new BigDecimal(Datavalue));
			}
			else if(DecimalValue.equals("0.0"))
			{
				Datavalue1 = String.format("%.0f", new BigDecimal(Datavalue));
			}
			return Datavalue1;
		}

	}
	/**
	 * Formats a numeric string to a specific number of decimal places with commas as thousand separators.
	 */
	public static String Decimalsetting2(String Datavalue, String DecimalValue)
			throws InterruptedException, AWTException {
		NumberFormat numberFormatter = null;
		String Datavalue1 = null;
		String amountOut;
		String amount = Datavalue;
		String value = null;
		MouseAdjFunction();
		if(Datavalue.isEmpty())
		{
			return "";	
		}
		else
		{
			if(DecimalValue.equals("2.0"))
			{
				Datavalue1 = String.format("%.2f", new BigDecimal(amount));
				String[] amountparts = Datavalue1.split("\\.");
				numberFormatter = NumberFormat.getNumberInstance();
				Double num = Double.parseDouble(amountparts[0]);
				amountOut = numberFormatter.format(num);
				value = amountOut+"."+amountparts[1];
			}
			else if(DecimalValue.equals("3.0"))
			{

				Datavalue1 = String.format("%.3f", new BigDecimal(amount));
				String[] amountparts = Datavalue1.split("\\.");
				numberFormatter = NumberFormat.getNumberInstance();
				Double num = Double.parseDouble(amountparts[0]);
				amountOut = numberFormatter.format(num);
				value = amountOut+"."+amountparts[1];
			}
			else if(DecimalValue.equals("4.0"))
			{
				Datavalue1 = String.format("%.4f", new BigDecimal(amount));
				String[] amountparts = Datavalue1.split("\\.");
				numberFormatter = NumberFormat.getNumberInstance();
				Double num = Double.parseDouble(amountparts[0]);
				amountOut = numberFormatter.format(num);
				value = amountOut+"."+amountparts[1];
				//System.out.println("value : "+value);

			}
		}
		return value;
	}
	/**
	 * Launches a browser based on the Config setting ("GC", "FF", "IE").
	 * Uses WebDriverManager for driver setup, falls back to local ChromeDriver if needed.
	 * Maximizes the window, sets implicit wait, and logs the result.
	 */
	public static void StartBrowser()
			throws InterruptedException, NullPointerException {

		// Method to initialize and launch the browser based on configuration
		try{

			// Custom function to adjust mouse settings (likely screen or UI adjustment)
			MouseAdjFunction();

			// Variable to store the full name of the browser being used
			String browserFullName = "";

			// If browser config is "FF", set up and launch Firefox using WebDriverManager
			if (Config.browser.equals("FF")) {

				WebDriverManager.firefoxdriver().setup();
				d= new FirefoxDriver();
				browserFullName = "Firefox";
			} 

			// If browser config is "IE", set up and launch Edge browser
			else if (Config.browser.equals("IE")) {
				WebDriverManager.edgedriver().setup();
				d= new EdgeDriver();
				browserFullName = "Edge";
			} 

			// If browser config is "GC", attempt to launch Chrome via WebDriverManager
			else if (Config.browser.equals("GC")) {
				try {

					WebDriverManager.chromedriver().setup();
					d=new ChromeDriver();
				}
				catch(Exception e){

					// Fallback: If WebDriverManager fails, use local driver path to launch Chrome manually
					System.setProperty("webdriver.chrome.driver",
							System.getProperty("user.dir")+"\\TestData\\Drivers\\chromedriver.exe");
					d = new ChromeDriver();
				}
				browserFullName = "Chrome";
			}

			// Log success message in the test report indicating browser was opened
			et.log(LogStatus.PASS, "Browser Opened: " + browserFullName);

			// Reset error counter on successful browser launch
			ErrorNumber = 0;

			// Set implicit wait timeout for locating web elements
			d.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));

			// Maximize browser window after opening
			d.manage().window().maximize();

			// Re-run mouse adjustment function, likely for UI element positioning
			MouseAdjFunction();
		}
		catch (Exception e)
		{
			// If any error occurs during browser launch, log it and increment error count
			ErrorNumber = ErrorNumber+1;
			et.log(LogStatus.FAIL, "Browser is Not Opened: " + e.getMessage());
			e.printStackTrace();
		}
	}
	/**
	 * Logs into the admin panel using predefined credentials from the configuration.
	 * Navigates to the admin URL, waits for elements to load, enters credentials,
	 * and verifies login. Handles zoom out for better visibility.
	 */
	// Method to log into the admin panel using predefined credentials
	public static void adminLogin() throws InterruptedException, NullPointerException {
		try {	

			// Logging the beginning of the admin login process to the test report
			et.log(LogStatus.INFO, "Admin Login: Starting login process.");

			// Adjust mouse settings or position (likely a utility function for screen automation)
			MouseAdjFunction();

			// Navigate to the admin panel URL defined in the configuration
			d.get(Config.Adminurl);

			// Log the loaded URL for debugging or traceability
			et.log(LogStatus.INFO, "Navigated to Admin URL: " + Config.Adminurl);

			// Maximize the browser window
			d.manage().window().maximize();
			Thread.sleep(2500); 

			// Initialize a FluentWait to handle dynamic wait conditions and element availability
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			// Use Java Robot to zoom out the browser window (Ctrl + -) for better UI visibility
			for (int k = 0; k < 4; k++) {
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_MINUS);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_MINUS);
			}

			// Wait until the username field is found using a custom wait condition
			waitfl.until(new Function<WebDriver, WebElement>() {
				public WebElement apply(WebDriver driver) {
					return driver.findElement(Property.UserName);
				}
			});

			Common.Wait.wait2Second();

			// Enter the admin username from config
			d.findElement(Property.UserName).sendKeys(Config.AdminNamel1);
			Common.Wait.wait2Second();

			// Enter the admin password from config
			d.findElement(Property.Password).sendKeys(Config.AdminPwdl1);
			Common.Wait.wait2Second();

			// Click the login button
			d.findElement(Property.LoginButton).click();

			// Adjust mouse again after login action (likely scrolls or re-positions screen)
			MouseAdjFunction();

			// Log the success of admin login to the report
			et.log(LogStatus.PASS, "Admin login successful.");

		} catch (Exception e) {

			// If any exception occurs, increment error count and take a screenshot
			ErrorNumber = ErrorNumber + 1;
			captureScreenshot(); // Ensure this method saves screenshot with file path

			// Log failure and exception message in the test report
			et.log(LogStatus.FAIL, "Admin login failed. Error #" + ErrorNumber);
			et.log(LogStatus.ERROR, e.toString());

			// Print stack trace for debugging
			e.printStackTrace();
		}
	}

	/**
	 * Configures general settings including decimal values, tax options, shipping,
	 * handling, and payment settings based on given parameters.
	 */
	public static void DecimalvalueSetting(String DecimalValue, String Tax, String IsShippingTaxable,
			String OrderAmountValue, String Weightdecimalvalue,String Weighttype,String userordershippingorhandlingfee,
			String PaymentSubOpt, String PaymentType,String CalculateTaxCondition, 
			String EnablePromotionsORDiscounts,String FullfilmentShippingOrHandlingFee,String FullfilmentShippingMarkupFee,
			String OrderBase,String EnableZeroAmountOrder, String TestData, String CostCenter,
			String ShipAddSameAsBillAdd, String WeightPerPackage,String OrderType)
					throws InterruptedException {
		try{
			MouseAdjFunction();

			Actions kb = new Actions(d);

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d)
					.withTimeout(Duration.ofSeconds(Config.ElementWaitTime))
					.pollingEvery(Duration.ofSeconds(5))
					.ignoring(NoSuchElementException.class)
					.ignoring(StaleElementReferenceException.class);

			et.log(LogStatus.INFO, "Navigating to General Settings");
			Common.Wait.wait15Second();
			d.findElement(Property.Settings).click();
			Common.Wait.wait2Second();
			d.findElement(Property.GeneralSettings).click();
			Common.Wait.wait2Second();
			d.findElement(Property.CollaseALL).click();

			//*******************Navigating to Accounting Tab*************************//

			d.findElement(Property.SettingAccounting).click();
			d.findElement(By.xpath("//button[normalize-space()='+ Show Advanced']")).click();

			WebElement taxCheckbox = d.findElement(By.xpath("(//input[@type='checkbox'])[3]"));

			// Check whether the checkbox is selected
			if (!taxCheckbox.isSelected()) { 
				d.findElement(Property.Taxenabled).click();
				System.out.println("Tax is Enabled");
				et.log(LogStatus.INFO, "Tax is now enabled");

			} else {
				System.out.println("Tax is Enabled in store");
				et.log(LogStatus.INFO, "Tax is already enabled");
			}

			// Select Tax provider in drop down
			String SelectedTaxProvider = d.findElement(Property.SelectedTaxProvider).getText();

			String[] TaxCalculationCondition = CalculateTaxCondition.split("_");

			if(!SelectedTaxProvider.equals(TaxCalculationCondition[0]))
			{

				d.findElement(Property.TaxProviderDropDown).click();

				Common.Wait.wait2Second();
				String Taxlist = d.findElement(Property.TaxProviderList).getText();

				String[] TaxValues = Taxlist.split("\n");

				for (int i = 0; i < TaxValues.length; i++) {
					if (TaxValues[i].equalsIgnoreCase(TaxCalculationCondition[0])) {
						Actions kb1 = new Actions(d);

						// Move to the top of the list
						kb1.sendKeys(Keys.HOME).perform();
						Thread.sleep(300);

						// Navigate down i times
						for (int j = 0; j < i; j++) {
							kb1.sendKeys(Keys.ARROW_DOWN).perform();
							Thread.sleep(200);
						}
						kb1.sendKeys(Keys.ENTER).perform();
						break;
					}
					et.log(LogStatus.INFO, "Tax Provider"+ CalculateTaxCondition);

				}
			}

			// Enable Calculate tax condition check box

			boolean IsTaxCheckBoxChecked = d.findElement(Property.Taxenabled).isSelected();

			if(TaxCalculationCondition[1].equals("ON"))
			{
				if(IsTaxCheckBoxChecked == true)
				{

				}
				else
				{
					d.findElement(Property.Taxenabled).click();
				}
			}
			else
			{
				if(IsTaxCheckBoxChecked == true)
				{
					d.findElement(Property.Taxenabled).click();
				}

			}

			// Enable promotions or Discounts
			boolean EnablePromotionOrDiscounts = d.findElement(Property.EnablePromotionsORDiscounts).isSelected();

			if(EnablePromotionsORDiscounts.equals("ON"))
			{
				if(EnablePromotionOrDiscounts == true)
				{

				}
				else
				{
					d.findElement(Property.EnablePromotionsORDiscounts).click();
					System.out.println("Promotions is Enabled");
				}
			}
			else
			{
				if(EnablePromotionOrDiscounts == true)
				{
					d.findElement(Property.EnablePromotionsORDiscounts).click();
					System.out.println("Promotions is Enabled");
				}

			}

			int Decimalvalue = Double.valueOf(DecimalValue).intValue();
			int Weightdecimalvalue1 = Double.valueOf(Weightdecimalvalue).intValue();
			int OrderAmountvalue = Double.valueOf(OrderAmountValue).intValue();
			Common.Wait.wait5Second();


			// ***************  Decimal value setting  ************************
			//*******Currency decimal setting *************//
			if(Decimalvalue == 2)  			
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.CurrencyDecimaldropdown));
				d.findElement(Property.CurrencyDecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Currency Decimals is Set= 2");
				et.log(LogStatus.INFO, "Currency Decimal set to 2");
			}
			else if (Decimalvalue == 3)
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.CurrencyDecimaldropdown));
				d.findElement(Property.CurrencyDecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Currency Decimals is Set= 3");
				et.log(LogStatus.INFO, "Currency Decimal set to 3");
			}
			else if(Decimalvalue  == 4)
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.CurrencyDecimaldropdown));
				d.findElement(Property.CurrencyDecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Currency Decimals is Set= 4");
				et.log(LogStatus.INFO, "Currency Decimal set to 4");
			}

			//*******Weight decimal value setting *************
			if(Weightdecimalvalue1 == 2)  			
			{
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.weightdecimaldropdown));
				d.findElement(Property.weightdecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Weight decimal is Set= 2");
				et.log(LogStatus.INFO, "Weight Decimal set to 2");
			}
			else if (Weightdecimalvalue1 == 3)
			{
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.weightdecimaldropdown));
				d.findElement(Property.weightdecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Weight decimal is Set= 3");
				et.log(LogStatus.INFO, "Weight Decimal set to 3");
			}
			else if(Weightdecimalvalue1  == 4)
			{
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.weightdecimaldropdown));
				d.findElement(Property.weightdecimaldropdown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Weight decimal is Set= 4");
				et.log(LogStatus.INFO, "Weight Decimal set to 4");
			}
			//******* Order Amount value setting *************

			if(OrderAmountvalue == 2)  			
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.OrderAmoutDecimal));
				d.findElement(Property.OrderAmoutDecimal).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Order Amount value is Set= 2");
				et.log(LogStatus.INFO, "Order Amount value set to 2");
			}
			else if (OrderAmountvalue == 3)
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.OrderAmoutDecimal));
				d.findElement(Property.OrderAmoutDecimal).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Order Amount value is Set= 3");
				et.log(LogStatus.INFO, "Order Amount value set to 3");
			}
			else if(OrderAmountvalue  == 4)
			{
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.OrderAmoutDecimal));
				d.findElement(Property.OrderAmoutDecimal).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
				System.out.println("Order Amount value is Set= 4");
				et.log(LogStatus.INFO, "Order Amount value set to 4");
			}

			//**** shipping taxable******//
			Common.Wait.wait5Second();
			WebElement option = d.findElement(Property.ShippingisTaxable);
			if (IsShippingTaxable.equalsIgnoreCase("Yes")) {
				if (!option.isSelected()) {
					option.click(); // enable it
					System.out.println("Shipping Taxable is Enabled");
					et.log(LogStatus.INFO, "Shipping taxable enabled");
				} else {
					System.out.println("Shipping Taxable was already Enabled");
					et.log(LogStatus.INFO, "Shipping taxable was already enabled");					
				}
			} else if (IsShippingTaxable.equalsIgnoreCase("No")) {
				if (option.isSelected()) {
					option.click(); // disable it
					System.out.println("Shipping Taxable is Disabled");
					et.log(LogStatus.INFO, "Shipping taxable Disabled");
				} else {
					System.out.println("Shipping Taxable was already Disabled");
					et.log(LogStatus.INFO, "Shipping taxable was already Disabled");	
				}
			}
			Common.Wait.wait2Second();

			//*******************Navigating to Order Management Tab*************************//

			d.findElement(Property.CollaseALL).click();
			Common.Wait.wait5Second();
			d.findElement(Property.SettingOrderManagement).click();
			Common.Wait.wait2Second();
			d.findElement(By.xpath("//button[normalize-space()='+ Show Advanced']")).click();

			// Enable Zero amount order setting

			boolean EnableZeroAmountOrderStatus = d.findElement(Property.EnableZeroAmountOrder).isSelected();
			d.findElement(Property.ShowBillingAddressToZeroAmount).isSelected();
			String[] ZeroAmountOrder = EnableZeroAmountOrder.split("_");

			if(EnableZeroAmountOrderStatus == true && ZeroAmountOrder[0].equals("NO"))
			{
				d.findElement(Property.EnableZeroAmountOrder).click();
				System.out.println("Zero Amount Order is Disabled");
				et.log(LogStatus.INFO, "Zero Amount Order is Disabled");
			}
			else if(EnableZeroAmountOrderStatus == false && ZeroAmountOrder[0].equals("YES"))
			{
				d.findElement(Property.EnableZeroAmountOrder).click();
				System.out.println("Zero Amount Order is Enabled");
				et.log(LogStatus.INFO, "Zero Amount Order is Enabled");
			}

			WebElement Option1 = d.findElement(Property.ShowBillingAddressToZeroAmount);
			if (ZeroAmountOrder[1].equalsIgnoreCase("YES")) {

				if (!Option1.isSelected()) {
					option.click(); // Enable it
					System.out.println("Billing Address for Zero Amount Order is Enabled");
					et.log(LogStatus.INFO, "Billing Address for Zero Amount Order is Enabled");
				}
			} else {
				if (Option1.isSelected()) {
					option.click(); // Disable it
					System.out.println("Billing Address for Zero Amount Order is Disabled");
					et.log(LogStatus.INFO, "Billing Address for Zero Amount Order is Disabled");
				}
			}
			//Save the general settings
			d.findElement(Property.SettingsSave).click();
			System.out.println("General Settings as been saved successfully");
			et.log(LogStatus.INFO, "Saving General settings");
			Common.Wait.wait5Second();
			//Back to setting page
			d.findElement(Property.SettingsBack).click();

			//*******************Shipping Settings********************//

			Common.Wait.wait2Second();
			d.findElement(Property.ShippingSttings).click();
			et.log(LogStatus.INFO, "Navigating to Shipping settings");

			//***Weight Package***
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.Enableweightpackage));	  	
			WebElement weightCheckbox = d.findElement(Property.Enableweightpackage);
			boolean isSelected = weightCheckbox.isSelected();
			String weightValue = WeightPerPackage.trim();

			if (isSelected && weightValue.equalsIgnoreCase("UnCheck")) {
				weightCheckbox.click();
				System.out.println("Weight per Package checkbox was checked. Now Unchecked.");
				et.log(LogStatus.INFO, "Weight per Package checkbox was checked. Now Unchecked.");
			} else if (!isSelected && !weightValue.equalsIgnoreCase("UnCheck")) {
				weightCheckbox.click();
				System.out.println("Weight per Package checkbox was unchecked. Now Checked.");
				et.log(LogStatus.INFO, "Weight per Package checkbox was unchecked. Now Checked.");
			} else {
				System.out.println("No action needed for Weight per Package checkbox.");
				et.log(LogStatus.INFO, "No action needed for Weight per Package checkbox.");
			}
			if(Weighttype.equalsIgnoreCase("KGS")|| Weighttype.equalsIgnoreCase("LBS"))
			{
				//enable weight per package in shipping
				boolean isweightperpackage = d.findElement(Property.Enableweightpackage).isSelected();
				if(isweightperpackage == true)
				{
					// System.out.println("check box already selected");
				}
				else
				{
					d.findElement(Property.Enableweightpackage).click();
				}

			
				//***UPS and USPS***
				if(Weighttype.equals("LBS"))
				{
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.upsdropdown));
					d.findElement(Property.upsdropdown).click();

					WebElement shippingmethodbox = d.findElement(Property.upsnextdayselectbox);
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.upsnextdayselectbox));
					if(shippingmethodbox.isSelected())
					{
						d.findElement(Property.Upsshippingselctbox).click();
						Common.Wait.wait2Second();
					}
					else
					{
						Common.Wait.wait2Second();
						d.findElement(Property.Upsshippingselctbox).click();
						Common.Wait.wait2Second();
					}
				}
				else if (Weighttype.equals("KGS")){
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.Uspsdropdown));

					d.findElement(Property.Uspsdropdown).click();
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.UspsPriorityselectbox));
					WebElement shippingmethodbox = d.findElement(Property.UspsPriorityselectbox);
					if(shippingmethodbox.isSelected())
					{
						d.findElement(Property.Uspsshippingselctbox).click();
						Common.Wait.wait2Second();
					}
					else{
						Common.Wait.wait2Second();
						d.findElement(Property.Uspsshippingselctbox).click();
						Common.Wait.wait2Second();
					}
				}
			}
		/*	
			if(Weighttype.equals("--Select--"))
			{
				//d.findElement(Property.shippingmethodselectbox).click();
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.Uspsdropdown));
				Common.Wait.wait5Second();
				d.findElement(Property.Uspsdropdown).click();

				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.UspsPriorityselectbox));
				WebElement shippingmethodbox1 = d.findElement(Property.UspsPriorityselectbox);
				Common.Wait.wait2Second();
				if(shippingmethodbox1.isSelected())
				{
					d.findElement(Property.Uspsshippingselctbox).click();
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.Uspsshippingselctbox));
				}
				else{
					Common.Wait.wait2Second();
				}
				d.findElement(Property.upsdropdown).click();
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.upsdropdown));

				WebElement shippingmethodbox = d.findElement(Property.upsnextdayselectbox);
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.upsnextdayselectbox));

				Common.Wait.wait2Second();
				if(shippingmethodbox.isSelected())
				{
					d.findElement(Property.Upsshippingselctbox).click();
					Common.Wait.wait2Second();
				}
				else{
					Common.Wait.wait2Second();
				}
			}
*/
			//***Shipping Handling***
			WebElement isshippingorhandlingfee = d.findElement(Property.ordershippinghandlingfee);
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.ordershippinghandlingfee));

			if(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||userordershippingorhandlingfee.equals("0.0000"))
			{
				if(isshippingorhandlingfee.isSelected())
				{
					isshippingorhandlingfee.click();
					System.out.println("Order Shipping/Handling Fee is Enabled");
					et.log(LogStatus.INFO, "Order Shipping/Handling Fee is Enabled");

				}
				else
				{
					System.out.println(" Order Shipping/Handling Fee is Disabled");
					et.log(LogStatus.INFO, "Order Shipping/Handling Fee is Disabled");
				}
			}
			else
			{
				waitfl.until(ExpectedConditions.elementToBeClickable((Property.ordershippinghandlingfee)));
				if(isshippingorhandlingfee.isSelected())
				{
					// System.out.println("shipping handling fee check box already selected");
				}
				else
				{
					isshippingorhandlingfee.click();
					System.out.println(" Order Shipping/Handling Fee is Enabled");
					et.log(LogStatus.INFO, "Order Shipping/Handling Fee is Enabled");
				}
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.feeentertextbox));

				d.findElement(Property.feeentertextbox).click();
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				kb.sendKeys(userordershippingorhandlingfee).perform();
			}
			if(OrderBase.equals("Split Ship"))
			{
				if(OrderType.equals("Mailinglist")){
					OrderBaseShippingSetting(OrderBase, TestData, WeightPerPackage);	
				}else{
					OrderBaseShippingSetting(OrderBase, TestData, WeightPerPackage);	
				}
			}else
			{
				OrderBaseShippingSetting(OrderBase, TestData, WeightPerPackage);
			}

			//Save the shipping settings
			d.findElement(Property.SettingsSave).click();
			System.out.println("Shipping Settings as been saved sucessfully");
			et.log(LogStatus.INFO, "Saving shipping settings");
			Common.Wait.wait5Second();
			//Back to setting page
			d.findElement(Property.SettingsBack).click();

			//****Payment Settings****

			Common.Wait.wait5Second();
			d.findElement(Property.PaymentSettings).click();
			et.log(LogStatus.INFO, "Navigating to Payment Settings");
			Common.Wait.wait5Second();

			if(!PaymentType.contains(","))
			{
				paymentTypeSelection(PaymentType, PaymentSubOpt);
			}
			else
			{
				String[] MultiPayments = PaymentType.split(",");
				String[] MultiSubOpts = PaymentSubOpt.split(",");
				int MultiPaymentLen = MultiPayments.length;
				System.out.println("MultiPaymentLen :"+MultiPaymentLen);
				for(int i =0 ; i<MultiPaymentLen; i++)
				{
					// System.out.println("MultiPayments :"+MultiPayments[i]);
					// System.out.println("MultiSubOpts :"+MultiSubOpts[i]);
					paymentTypeSelection(MultiPayments[i], MultiSubOpts[i]);
				}
			}
			//Get cost center check box status
			boolean CostcenterStatus = d.findElement(Property.CostCenterStatus).isSelected();
			// Update the cost center value based on data sheet value (Excel file)
			if(CostCenter.equals("YES") && CostcenterStatus ==  false)
			{
				d.findElement(Property.CostCenterStatus).click();
				Thread.sleep(500);
				d.findElement(Property.DisplayAllCostCenterYes).click();
				System.out.println("Cost Center is Enabled");
				et.log(LogStatus.INFO, "Cost Center is Enabled");
			}
			else if(CostCenter.equals("NO") && CostcenterStatus == true)
			{
				d.findElement(Property.CostCenterStatus).click();
				System.out.println("Cost Center is Disabled");
				et.log(LogStatus.INFO, "Cost Center is Disabled");
			}

			// Shipping Address same as Billing Address
			Thread.sleep(500);

			WebElement toggle = d.findElement(Property.ShipAddSameAsBill);
			boolean isSelected1 = toggle.isSelected(); // current state

			if (ShipAddSameAsBillAdd.equalsIgnoreCase("YES") && !isSelected1) {
				toggle.click(); // Turn it ON
				System.out.println("Shipping Address is same as Billing Address is enabled");
				et.log(LogStatus.INFO, "Shipping Address is same as Billing Address is enabled");
			} else if (ShipAddSameAsBillAdd.equalsIgnoreCase("NO") && isSelected1) {
				toggle.click(); // Turn it OFF
				System.out.println("Shipping Address is NOT same as Billing Address is enabled");
				et.log(LogStatus.INFO, "Shipping Address is NOT same as Billing Address is enabled");
			} else {
				System.out.println("Shipping Address is NOT same as Billing Address is already in Enabled.");
				et.log(LogStatus.INFO, "Shipping Address is NOT same as Billing Address is already in Enabled");
			}

			//save the payment settings
			Thread.sleep(5000);
			d.findElement(Property.SettingsSave).click();
			et.log(LogStatus.INFO, "Saving payment settings");
			System.out.println("Payment Settings as been saved sucessfully");
			Common.Wait.wait5Second();
			//Back to setting page
			d.findElement(Property.SettingsBack).click();

			//BackTo Home
			Common.Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();
			et.log(LogStatus.PASS, "All General settings configured and saved successfully");
			Common.Wait.wait2Second();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Exception occurred while configuring General settings: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Sets the shipping basis (Order or Split Ship) and configures shipping price settings.
	 */
	public static void OrderBaseShippingSetting(String OrderBase, String TestStep, String WeightPerPackage)
			throws InterruptedException , AWTException{
		try{
			MouseAdjFunction();		
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);
			Actions kb = new Actions(d);
			String SelectedText ;
			SelectedText = d.findElement(Property.shippingbasis).getText();
			SelectedText.split("\n");

			if(OrderBase.equals("Order"))
			{
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.shippingbasis));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.shippingbasis));
				d.findElement(Property.shippingbasis).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys("Order").perform();
				kb.sendKeys(Keys.ENTER).perform();
				et.log(LogStatus.INFO, "Shipping Basis: Order");
				Common.Wait.wait2Second();	
				OrderBaseShippingpriceSetting(TestStep,OrderBase,WeightPerPackage);
			}else{

				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.shippingbasis));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.shippingbasis));
				d.findElement(Property.shippingbasis).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys("Split Ship").perform();
				et.log(LogStatus.INFO, "Shipping Basis: Split Ship");
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();	
				OrderBaseShippingpriceSetting(TestStep,OrderBase,WeightPerPackage);
			}
			Common.Wait.wait2Second();	

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Sets the shipping price basis depending on weight per package.
	 */
	public static void OrderBaseShippingpriceSetting(String TestStep,String OrderBase, String WeightPerPackage)
			throws InterruptedException , AWTException{
		try{
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);
			Actions kb = new Actions(d);
			String SelectedText ;
			//  int TestStepvalue1 = Double.valueOf(TestStepvalue).intValue();
			if(!WeightPerPackage.equals("UnCheck"))
			{
				Common.Wait.wait5Second();
				String[] WeightPerpackageOptions =WeightPerPackage.split("_"); 
				waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.Shippingpicebasis));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.Shippingpicebasis));
				d.findElement(Property.Shippingpicebasis).click();
				if(WeightPerpackageOptions[1].equals("Individual"))
				{
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(400);
					kb.sendKeys("Individual").perform();
					kb.sendKeys(Keys.ENTER).perform();
					et.log(LogStatus.INFO, "Shipping Price Basis: Individual");
				}
				else
				{
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(300);
					kb.sendKeys("Consolidated").perform();
					kb.sendKeys(Keys.ENTER).perform();
					et.log(LogStatus.INFO, "Shipping Price Basis: Consolidated");
					Common.Wait.wait5Second();	

					waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.Defaultlocation));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.Defaultlocation));
					d.findElement(Property.Defaultlocation).click();
					Common.Wait.wait2Second();
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(300);
					kb.sendKeys(Keys.DOWN).perform();
					kb.sendKeys(Keys.ENTER).perform();
					Common.Wait.wait5Second();

					waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.DefaultPackageWeight));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.DefaultPackageWeight));
					d.findElement(Property.DefaultPackageWeight).click();
					kb.sendKeys(Keys.HOME).perform();
					kb.sendKeys(WeightPerpackageOptions[2]).perform();
					kb.sendKeys(Keys.ENTER).perform();
				}
			}
			else
			{
				SelectedText = d.findElement(Property.Shippingpicebasis).getText();
				String[] OrderPriceBasis = SelectedText.split("\n");
				//System.out.println(OrderPriceBasis[0]);
				if(!OrderPriceBasis[0].equals("Individual"))
				{
					waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.Shippingpicebasis));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.Shippingpicebasis));
					d.findElement(Property.Shippingpicebasis).click();
					Thread.sleep(2000);
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(300);
					kb.sendKeys(Keys.DOWN).perform();
					kb.sendKeys(Keys.ENTER).perform();
					Thread.sleep(2000);
				}
			}

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Selects a payment method and its sub-options on the payment screen.
	 * Based on the payment type (e.g., Billing, Co-Op Fund, Credit Card, Gift Card),
	 * this method selects the appropriate checkboxes and radio buttons in the UI.
	 * Handles first-level and second-level sub-option selections using provided values.
	 */
	public static void paymentTypeSelection(String PaymentType, String PaymentSubOpt) throws InterruptedException, AWTException
	{
		MouseAdjFunction();
		String[] PaymentSubOptions = PaymentSubOpt.split("_");
		if(PaymentType.equals("Billing"))
		{
			boolean BillingCheckBoxStatus = d.findElement(Property.BillingCheckBox).isSelected();
			if(BillingCheckBoxStatus == false)
			{   
				Common.Wait.wait2Second();
				d.findElement(Property.BillingCheckBox).click();
				Common.Wait.wait2Second();
				d.findElement(Property.BillingSubOption).click();
				Common.Wait.wait2Second();
			}
			System.out.println("Biling is selected as Payment Method");
			et.log(LogStatus.INFO, "Biling is selected as Payment Method");
		}
		else if(PaymentType.equals("Co-Op Fund") || PaymentType.equals("Money on Account"))
		{
			boolean CoopfundCheckBoxStatus = d.findElement(Property.CoOpFund).isSelected();
			if(CoopfundCheckBoxStatus ==  false)
			{
				d.findElement(Property.CoOpFund).click();
				Common.Wait.wait2Second();
			}
			System.out.println("Coop Fund is selected as Payment Method");
			et.log(LogStatus.INFO, "Coop Fund is selected as Payment Method");

			// Sub option selection
			boolean CoopFundSubtionStatus = d.findElement(Property.CoopFundSubOption).isSelected();
			boolean CoopFundMoneyOnAcSQLStatus = d.findElement(Property.MoneyOnAccountSQL).isSelected();
			if(PaymentType.equals("Co-Op Fund") && CoopFundSubtionStatus == false)
			{
				d.findElement(Property.CoopFundSubOption).click();
				Common.Wait.wait2Second();
			}
			// Money on Account SQL option select/ De-select condition
			else if(PaymentType.equals("Money on Account") && CoopFundMoneyOnAcSQLStatus == false)
			{
				d.findElement(Property.MoneyOnAccountSQL).click();
			}

			//Write else part While using "Money on Account SQL"

			// Second Level sub option selection
			if(PaymentSubOptions[0].equals("ExpiryDate"))
			{
				d.findElement(Property.CFExpiryDate).click();
				Common.Wait.wait2Second();
			}
			else if(PaymentSubOptions[0].equals("FIFO"))
			{
				d.findElement(Property.CFFIFO).click();
				Common.Wait.wait2Second();
			}
		}
		else if(PaymentType.equals("Credit Card"))
		{
			boolean CreditCardStatus = d.findElement(Property.CreditCardStatus).isSelected();
			if(CreditCardStatus ==  false)
			{
				d.findElement(Property.CreditCardStatus).click();
				System.out.println("Credit Card is selected as Payment Method");
				et.log(LogStatus.INFO, "Credit Card is selected as Payment Method");
				Common.Wait.wait2Second();
			}

			// Sub option selection
			if(PaymentSubOptions[0].equals("AuthNet"))
			{
				d.findElement(Property.CCAutoNet).click();
				et.log(LogStatus.INFO, "Credit Card > Authroize.Net is selected");
				Common.Wait.wait2Second();
			}
			else if(PaymentSubOptions[0].equals("BrainTree"))
			{
				d.findElement(Property.CCBrainTree).click();
				et.log(LogStatus.INFO, "Credit Card > Braintree is selected");
				Common.Wait.wait2Second();
			}
			else if(PaymentSubOptions[0].equals("PayPal"))
			{
				d.findElement(Property.CCPayPal).click();
				et.log(LogStatus.INFO, "Credit Card > Paypal is selected");
				Common.Wait.wait2Second();
			}
			// Second Level sub option selection
			if(PaymentSubOptions[1].equals("Aonly"))
			{
				d.findElement(Property.CCAonly).click();
				Common.Wait.wait2Second();
			}
			else if(PaymentSubOptions[1].equals("Acap"))
			{
				d.findElement(Property.CCACap).click();
				Common.Wait.wait2Second();
			}
			else if(PaymentSubOptions[1].equals("AChaLat"))
			{
				d.findElement(Property.CCAChaLat).click();
				Common.Wait.wait2Second();
			}

		}
		else if(PaymentType.equals("Gift Card"))
		{
			boolean GiftCardStatus = d.findElement(Property.GiftcardStatus).isSelected();
			if(GiftCardStatus == false)
			{
				d.findElement(Property.GiftcardStatus).click();
				Common.Wait.wait2Second();
				d.findElement(Property.GiftcardSQLRadioButton).click();
				Common.Wait.wait2Second();
			}
			et.log(LogStatus.INFO, "Gift Card is selected as Payment Method");
			System.out.println("Gift Card is selected as Payment Method");
		}
	}
	/**
	 * Selects a second-level sub-option for Credit Card payments.
	 * Sub-options define how the transaction is processed (e.g., authorize only or capture now).
	 * sub option The sub-option to be selected:
	 *              "AOnly" for Authorize Only,
	 *              "ACap" for Authorize and Capture,
	 *              "AChaLat" for Authorize and Charge Later.
	 */
	public static void suboptions(String suboption) throws InterruptedException, AWTException {
		MouseAdjFunction();
		Common.Wait.wait2Second();	
		if(suboption.equals("AOnly"))  // Authorize only
		{
			d.findElement(Property.CCAonly).click();
		}
		else if(suboption.equals("ACap")) //Authorize and Capture
		{
			d.findElement(Property.CCACap).click();
		}
		else if(suboption.equals("AChaLat")) // Authorize and Charge Later
		{
			d.findElement(Property.CCAChaLat).click();
		}
	}

	/**
	 * Updates the pricing details of a product based on quantity, price type (flat/per piece),
	 * and shipping weight/type configuration.
	 *   Navigates to the product pricing section
	 *   Searches for a specific product using {@code Config.ProductPriceCode}
	 *   Edits the price based on quantity range
	 *   Sets either a flat rate or per-piece price
	 * 	 Handles weight and shipping type input
	 */	
	public static void ItemPerPrice(String ItemPerPrice, String FlatRate,String Weighttype,String Weightdecimaltext, String Quantity)
			throws InterruptedException {
		try{
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);											
			waitfl.ignoring(StaleElementReferenceException.class);
			Common.Wait.wait2Second();

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.Product));
			Common.Wait.wait5Second();
			d.findElement(Property.Product).click();
			et.log(LogStatus.INFO, "Navigated to Product section.");
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PricingIcon));
			Common.Wait.wait2Second();
			d.findElement(Property.PricingIcon).click();
			et.log(LogStatus.INFO, "Navigating to Pricing Settings For Product Price");

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PricingSearchBox));
			Common.Wait.wait2Second();
			d.findElement(Property.PricingSearchBox).sendKeys(Config.ProductPriceCode);
			et.log(LogStatus.INFO, "Searched for Product Code: " + Config.ProductPriceCode);
			Common.Wait.wait2Second();
			d.findElement(Property.priceDetailsLink).click();

			BigDecimal inputQty = new BigDecimal(Quantity.trim());

			List<WebElement> rows = d.findElements(By.xpath("//tbody/tr"));

			for (int i = 0; i < rows.size(); i++) {
				WebElement row = rows.get(i);
				List<WebElement> cells = row.findElements(By.tagName("td"));

				if (cells.size() >= 2) {
					BigDecimal minQty = new BigDecimal(cells.get(0).getText().trim());
					String maxQtyText = cells.get(1).getText().trim();

					BigDecimal maxQty;
					if (maxQtyText.equalsIgnoreCase("up")) {
						maxQty = new BigDecimal("999999"); // Use a very high number as a proxy for "up"
					} else {
						maxQty = new BigDecimal(maxQtyText);
					}

					// Compare inputQty with min and max range
					if (inputQty.compareTo(minQty) >= 0 && inputQty.compareTo(maxQty) <= 0) {
						String editBtnXPath = "//tbody/tr[" + (i + 1) + "]/td[7]/button[1]";
						WebElement editButton = d.findElement(By.xpath(editBtnXPath));
						waitfl.until(ExpectedConditions.elementToBeClickable(editButton));
						editButton.click();
						System.out.println("Clicked edit for quantity: " + inputQty);
						break;
					}
				}
			}
			Common.Wait.wait2Second();
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PriceEntertextBox));
			Common.Wait.wait2Second();
			Actions kb = new Actions(d);
			for(int i=1; i<= 1; i++)
			{
				d.findElement(Property.PriceEntertextBox).click();
				Thread.sleep(1000);
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				d.findElement(Property.PriceMinimumQuantity).click();
			}
			Common.Wait.wait2Second();
			double ItemPerPrice1 =  Double.valueOf(ItemPerPrice).doubleValue();
			if(ItemPerPrice1 == 0.0)
			{
				//d.findElement(Property.PriceEntertextBox).sendKeys(FlatRate);
				d.findElement(Property.PriceEntertextBox).click();
				Thread.sleep(500);
				//System.out.println("FlatRate : "+FlatRate);
				kb.sendKeys(FlatRate).perform();
				Common.Wait.wait2Second();
				d.findElement(Property.PriceTypeDropDown).click();
				//select the flat rate price code
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
			}
			else
			{
				d.findElement(Property.PriceEntertextBox).click();
				Thread.sleep(500);
				kb.sendKeys(ItemPerPrice).perform();
				Common.Wait.wait2Second();
				d.findElement(Property.PriceTypeDropDown).click();
				//select the per piece price code
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
			}
			d.findElement(Property.Weightentertextbox).click();
			Thread.sleep(500);
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
			Thread.sleep(500);
			d.findElement(Property.PriceMinimumQuantity).click();
			d.findElement(Property.Weightentertextbox).click();
			Thread.sleep(500);
			if(Weightdecimaltext.equals("0.00")|| Weightdecimaltext.equals("0.000")||Weightdecimaltext.equals("0.0000")||Weighttype.equalsIgnoreCase("--Select--"))
			{
				System.out.println("No shipping weight");
				et.log(LogStatus.INFO, "Weight is not set or skipped (0)");
				kb.sendKeys("0").perform();
			}
			else{ 
				kb.sendKeys(Weightdecimaltext).perform();
				et.log(LogStatus.INFO, "Weight set to: " + Weightdecimaltext);
				String s=d.findElement(Property.selectshippingdropdoen).getText();
				//System.out.println("selected type is"+ s);
				d.findElement(Property.shippingtypedropdown).click();
				Thread.sleep(500);
				kb.sendKeys(Weighttype).perform();
				if(Weighttype.equals(s))
				{
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(300);
					kb.sendKeys(Keys.ARROW_DOWN).perform();
					kb.sendKeys(Keys.ENTER).perform();
				}
				else
				{
					kb.sendKeys(Keys.HOME).perform();
					Thread.sleep(300);
					kb.sendKeys(Keys.ARROW_DOWN).perform();
					kb.sendKeys(Keys.ARROW_DOWN).perform();
					kb.sendKeys(Keys.ENTER).perform();
					Thread.sleep(500);
				}
				et.log(LogStatus.INFO, "Shipping type set to: " + Weighttype);
			}
			d.findElement(Property.PriceDetailsSaveButton).click();
			et.log(LogStatus.PASS, "Price details saved successfully.");
			Common.Wait.wait2Second();

			Common.Wait.wait2Second();
			d.findElement(Property.AdminHomeLink).click();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Pricing setup failed: Unable to assign item price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Configures base pricing and tax settings for a given product type.
	 * 
	 * This method performs the following actions:
	 *   Navigates to the Products Overview page
	 *   Searches for a product using the provided type
	 *   Sets the product offline if it's currently active
	 *   Updates the product code and navigates to pricing section
	 *   Sets the base price (with dynamic handling for specific product types)
	 *   Configures tax exemption based on the provided flag
	 *   Saves the changes and re-activates the product
	 */
	public static void BasePriceSetting(String BasePrice, String ProdutType, String DownloadPrice, String IsTaxExempt)
			throws InterruptedException {
		try{
			//  WebDriverWait wait = new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ProductsOverview));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ProductsOverview));
			Common.Wait.wait5Second();
			d.findElement(Property.ProductsOverview).click();
			et.log(LogStatus.INFO, "Navigated to Products Overview.");

			Actions kb = new Actions(d);	
			Common.Wait.wait2Second();
			d.findElement(Property.ProductsSearchBox).sendKeys(ProdutType);
			Common.Wait.wait5Second();
			d.findElement(Property.EditLink).click();
			Common.Wait.wait10Second();
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ProductStatus));
			WebElement statusElement = d.findElement(Property.ProductStatus);
			// Check if checkbox is selected (checked)
			boolean isActive = statusElement.isSelected();
			System.out.println("Checkbox selected: " + isActive);
			if (isActive) {
				System.out.println("Product is ACTIVE. Proceeding to turn it OFF...");
				// Click the toggle to uncheck
				if (statusElement.isDisplayed() && statusElement.isEnabled()) {
					statusElement.click();
				}

				WebDriverWait wait = new WebDriverWait(d, Duration.ofSeconds(10));
				WebElement alertOk = wait.until(ExpectedConditions.elementToBeClickable(Property.ProductAlertOK));
				alertOk.click();

				Thread.sleep(5000);
				System.out.println("Product has been set to OFFLINE.");
			} else {
				System.out.println("Product is not ACTIVE. No action taken.");
			}
			// Product code update related script
			d.findElement(Property.ProductCodeTextBox).clear();
			Thread.sleep(500);
			d.findElement(Property.ProductCodeTextBox).sendKeys(ProdutType);

			//pricing section
			Thread.sleep(1000);
			d.findElement(Property.CollaseALL).click();
			Thread.sleep(2000); // or Common.Wait.wait2Second();
			d.findElement(Property.ProdPricing).click();
			Thread.sleep(2000); // or Common.Wait.wait2Second();
			for(int i=1; i<= 1; i++)
			{
				WebElement BasePriceText=d.findElement(Property.BasePriceTextBoxDynamic);
				BasePriceText.click();
				Thread.sleep(1000);
				for(int b=1;b<=5;b++) {
					BasePriceText.sendKeys(Keys.BACK_SPACE);		
					}
			//	kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				Thread.sleep(500);
			}
			kb.sendKeys(BasePrice).perform();
			Common.Wait.wait2Second();

			//Provide Download Price
			d.findElement(Property.DownloadPrice).click();
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
			Thread.sleep(500);
			kb.sendKeys(DownloadPrice).perform();
			Thread.sleep(500);

			d.findElement(By.xpath("//button[normalize-space()='+ Show Advanced']")).click();
			for(int i=1; i<= 1; i++)
			{
				boolean taxexmpt=d.findElement(Property.TaxExemptCheckBox).isSelected();

				if(IsTaxExempt.equalsIgnoreCase("YES")){
					if(taxexmpt==true){
						//System.out.println("******do nothing********");
						et.log(LogStatus.INFO, "Tax Exempt set to YES.");
					}else{
						d.findElement(Property.TaxExemptCheckBox).click();
						et.log(LogStatus.INFO, "Tax Exempt set to YES.");
					}

				}else if(IsTaxExempt.equalsIgnoreCase("NO")){
					if(taxexmpt==true){
						//System.out.println("******click on check box****");
						d.findElement(Property.TaxExemptCheckBox).click();
						et.log(LogStatus.INFO, "Tax Exempt set to NO.");
					}else{
						//System.out.println("do nothing");	
						et.log(LogStatus.INFO, "Tax Exempt set to NO.");
					}
				}

			}
			Common.Wait.wait2Second();
			d.findElement(Property.ProductInfoSave).click();
			et.log(LogStatus.INFO, "Product Info saved.");

			Common.Wait.wait10Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ProductStatus));
			d.findElement(Property.ProductStatus).click();
			Common.Wait.wait5Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ProductAlertOK));
			d.findElement(Property.ProductAlertOK).click();
			System.out.println("*Product Alert ok working fine*");
			Common.Wait.wait10Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.AdminHomeLink));
			System.out.println("*Product Home link pre-click *");
			Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();
			System.out.println("*Product Home link fine*");
			Common.Wait.wait10Second();
			System.out.println("Base Pricing Settings as been Done");
			et.log(LogStatus.PASS, "Product Setup completed : " + ProdutType);
		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Product Setup Failed: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Configures a discount value for a promotion by toggling its active state and
	 * entering the specified discount value.
	 *
	 * This method performs the following operations:
	 *   Navigates to the Promotions section
	 *   Searches for the promotion using {@code Config.DiscountName}
	 *   Opens the promotion edit page
	 *   Toggles the discount setting ON or OFF based on input parameters
	 *   Sets the discount value if it's greater than zero
	 *   Saves the changes
	 */
	public static void Discount(String Discount, String DiscountPercentage, String EnablePromotionsORDiscounts)
			throws InterruptedException {
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			if(EnablePromotionsORDiscounts.equals("ON")) {
			Common.Wait.wait5Second();
			d.findElement(Property.PromotionsIconL1).click();
			et.log(LogStatus.INFO, "Navigated to Promotion section for Discounts.");

			Actions kb = new Actions(d);
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PromotionsSearch));
			Common.Wait.wait2Second();
			d.findElement(Property.PromotionsSearch).sendKeys(Config.DiscountName);
			et.log(LogStatus.INFO, "Searched for promotion: " + Config.DiscountName);
			Common.Wait.wait2Second();
			d.findElement(Property.PromotionsEdit).click();
			et.log(LogStatus.INFO, "Clicked edit on promotion");

			// Wait for the toggle to be present and visible
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PromotionDiscountToggle));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PromotionDiscountToggle));

			Common.Wait.wait2Second();

			// Locate the toggle element
			WebElement toggleElement = d.findElement(Property.PromotionDiscountToggle);

			// Check if toggle is currently selected (ON)
			boolean isToggleOn = toggleElement.isSelected();
			System.out.println("Current Promotion Discount Toggle state (isSelected): " + isToggleOn);

			// Determine if toggle needs to be changed
			if (DiscountPercentage.equals("N") && isToggleOn) {
				System.out.println("DiscountPercentage is 'N' and toggle is ON. Turning it OFF...");

				if (toggleElement.isDisplayed() && toggleElement.isEnabled()) {
					toggleElement.click();
					System.out.println("Promotion Discount Toggle has been turned OFF.");
				}

			} else if (!DiscountPercentage.equals("N") && !isToggleOn) {
				System.out.println("DiscountPercentage is NOT 'N' and toggle is OFF. Turning it ON...");

				if (toggleElement.isDisplayed() && toggleElement.isEnabled()) {
					toggleElement.click();
					System.out.println("Promotion Discount has been turned ON.");
				}

			} else {
				System.out.println("Promotion Discount Toggle is already in the desired state. No action taken.");
			}
			WebElement toggle = d.findElement(Property.PromotionDiscountactive);
			boolean toggleStatus = toggle.isSelected(); // Example, adjust based on actual HTML

			if (Discount.equals("0") || Discount.equals("0.00") || Discount.equals("0.000") || Discount.equals("0.0000")) {
				// Discount is zero → ensure toggle is OFF
				if (toggleStatus) {  // toggle is ON, needs to be turned OFF
					toggle.click();
				}
			} else {
				// Discount is non-zero → ensure toggle is ON
				for (int i = 1; i <= 1; i++) {
					d.findElement(Property.PromotionValue).click();
					Thread.sleep(1000);
					kb.sendKeys(Keys.END)
					.keyDown(Keys.SHIFT)
					.sendKeys(Keys.HOME)
					.keyUp(Keys.SHIFT)
					.sendKeys(Keys.BACK_SPACE)
					.perform();
					Thread.sleep(500);
				}

				Common.Wait.wait2Second();
				d.findElement(Property.PromotionValue).click();
				Thread.sleep(500);
				kb.sendKeys(Discount).perform();
				et.log(LogStatus.INFO, "Entered Discount value: " + Discount);
				Common.Wait.wait2Second();

				// Ensure toggle is ON
				if (toggleStatus){  // toggle is OFF, needs to be turned ON
					toggle.click();
				}
			}
			d.findElement(Property.PromotionSave).click();
			Common.Wait.wait2Second();
			et.log(LogStatus.PASS, "Promotion> Discount details saved successfully.");
			d.findElement(Property.AdminHomeLink).click();
			}
		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Promotion> Discount setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Updates the pricing settings for an add-on product based on the provided price values.
	 *
	 * This method performs the following operations:
	 *   Navigates to the Pricing section
	 *   Searches for the add-on using {@code Config.AddonPriceCode}
	 *   Edits the pricing row
	 *   Clears the existing price and inputs a new price
	 *   Selects the appropriate pricing type (flat rate or per piece)
	 *   Saves the changes and returns to the admin home page
	 */
	public static void AddonPrice(String AddonPrice, String AddonPricePerPiece)
			throws InterruptedException {
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PricingIcon));
			Common.Wait.wait2Second();
			d.findElement(Property.PricingIcon).click();
			et.log(LogStatus.INFO, "Navigating to Pricing Settings For ADD-Ons Price");

			Actions kb = new Actions(d);
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PricingSearchBox));
			Common.Wait.wait2Second();
			d.findElement(Property.PricingSearchBox).sendKeys(Config.AddonPriceCode);
			et.log(LogStatus.INFO, "Entered AddonPrice code: " + Config.AddonPriceCode);
			Common.Wait.wait5Second();
			d.findElement(Property.priceDetailsLink).click();
			Thread.sleep(3000);
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PriceEditLink));
			Common.Wait.wait2Second();
			d.findElement(Property.PriceEditLink).click();
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PriceEntertextBox));

			Common.Wait.wait5Second();
			for(int i=1; i<= 1; i++)
			{
				d.findElement(Property.PriceEntertextBox).click();
				Thread.sleep(1000);
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				Thread.sleep(2000);
				d.findElement(Property.PriceMinimumQuantity).click();
			}
			Common.Wait.wait2Second();
			d.findElement(Property.PriceEntertextBox).click();
			Thread.sleep(2000);
			kb.sendKeys(AddonPrice).perform();
			Common.Wait.wait2Second();
			if(AddonPricePerPiece.equals("0") || AddonPricePerPiece.equals("0.00") || AddonPricePerPiece.equals("0.000") || AddonPricePerPiece.equals("0.0000"))
			{
				d.findElement(Property.PriceTypeDropDown).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
			}
			else
			{
				d.findElement(Property.PriceTypeDropDown).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
				kb.sendKeys(Keys.ENTER).perform();
				Thread.sleep(500);
			}
			Common.Wait.wait5Second();
			d.findElement(Property.PriceDetailsSaveButton).click();
			Common.Wait.wait5Second();
			et.log(LogStatus.PASS, "ADD_Ons price setting completed successfully");
			d.findElement(Property.AdminHomeLink).click();
		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Add-on setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Configures the coupon discount settings within the Promotions section.
	 *
	 * This method:
	 *   Navigates to the Promotions section</li>
	 *   Searches for the coupon promotion based on {@code Config.CouponCodeName}
	 *   Edits the promotion coupon details
	 *   Toggles the discount feature ON or OFF depending on the {@code PromotionDiscountPercentage}
	 *   Updates the coupon value field with the provided {@code PromotionCoupon} value
	 *   Saves the changes and returns to the admin home page
	 */
	public static void CouponCodePrice(String PromotionDiscountPercentage, String PromotionCoupon, String EnablePromotionsORDiscounts)
			throws InterruptedException {
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);


			Common.Wait.wait5Second();
			if(EnablePromotionsORDiscounts.equals("ON")) {
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PromotionsIconL1));
			d.findElement(Property.PromotionsIconL1).click();
			et.log(LogStatus.INFO, "Navigated to Promotion section for Coupon.");

			Actions kb = new Actions(d);
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PromotionsSearch));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PromotionsSearch));

			Common.Wait.wait2Second();
			d.findElement(Property.PromotionsSearch).sendKeys(Config.CouponCodeName);
			et.log(LogStatus.INFO, "Searched for promotion: " + Config.CouponCodeName);
			Common.Wait.wait2Second();
			d.findElement(Property.PromotionsEdit).click();
			et.log(LogStatus.INFO, "Clicked edit on promotion coupon");

			// Wait for the toggle to be present and visible
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PromotionDiscountToggle));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PromotionDiscountToggle));

			// Get the current state of the toggle
			WebElement toggleElement = d.findElement(Property.PromotionDiscountToggle);
			boolean isToggleOn = toggleElement.isSelected(); // replaced aria-pressed with isSelected()
			Common.Wait.wait2Second();

			if (PromotionDiscountPercentage.equals("N") && isToggleOn) {
				toggleElement.click(); // Turn OFF
			} else if (!PromotionDiscountPercentage.equals("N") && !isToggleOn) {
				toggleElement.click(); // Turn ON
			}

			WebElement toggle = d.findElement(Property.PromotionDiscountactive);

			// Use isSelected() to get the toggle state
			boolean isToggleOn1 = toggle.isSelected();

			if (PromotionCoupon.equals("0") || PromotionCoupon.equals("0.00") || PromotionCoupon.equals("0.000") || PromotionCoupon.equals("0.0000")) {
				// Coupon is zero → ensure toggle is OFF
				if (isToggleOn1) {  // toggle is ON, needs to be turned OFF
					toggle.click();
				}
			} else {
				for (int i = 1; i <= 1; i++) {
					d.findElement(Property.PromotionCouponValue).click();
					Thread.sleep(1000);
					kb.sendKeys(Keys.END)
					.keyDown(Keys.SHIFT)
					.sendKeys(Keys.HOME)
					.keyUp(Keys.SHIFT)
					.sendKeys(Keys.BACK_SPACE)
					.perform();
				}
				Common.Wait.wait2Second();
				d.findElement(Property.PromotionCouponValue).click();
				Thread.sleep(500);
				kb.sendKeys(PromotionCoupon).perform();
				et.log(LogStatus.INFO, "Entered Discount value: " + PromotionCoupon);
				Common.Wait.wait2Second();
				d.findElement(Property.PromotionDiscountactive).click();
			}
			d.findElement(Property.PromotionSave).click();
			et.log(LogStatus.PASS, "Promotion> Coupon details saved successfully.");
			Common.Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();
			}
		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Promotion> Coupon setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Updates the postage price settings for a specific postage code.
	 *
	 * This method:
	 *   Navigates to the Postage Price Settings section
	 *   Searches for the postage code specified in {@code Config.PostagePriceCode}
	 *   Edits the postage price by clearing the existing value and entering the new {@code Postage} amount
	 *   Saves the updated postage price
	 *   Returns to the Admin Home page
	 */
	public static void PostageSetting(String Postage, String PostagePriceperpiece)
			throws InterruptedException {
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PostageIcon));
			Common.Wait.wait2Second();
			d.findElement(Property.PostageIcon).click();
			et.log(LogStatus.INFO, "Navigating to Postage Price Settings");


			Actions kb = new Actions(d);
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PostageSearch));
			Common.Wait.wait2Second();
			d.findElement(Property.PostageSearch).sendKeys(Config.PostagePriceCode);
			et.log(LogStatus.INFO, "Entered Postage Price code: " + Config.PostagePriceCode);
			Common.Wait.wait5Second();
			d.findElement(Property.PosatgeEditLink).click();
			Thread.sleep(3000);

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PostagePrice));

			Common.Wait.wait5Second();
			for(int i=1; i<= 1; i++)
			{
				d.findElement(Property.PostagePrice).click();
				Thread.sleep(1000);
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				Thread.sleep(2000);
				d.findElement(Property.MinimumQuantity).click();
			}
			Common.Wait.wait2Second();
			d.findElement(Property.PostagePrice).click();
			Thread.sleep(2000);
			kb.sendKeys(Postage).perform();
			Common.Wait.wait2Second();

			d.findElement(Property.PostagePriceSave).click();
			Common.Wait.wait5Second();
			et.log(LogStatus.PASS, "Postage price setting completed successfully");
			d.findElement(Property.AdminHomeLink).click();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Postage Price Setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Updates the tax rate in the application settings.
	 * 
	 * This method performs the following steps:
	 *   Navigates to the TAX Settings section via the main Settings menu
	 *   Clears the existing tax value and inputs the new tax rate provided in {@code Tax}
	 *   Uses Robot class to navigate through the form and trigger the Save button
	 *   Attempts to locate and click the Save button explicitly as a fallback
	 *   Logs the success or failure of the operation
	 *   Returns to the Admin Home page
	 */
	public static void TaxSettings(String Tax)
			throws InterruptedException {
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			d.findElement(Property.Settings).click();
			Common.Wait.wait2Second();
			d.findElement(Property.TaxSettings).click();
			et.log(LogStatus.INFO, "Navigating to TAX Settings");
			Common.Wait.wait5Second();

			Actions kb = new Actions(d);
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.TaxEditLink));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.TaxEditLink));
			Common.Wait.wait5Second();
			d.findElement(Property.TaxEditLink).click();

			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.TaxValue));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.TaxValue));
			Common.Wait.wait2Second();
			for(int i=0; i<=1; i++)
			{
				d.findElement(Property.TaxValue).click();
				Thread.sleep(1000);
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				Thread.sleep(500);
			}
			d.findElement(Property.TaxValue).click();
			Thread.sleep(500);
			// Wait a little for stability
			Common.Wait.wait2Second();

			// Enter tax value using keyboard
			kb.sendKeys(Tax).perform();  // Assuming Tax is the string with value like "9.8"
			Thread.sleep(500);

			Robot robot = new Robot();
			robot.setAutoDelay(500);  // Delay between key events

			// Press TAB to move to Effective Date field
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);

			// Press TAB again to move to Save button
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);

			// Press ENTER to activate Save button
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);

			try {
				WebDriverWait wait = new WebDriverWait(d, Duration.ofSeconds(10));
				WebElement saveBtn = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@title='Save Tax Rate']")));

				// Scroll into view using JavaScript
				((JavascriptExecutor) d).executeScript("arguments[0].scrollIntoView(true);", saveBtn);

				// Click the Save button
				saveBtn.click();
			}
			catch(Exception e) {

			} 
			Common.Wait.wait5Second();
			et.log(LogStatus.PASS, "TAX Settings saved successfully");
			Common.Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "TAX Setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Updates the shipping base price and price per piece settings in the pricing configuration.
	 */
	public static void ShippingPriceSetting(String OrderBaseShipping, String ShippingPricePerPiece)
			throws InterruptedException {
		try{
			Actions kb = new Actions(d);
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));	
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.PricingIcon));
			Common.Wait.wait2Second();
			d.findElement(Property.PricingIcon).click();
			et.log(LogStatus.INFO, "Navigating to Pricing Settings For Shipping Price");

			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PricingSearchBox));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PricingSearchBox));
			//d.findElement(Property.PricingSearchBox).isDisplayed();

			waitfl.until(new Function<WebDriver, WebElement>() 
			{
				public WebElement apply(WebDriver driver) {
					return driver.findElement(Property.PricingSearchBox);
				}
			});

			Common.Wait.wait2Second();
			d.findElement(Property.PricingSearchBox).sendKeys(Config.ShippingPriceCode);
			et.log(LogStatus.INFO, "Entered ShippingPriceCode: " + Config.ShippingPriceCode);
			Common.Wait.wait2Second();
			d.findElement(Property.priceDetailsLink).click();
			Common.Wait.wait2Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PriceEditLink));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PriceEditLink));
			//d.findElement(Property.PriceEditLink).isDisplayed();


			waitfl.until(new Function<WebDriver, WebElement>() 
			{
				public WebElement apply(WebDriver driver) {
					return driver.findElement(Property.PriceEditLink);
				}
			});

			Common.Wait.wait2Second();
			d.findElement(Property.PriceEditLink).click();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.PriceEntertextBox));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.PriceEntertextBox));

			Common.Wait.wait2Second();
			for(int i=1; i<= 1; i++)
			{
				d.findElement(Property.PriceEntertextBox).click();
				Thread.sleep(1000);
				kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
				Thread.sleep(500);
				d.findElement(Property.PriceMinimumQuantity).click();
			}
			Common.Wait.wait2Second();
			d.findElement(Property.PriceEntertextBox).click();
			Thread.sleep(500);
			kb.sendKeys(OrderBaseShipping).perform();
			et.log(LogStatus.INFO, "Entered base shipping price: " + OrderBaseShipping);
			Thread.sleep(500);
			if(ShippingPricePerPiece.equals("0") || ShippingPricePerPiece.equals("0.00") || ShippingPricePerPiece.equals("0.000") || ShippingPricePerPiece.equals("0.0000") )
			{
				d.findElement(Property.PriceTypeDropDown).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
			}
			else
			{
				d.findElement(Property.PriceTypeDropDown).click();
				Common.Wait.wait2Second();
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
				kb.sendKeys(Keys.DOWN).perform();
				Thread.sleep(500);
				kb.sendKeys(Keys.ENTER).perform();
				Thread.sleep(500);
			}

			Common.Wait.wait2Second();
			d.findElement(Property.PriceDetailsSaveButton).click();
			et.log(LogStatus.PASS, "Shipping price setting completed successfully");
			Common.Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Shipping Setup failed: Unable to assign price: " + e.getMessage());
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Configures fulfillment shipping and handling fees along with their fee types (amount or percentage).
	 */
	public static void fullfillmentdetails(String FullfilmentShippingOrHandlingFee,String FullfilmentShippingMarkupFee,
			String FullfilmentShippingOrHandlingFeeTypeAmountPercent,String FullfilmentShippingMarkupFeeAmountPercentTypeAmount)
					throws InterruptedException{
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			d.findElement(Property.Settings).click();
			Common.Wait.wait2Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.FulfillmentLocationsSettings));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FulfillmentLocationsSettings));
			d.findElement(Property.FulfillmentLocationsSettings).click();
			et.log(LogStatus.INFO, "Navigating to Fulfillment Loacations Settings");
			Common.Wait.wait2Second();		

			Actions kb = new Actions(d);            
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.FulFillmentEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FulFillmentEdit));
			d.findElement(Property.FulFillmentEdit).click();

			Thread.sleep(1000);
			d.findElement(Property.CollaseALL).click();
			Thread.sleep(1000);
			d.findElement(By.xpath("//h3[normalize-space()='Shipping Information']"));
			Thread.sleep(1000);

			//shipping/handling Fee
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShiipingHandilingEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShiipingHandilingEdit));
			d.findElement(Property.ShiipingHandilingEdit).click();
			Thread.sleep(1000);
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
			kb.sendKeys(FullfilmentShippingOrHandlingFee).perform();
			et.log(LogStatus.INFO, "Entered FullfilmentShippingOrHandlingFee value: " + FullfilmentShippingOrHandlingFee);

			//Shipping markup fee
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingMarkupEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingMarkupEdit));

			d.findElement(Property.ShippingMarkupEdit).click();
			Thread.sleep(1000);
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();
			kb.sendKeys(FullfilmentShippingMarkupFee).perform();
			et.log(LogStatus.INFO, "Entered FullfilmentShippingMarkupFee value: " + FullfilmentShippingMarkupFee);

			if((FullfilmentShippingOrHandlingFeeTypeAmountPercent.equals("N"))&&(FullfilmentShippingMarkupFeeAmountPercentTypeAmount.equals("N")))
			{
				//select the amount
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingDropDown));
				Common.Wait.wait2Second();
				d.findElement(Property.ShippingDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();

				Common.Wait.wait2Second();
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.MarkupDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.MarkupDropDown));
				d.findElement(Property.MarkupDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
			}
			else if((FullfilmentShippingOrHandlingFeeTypeAmountPercent.equals("Y"))&&(FullfilmentShippingMarkupFeeAmountPercentTypeAmount.equals("Y")))
			{
				//select the Percentage
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingDropDown));
				Common.Wait.wait2Second();
				d.findElement(Property.ShippingDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();

				Common.Wait.wait2Second();
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.MarkupDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.MarkupDropDown));
				d.findElement(Property.MarkupDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
			}
			else if((FullfilmentShippingOrHandlingFeeTypeAmountPercent.equals("Y"))&&(FullfilmentShippingMarkupFeeAmountPercentTypeAmount.equals("N")))
			{
				//select the shipping/handling is Percentage and Markup fee is amount
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingDropDown));
				Common.Wait.wait2Second();
				d.findElement(Property.ShippingDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();

				Common.Wait.wait2Second();
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.MarkupDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.MarkupDropDown));
				d.findElement(Property.MarkupDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();
			}
			else
			{
				//select the shipping/handling is Amount and Markup fee is Percentage
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingDropDown));
				Common.Wait.wait2Second();
				d.findElement(Property.ShippingDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second();

				Common.Wait.wait2Second();
				waitfl.until(ExpectedConditions.elementToBeClickable(Property.MarkupDropDown));
				waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.MarkupDropDown));
				d.findElement(Property.MarkupDropDown).click();
				kb.sendKeys(Keys.HOME).perform();
				Thread.sleep(300);
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ARROW_DOWN).perform();
				kb.sendKeys(Keys.ENTER).perform();
				Common.Wait.wait2Second(); 
			}
			Common.Wait.wait2Second();
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.FullfillmentSave));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FullfillmentSave));

			waitfl.until(new Function<WebDriver, WebElement>() 
			{
				public WebElement apply(WebDriver driver) {
					return driver.findElement(Property.FullfillmentSave);
				}
			});

			d.findElement(Property.FullfillmentSave).click();
			et.log(LogStatus.PASS, "Fulfillment Settings saved successfully");
			Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();
			Wait.wait5Second();

		}catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			et.log(LogStatus.FAIL, "Fulfillment settings Setup failed: Unable to assign price:" + e.getMessage());
			captureScreenshot();
			e.printStackTrace();}
	}
	/**
	 * Clears and deletes fulfillment shipping and handling fee details.
	 */
	public static void deleteingfullfillmentdetails(String FullfilmentShippingOrHandlingFee,String FullfilmentShippingMarkupFee,
			String FullfilmentShippingOrHandlingFeeTypeAmountPercent,String FullfilmentShippingMarkupFeeAmountPercentTypeAmount)
					throws InterruptedException{
		try{
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			MouseAdjFunction();
			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			d.findElement(Property.Settings).click();
			Common.Wait.wait2Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.FulfillmentLocationsSettings));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FulfillmentLocationsSettings));
			d.findElement(Property.FulfillmentLocationsSettings).click();
			Common.Wait.wait2Second();		


			Actions kb = new Actions(d);            
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.FulFillmentEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FulFillmentEdit));
			d.findElement(Property.FulFillmentEdit).click();

			Thread.sleep(1000);
			d.findElement(Property.CollaseALL).click();
			Thread.sleep(1000);
			d.findElement(By.xpath("//h3[normalize-space()='Shipping Information']")).click();
			Thread.sleep(1000);

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShiipingHandilingEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShiipingHandilingEdit));

			Common.Wait.wait2Second();	
			d.findElement(Property.ShiipingHandilingEdit).click();
			Thread.sleep(1000);
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShippingMarkupEdit));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ShippingMarkupEdit));
			d.findElement(Property.ShippingMarkupEdit).click();
			Thread.sleep(1000);
			kb.sendKeys(Keys.END).keyDown(Keys.SHIFT).sendKeys(Keys.HOME).keyUp(Keys.SHIFT).sendKeys(Keys.BACK_SPACE).perform();

			waitfl.until(ExpectedConditions.elementToBeClickable(Property.FullfillmentSave));
			waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.FullfillmentSave));
			d.findElement(Property.FullfillmentSave).click();
			Wait.wait5Second();
			d.findElement(Property.AdminHomeLink).click();
			Wait.wait5Second(); 
		}
		catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();}
	}
	/**
	 * Logs in the user by navigating to the user URL,
	 * entering the username and password, and clicking the login button.
	 * Logs success or failure of the login process.
	 */
	public static void userLogin()
			throws InterruptedException, IOException {
		try{
			MouseAdjFunction();
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			d.get(Config.Userurl);
			et.log(LogStatus.INFO, "Navigated to User URL: " + Config.Userurl);
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.username));
			Common.Wait.wait2Second();

			Wait.wait5Second();
			d.findElement(Property.username).click();
			d.findElement(Property.username).sendKeys(Config.UserNamel1);
			d.findElement(Property.Password).sendKeys(Config.UserPwdl1);

			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.UserLoginButton));
			d.findElement(Property.UserLoginButton).click();
			et.log(LogStatus.PASS, "User login successful.");

			Common.Wait.wait5Second();
		}
		catch(Exception e)
		{
			captureScreenshot();
			e.printStackTrace();
			et.log(LogStatus.FAIL, "user login failed. Error #" + ErrorNumber);
		}

	}
	/**
	 * Calculates the subtotal price based on item price, flat rate, quantity, discount,
	 * and base price, adjusting according to the order amount format and order type.
	 */
	public static String SubTotalCalculation(String ItemPerPrice, String FlatRate,
			String Quantity1, String DiscountCalculationFromSubTotal, String OrderAmountValue,
			String BasePrice, String OrderType)
	{
		double Itemprice = 0;
		String ItemPrice3 = "";
		try
		{
			double ItemPerPrice1 = Double.parseDouble(ItemPerPrice);
			double FlatRate1 = Double.parseDouble(FlatRate);
			double Quantity11 = Double.parseDouble(Quantity1);
			double Discount = Double.parseDouble(DiscountCalculationFromSubTotal);
			double Subtotal = 0;

			if(FlatRate1 > 0)
			{
				Subtotal = FlatRate1 - Discount;
				//System.out.println("A");
			}
			else
			{
				Subtotal = (ItemPerPrice1 * Quantity11) - Discount;
				//	System.out.println("B");
			}
			String Subtotal1 = ""+Subtotal;
			String Subtotal2 = Decimalsetting(Subtotal1, OrderAmountValue);
			String Subtotal3 = Subtotal2;

			Subtotal = Double.parseDouble(Subtotal3);
			double BasePrice1 = Double.parseDouble(BasePrice);

			if(Subtotal >= BasePrice1)
			{
				Itemprice = Subtotal+Discount;
				//System.out.println("C");
			}
			else
			{
				Itemprice = BasePrice1;
				OrderFlow.BasePriceIncrementValue = 1;
				//System.out.println("D");
			}

			Itemprice = Itemprice+0;// No value for zero
			//System.out.println("F");

			String ItemPrice2 = Decimalsetting(""+Itemprice, OrderAmountValue);
			ItemPrice3 = ItemPrice2;
			// System.out.println("ItemPrice3 :"+ItemPrice3);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ItemPrice3;
	}
	/**
	 * Verifies if the displayed addon price matches the expected addon price
	 * based on the given addon quantities and prices per piece.
	 */
	public static void AddonPriceverify(String Addons, String AddonPricePerPiece, 
			String TestStep, String Parameters, String OrderType) throws InterruptedException {
		try{
			MouseAdjFunction();
			String ActualAddonPrice; 
			if(OrderType.equals("Mailinglist")) {
				ActualAddonPrice = d.findElement(Property.AddonpriceB1).getText();
			}
			else {
				ActualAddonPrice = d.findElement(Property.Addonprice).getText();
			}
			if(AddonPricePerPiece.equals("0") || AddonPricePerPiece.equals("0.00") || AddonPricePerPiece.equals("0.000") || AddonPricePerPiece.equals("0.0000"))
			{
				String ExpectedAddonPrice = null;
				if(Addons.equals("0") || Addons.equals("0.00") || Addons.equals("0.000") || Addons.equals("0.0000"))
				{
					ExpectedAddonPrice = "Addon1"; 
				}
				else
				{
					ExpectedAddonPrice = Config.Currency+Addons;
				}
				if(ExpectedAddonPrice.equals(ActualAddonPrice))
				{
					System.out.println("Both addon prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- View Summary Both Addon prices are different ----->"+ErrorNumber);
					System.out.println("Actual Addon price is: "+ActualAddonPrice);
					System.out.println("Expected Addon price is: "+ExpectedAddonPrice); 
				}
			}
			else
			{
				String ExpectedAddonPrice1 = null;
				if(Addons.equals("0") || Addons.equals("0.00") || Addons.equals("0.000") || Addons.equals("0.0000"))
				{
					ExpectedAddonPrice1 = "Addon1";
				}
				else
				{
					ExpectedAddonPrice1 = Config.Currency+AddonPricePerPiece;
				}
				if(ExpectedAddonPrice1.equals(ActualAddonPrice))
				{
					/**/System.out.println("Both addon prices are same");
					et.log(LogStatus.INFO,"Both addon prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- View Summary Both Addon prices are different ----->"+ErrorNumber);
					System.out.println("Actual Addon price is: "+ActualAddonPrice);
					System.out.println("Expected Addon price is: "+ExpectedAddonPrice1);  
				}
			}
		}catch (Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Validates various pricing details displayed on the View Summary page against expected values.
	 * This includes quantity, item price, subtotal, discounts (amount or percentage), addons, postage, and total price.
	 * It handles different order types, such as "Mailinglist", and supports scenarios with or without promotions/discounts.
	 * Screenshots are captured on any verification failure, and errors are logged accordingly.
	 */
	public static void ViewSummaryPriceInformation(String Quantity, String ItemPerPrice, String SubTotal, String Discount, String Addons,
			String Postage, String DownloadPrice, String TotalPrice, String DiscountPercentage, String DiscountCalculationFromSubTotal,
			String OrderType, String TestStep, String Parameters, String ProdutType,
			String OrderBase, String EnablePromotionsORDiscounts,String Weighttype,String DiscountcalculationfromSubTotal,
			String Priceafterapplyingfulfillmentshippingmarkupfee,String OrderAmountValue, String DecimalValue, 
			int BasePriceIncrementValue) throws InterruptedException {
		try{
			MouseAdjFunction();
			String ExpectedItemPrice = Config.Currency+ItemPerPrice;
			String ExpectedSubTotal = Config.Currency+SubTotal;
			String ExpectedDownloadPrice = Config.Currency+DownloadPrice;
			String ExpectedDiscount = null;
			String ExpectedDiscountpercentage = null;
			// For scroll down to screen, we click on total amount attribute (Screen shot purpose)
			Wait.wait5Second();
			//d.findElement(Property.VSTotalPrice).click();


			String ActualDiscountPercentage= null;
			//We get the expected Discount percentage based on Discount type like Amount or percentage
			if(DiscountPercentage.equals("N"))
			{
				System.out.println("OrderFlow.IsBaseDiscountZero :"+OrderFlow.IsBaseDiscountZero);
				System.out.println("BasePriceIncrementValue :"+BasePriceIncrementValue);
				// This part works if Discount value is amount
				//First if block related to Base price cases only
				if(Discount.equals("0") ||
						Discount.equals("0.00") || Discount.equals("0.000") ||
						Discount.equals("0.0000"))
				{
					ExpectedDiscount = "-"+Config.Currency+Discount;
					if(BasePriceIncrementValue == 1 && OrderFlow.IsBaseDiscountZero == true)
					{
						ExpectedDiscount = "-"+Config.Currency+Discount;
					}
					else if(BasePriceIncrementValue == 1 && OrderFlow.IsBaseDiscountZero == false)
					{
						ExpectedDiscount = "-( "+Config.Currency+Discount+" )";
					}
				}
				else 
				{
					// Execute this code when Discount value is more than zero.;
					ExpectedDiscount = "-( "+Config.Currency+Discount+" )";
				}

			}
			else
			{
				// This part works if Discount value is Percentage

				// Get the expected Discount values
				ExpectedDiscount = "-"+Config.Currency+DiscountCalculationFromSubTotal;
				// Get expected Discount percentage  (We need to verify both amount and percentage values if discount type is percentage)
				ExpectedDiscountpercentage = "("+Discount+" "+Config.PercentageSymbol+")";


				//Below code related to Discount percentage verification in user
				if(EnablePromotionsORDiscounts.equals("ON"))
				{
					if(BasePriceIncrementValue == 0)
					{

						if(!OrderType.equals("Download")) {
							ActualDiscountPercentage = d.findElement(Property.Layout2VSDiscountPercent).getText();
						}
						else {
							ActualDiscountPercentage = d.findElement(Property.Layout2VSDiscountPercentB1).getText();
						}

						// This part execute when admin enable PromotionsORDiscounts check box
						if(ExpectedDiscountpercentage.equals(ActualDiscountPercentage))
						{
							System.out.println("View Summary Both Discount percentages are same");
							et.log(LogStatus.INFO, "View Summary Both Discount percentages are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<---- View Summary Both Discount percentages are different ------> "+ErrorNumber);
							et.log(LogStatus.ERROR, "<---- View Summary Both Discount percentages are different ------> "+ErrorNumber);
							System.out.println("Actual Discount percentage is : "+ActualDiscountPercentage);
							System.out.println("Expected Discount percentage is is : "+ExpectedDiscountpercentage);
						}
					}
				}
			}
			//System.out.println(ActualDiscountPercentage);
			String ExpectedAddons = Config.Currency+Addons;
			//System.out.println("expected shipping in view summary is" +ExpectedShippingPrice);
			String ExpectedPostage = Config.Currency+Postage;
			String ExpectedTotal = Config.Currency+TotalPrice;
			String ActualQuantity = null;
			// Get Actual quantity value from View summary page based on layouts

			if(OrderType.equals("Mailinglist")) {	
				ActualQuantity = d.findElement(Property.VSTQuantityB1).getText();
			}
			else if(OrderType.equals("Download")) {
				ActualQuantity = d.findElement(Property.VSTQuantityB2).getText();
			}
			else {
				ActualQuantity = d.findElement(Property.VSTQuantity).getText();
			}

			int ActualQuantity1 = Double.valueOf(ActualQuantity).intValue();
			int Quantity1 = Double.valueOf(Quantity).intValue();
			System.out.println("Quantity: "+Quantity1);
			// Compare the Expected Quantity value with Actual Quantity value. 
			if(Quantity1 == ActualQuantity1)
			{
				System.out.println("View Summary Both Quantitys are same ");
				et.log(LogStatus.INFO, "View Summary Both Quantity are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<---- View Summary Both Quantitys are different ----> "+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- View Summary Both Quantity are different ------>"+ErrorNumber);
				System.out.println("Actual quantity is : "+ActualQuantity);
				System.out.println("Expected quantity is : "+Quantity);
			}


			String ActualItemPerPrice = "";
			// We get the Item price value for the remaining conditions
			if(!OrderType.equals("Download")) {
				if(OrderType.equals("Mailinglist")) {
					ActualItemPerPrice = d.findElement(Property.VSItemPriceB1).getText();
				}
				else {
					ActualItemPerPrice = d.findElement(Property.VSItemPrice).getText();
				}
				if(ExpectedItemPrice.equals(ActualItemPerPrice))
				{
					System.out.println("View Summary Both Item prices are same");
					et.log(LogStatus.INFO, "View Summary Both Iem Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- View Summary Both item prices are different ----->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both item prices are different ------>"+ErrorNumber);
					System.out.println("Actual Item price is: "+ActualItemPerPrice);
					System.out.println("Expected Item price is: "+ExpectedItemPrice);
				}
			}

			// Get subtotal value and compared with expected value with Actual value.
			if(!OrderType.equals("Download")) {
				String ActualSubTotal;
				if(OrderType.equals("Mailinglist")) {
					ActualSubTotal = d.findElement(Property.VSSubTotalB1).getText();
				}

				else {
					ActualSubTotal = d.findElement(Property.VSSubTotal).getText();
				}
				if(ExpectedSubTotal.equals(ActualSubTotal))
				{
					System.out.println("View Summary Both Sub totals Prices are same ");
					et.log(LogStatus.INFO, "View Summary Both Sub totals Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- View Summary Both subtotals price are different ----->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both subtotals price are different ------>"+ErrorNumber);
					System.out.println("Actual subtotal is : "+ActualSubTotal);
					System.out.println("Expected subtotal is : "+ExpectedSubTotal);
				}
			}
			// Get Download Price value and compared with expected value with Actual value.
			if(OrderType.equals("Download")){
				String ActualDownloadPrice;
				ActualDownloadPrice = d.findElement(Property.VSDownload).getText();

				if(ExpectedDownloadPrice.equals(ActualDownloadPrice))
				{
					System.out.println("View Summary Both Download Prices are same ");
					et.log(LogStatus.INFO, "View Summary Both Download Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- View Summary Both Download price are different ----->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both Download price are different ------>"+ErrorNumber);
					System.out.println("Actual Download Price is : "+ActualDownloadPrice);
					System.out.println("Expected Download Price is : "+ExpectedDownloadPrice);
				}
			}
			// Get and verify the Discount value in view summary page.
			String ActualDiscount="";
			if(OrderType.equals("Mailinglist")) {
				ActualDiscount = d.findElement(Property.VSDiscountAppliedB1).getText();
			}
			else if(OrderType.equals("Download")) {
				ActualDiscount = d.findElement(Property.VSDiscountAppliedB2).getText();
			}
			else {
				ActualDiscount = d.findElement(Property.VSDiscountApplied).getText();
			}
			if(EnablePromotionsORDiscounts.equals("ON"))
			{
				// This part execute when admin enable promotionsORDiscount check box enables.
				if(ExpectedDiscount.equals(ActualDiscount))
				{
					System.out.println("View Summary Both Discount prices are same ");
					et.log(LogStatus.INFO, "View Summary Both Discount Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<------ View Summary Both Discount prices are different ------>"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both Discount prices are different ------>"+ErrorNumber);
					System.out.println("Actual Discount price is : "+ActualDiscount);
					System.out.println("Expected Discount price is : "+ExpectedDiscount);
				}
			}
			else
			{
				System.out.println("ActualDiscount: "+ActualDiscount);

				// This part execute when admin disable promotionsORDiscount check box enables. And application not displays discount value
				if(ActualDiscount.equals(""))
				{
					System.out.println("View Summary Both Discount prices are same ");
					et.log(LogStatus.INFO, "View Summary Both Discount Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---- View Summary Both Discount prices are different ------>"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both Discount prices are different ------>"+ErrorNumber);
					System.out.println("Actual Discount price is : "+ActualDiscount);
					System.out.println("Expected Discount price is : "+ExpectedDiscount);
				}

			}
			if(!OrderType.equals("Download")) {
				// Get the addon price and compare expected value with actual value
				String ActualAddons;
				if(OrderType.equals("Mailinglist")) {
					ActualAddons = d.findElement(Property.VSAddonosB1).getText();
				}
				else {
					ActualAddons = d.findElement(Property.VSAddonos).getText();
				}
				if(ExpectedAddons.equals(ActualAddons))
				{
					System.out.println("View Summary Both Addon prices are same");
					et.log(LogStatus.INFO, "View Summary Both Addon Prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<------- View Summary Both Addon prices are different ------>"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- View Summary Both Addon prices are different ------>"+ErrorNumber);
					System.out.println("Actual Addon price is : "+ActualAddons);
					System.out.println("Expected Addon price is : "+ExpectedAddons);  
				}

				// Below code related to get Postage price value and verify expected value with actual value
				String ActualPostage;
				if(OrderType.equals("Mailinglist")) {
					ActualPostage = d.findElement(Property.VSPostagePrice).getText();

					// Below code execute when third party shipping is available
					// Else part execute general shipping cases (with out third party)
					if(ExpectedPostage.equals(ActualPostage))
					{
						System.out.println("View Summary Both Postage Price values are same ");
						et.log(LogStatus.INFO, "View Summary Both Postage Prices are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<------View Summary Both Postage values are different ----->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----View Summary Both Postage values are different ------>"+ErrorNumber);
						System.out.println("Actual Postage value is :"+ActualPostage);
						System.out.println("Expected Postage value is :"+ExpectedPostage);
					}
				}
			}
			// Below code related to get total price value and verify expected value with actual value
			String ActualTotal;
			if(OrderType.equals("Mailinglist")) {
				ActualTotal = d.findElement(Property.VSTotalPriceB1).getText();
			}
			else if(OrderType.equals("Download")) {
				ActualTotal = d.findElement(Property.VSTotalPriceB2).getText();
			}
			else {
				ActualTotal = d.findElement(Property.VSTotalPrice).getText();
			}
			// Below code execute when third party shipping is available
			// Else part execute general shipping cases (with out third party)
			if(ExpectedTotal.equals(ActualTotal))
			{
				System.out.println("View Summary Both Total values are same ");
				et.log(LogStatus.INFO, "View Summary Both Total Prices are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<------View Summary Both Total values are different ----->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<----View Summary Both Total values are different ------>"+ErrorNumber);
				System.out.println("Actual total value is :"+ActualTotal);
				System.out.println("Expected total value is :"+ExpectedTotal);
			}

		}catch (Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Verifies the quantity, item price, and total price displayed in the shopping cart.
	 */
	public static void ShoppingCartPriceInformation(String Quantity,String SubTotal, String ItemPerPrice, String Discount, String Addons,
			String TotalPrice, String DiscountPercentage,
			String OrderType, String TestStep, String Parameters, String ProdutType,String OrderBase,String Weighttype,
			String EnablePromotionsORDiscounts,String DiscountcalculationfromSubTotal,String Priceafterapplyingfulfillmentshippingmarkupfee,
			String OrderAmountValue,String DiscountCalculationFromSubTotal,String DecimalValue) throws InterruptedException {
		try{

			MouseAdjFunction();
			String ShoppingCartTotal = d.findElement(Property.SCartTotal).getText();	
			int Quantity1 = Double.valueOf(Quantity).intValue();
			// Generated the expected Item price value
			String ExpectedItemPrice = null;
			if(OrderType.equals("StaticInventoryShipTOMultipleLocations"))
			{// In shopping cart Item price displays "-" value for the above conditions
				ExpectedItemPrice ="-";	
			}
			else
			{// In other all conditions directly Item price value displays.
				ExpectedItemPrice =Config.Currency+ItemPerPrice;	
			}

			String ExpectedTotal = Config.Currency+TotalPrice;

			String ExpectedTotal2 = ExpectedTotal;
			// Get the Actual quantity and compare with expected quantity
			String ActualQuantity = d.findElement(Property.SCartQuantity).getText();
			int ActualQuantity1 = Double.valueOf(ActualQuantity).intValue();
			if(Quantity1 == ActualQuantity1)
			{
				System.out.println("Shopping Cart Both Quantity are same");
				et.log(LogStatus.INFO, "Shopping Cart Both Quantity are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<----- Shopping cart Both quantitys are different ----->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Shopping cart Both quantitys are different ------>"+ErrorNumber);
				System.out.println("Actual quantity is : "+ActualQuantity);
				System.out.println("Expected quantity is : "+Quantity);
			}
			if(!OrderType.equals("Download")) {
				// Get the Item price based on product type conditions
				String ActualItemPerPrice = null;
				ActualItemPerPrice = d.findElement(Property.SCartItemPrice).getText();

				//Compare the Expected item price with actual item price.
				if(ExpectedItemPrice.equals(ActualItemPerPrice))
				{
					System.out.println("Shopping Cart Both Item prices are same");
					et.log(LogStatus.INFO, "Shopping Cart Both Item prices are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- Shopping cart Both item prices are different -------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Shopping cart Both item prices are different ------>"+ErrorNumber);
					System.out.println("Actual Item price is: "+ActualItemPerPrice);
					System.out.println("Expected Item price is: "+ExpectedItemPrice);
				}
			}
			// Compare total value general condition (with out third party shipping)
			if((ExpectedTotal2.equals(ShoppingCartTotal)))
			{
				System.out.println("Shopping Cart Both Total value Price are same");
				et.log(LogStatus.INFO, "Shopping Cart Both Toatl Value price are same");
			}
			else
			{
				//System.out.println("3333");
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<---------- Shopping cart Both Total values are different ---------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Shopping cart Both Total values are different  ------>"+ErrorNumber);
				System.out.println("Actual total value is :"+ShoppingCartTotal);
				System.out.println("Expected total value is :"+ExpectedTotal2);
			}
			// else part completed.

		}catch (Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Verifies multiple applied payment amounts and the remaining balance on the order checkout page.
	 * This method checks whether the displayed applied payment amounts (up to 5) and remaining balance
	 * match the expected values. If any mismatch occurs, a screenshot is captured and an error is logged.
	 */	
	public static void VerifyMultiAppliedAndRemaingPayments(String Payment1Price, String Payment2Price,
			String Payment3Price, 
			String Payment4Price, String Payment5Price, String DecimalValue,
			String PaymentType, int paymentLength) throws InterruptedException {
		try{
			String ActualAppliedPayment1Price  = null;
			String ActualAppliedPayment2Price  = null;
			String ActualAppliedPayment3Price  = null;
			String ActualAppliedPayment4Price  = null;
			String ActualAppliedPayment5Price  = null;
			MouseAdjFunction();
			if(paymentLength >= 1)
			{
				ActualAppliedPayment1Price = d.findElement(Property.AppliedPayment).getText();
			}
			if(paymentLength >= 2)
			{
				ActualAppliedPayment2Price = d.findElement(Property.AppliedPayment2).getText();
			}
			if(paymentLength >= 3)
			{
				ActualAppliedPayment3Price = d.findElement(Property.AppliedPayment3).getText();
				
			}
			if(paymentLength >= 4)
			{
			ActualAppliedPayment4Price = d.findElement(Property.AppliedPayment4).getText();
				
			}
			if(paymentLength == 5)
			{
			ActualAppliedPayment5Price = d.findElement(Property.AppliedPayment5).getText();				
			}

	     	String ActualRemainingPayment = d.findElement(Property.RemainingBalance).getText();	

			String ExpectedAppliedPayment1Price = Config.Currency+Payment1Price;

			String ExpectedAppliedPayment2Price = Config.Currency+Payment2Price;
			String ExpectedAppliedPayment3Price = Config.Currency+Payment3Price;
			String ExpectedAppliedPayment4Price = Config.Currency+Payment4Price;
			String ExpectedAppliedPayment5Price = Config.Currency+Payment5Price;

			String ExpectedRemainingPayment = Config.Currency+DecimalValue;	

			if(paymentLength >= 1)
			{
				if(ExpectedAppliedPayment1Price.equals(ActualAppliedPayment1Price))
				{
					System.out.println("Both Total values are same ");
					et.log(LogStatus.INFO,"Both Total values are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<----- Ordercheckout Both applied payments values are different ---------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
					System.out.println("Actual Applied value is :"+ActualAppliedPayment1Price);
					System.out.println("Expected Applied value is :"+ExpectedAppliedPayment1Price);

				}
			}

			if(paymentLength >= 2)
			{
				if(ExpectedAppliedPayment2Price.equals(ActualAppliedPayment2Price))
				{
					System.out.println("Both Total values are same ");
					et.log(LogStatus.INFO,"Both Total values are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Ordercheckout Both applied payments values are different ----->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
					System.out.println("Actual Applied value is :"+ActualAppliedPayment2Price);
					System.out.println("Expected Applied value is :"+ExpectedAppliedPayment2Price);
				}
			}

			if(paymentLength >= 3)
			{
				if(ExpectedAppliedPayment3Price.equals(ActualAppliedPayment3Price))
				{
					System.out.println("Both Total values are same ");
					et.log(LogStatus.INFO,"Both Total values are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<------- Ordercheckout Both applied payments values are different -------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
					System.out.println("Actual Applied value is :"+ActualAppliedPayment3Price);
					System.out.println("Expected Applied value is :"+ExpectedAppliedPayment3Price);
				}
			}

			if(paymentLength >= 4)
			{
				if(ExpectedAppliedPayment4Price.equals(ActualAppliedPayment4Price))
				{
					System.out.println("Both Total values are same ");
					et.log(LogStatus.INFO,"Both Total values are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<------- Ordercheckout Both applied payments values are different -------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
					System.out.println("Actual Applied value is :"+ActualAppliedPayment4Price);
					System.out.println("Expected Applied value is :"+ExpectedAppliedPayment4Price);

				}
			}

			if(paymentLength == 5)
			{
				if(ExpectedAppliedPayment5Price.equals(ActualAppliedPayment5Price))
				{
					System.out.println("Both Total values are same ");
					et.log(LogStatus.INFO,"Both Total values are same");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<------- Ordercheckout Both applied payments values are different -------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different------>"+ErrorNumber);
					System.out.println("Actual Applied value is :"+ActualAppliedPayment5Price);
					System.out.println("Expected Applied value is :"+ExpectedAppliedPayment5Price);
				}
			}
/*
			if(ExpectedRemainingPayment.equals(ActualRemainingPayment))
			{
				System.out.println("Both Total values are same ");
				et.log(LogStatus.INFO,"Both Total values are same");
			}
			else
			{

				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Ordercheckout Both Remaining payments values are different ---------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Ordercheckout Both Remaining payments values are different------>"+ErrorNumber);
				System.out.println("Actual Remaining value is :"+ActualRemainingPayment);
				System.out.println("Expected Remaining value is :"+ExpectedRemainingPayment);
			}
			*/
		}catch (Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Selects a payment type during checkout and performs required actions based on the payment method.
	 * This method handles different payment types such as Billing, Co-Op Fund, and Credit Card.
	 * It fills in necessary details and clicks the apply button for each type accordingly.
	 * Screenshots are captured if any error occurs during the process.
	 */
	public static void SelectPaymentTypeInCheckOutPage(String OrderBase, String PaymentType, String TestStep,
			String PaymentPrice, String PaymentSubOpt, boolean paymentsListBoxCount,
			int s, int paymentLength1) throws InterruptedException, AWTException
	{
		MouseAdjFunction();
		try{
			Common.Wait.wait2Second();
			//System.out.println("paymentsListBoxCount :"+paymentsListBoxCount);

			Common.Wait.wait10Second();

			if(PaymentType.equals("Billing"))
			{
				d.findElement(Property.BillingPayment).click();
				Common.Wait.wait5Second();
				et.log(LogStatus.INFO, "Payment Method: Billing");
				WebElement BillingAmount=d.findElement(Property.PaymentValue);
				Common.Wait.wait2Second();
				BillingAmount.clear();
				BillingAmount.sendKeys(PaymentPrice);				
				d.findElement(Property.PONumber).sendKeys("124");
				Common.Wait.wait5Second();
				d.findElement(Property.BillingApplyButton).click();

			}
			else if(PaymentType.equals("Co-Op Fund"))
			{
				// No need to do other actions
				d.findElement(Property.COOPPayment).click();
				Common.Wait.wait5Second();
				et.log(LogStatus.INFO, "Payment Method: Co-op Fund");
				d.findElement(Property.coopApplyButton).click();

			}
			else if(PaymentType.equals("Credit Card"))
			{
				d.findElement(Property.CreditCardPayment).click();
				et.log(LogStatus.INFO, "Payment Method: Credit Card");

				System.out.println("** CAME UP TO HERE **");
				Thread.sleep(2000);
				d.findElement(Property.CreditCardNumber).sendKeys("4111111111111111");
				d.findElement(Property.CreditCardNameOnCard).sendKeys("sat");
				d.findElement(Property.CreditCardLastName).sendKeys("SelTest");
				d.findElement(Property.CreditCardCVVNumber).sendKeys("123");
				Common.Wait.wait5Second();
				d.findElement(Property.creditcardApplyButton).click();
				Common.Wait.wait5Second();
			}
		}
		catch(Exception e)
		{
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}
	}

	/**
	 * Verifies that the applied payment amount in the checkout page matches the expected total,
	 * based on various pricing, tax, and order configuration conditions.
	 *
	 * This method handles different tax calculation types (e.g., Vertex), shipping conditions,
	 * and order base types. It performs detailed
	 * checks depending on whether handling/shipping fees are applied and how they are calculated.
	 * If the actual applied payment does not match the expected value, a screenshot is captured
	 * and an error is logged.
	 */

	public static void VerifyAppliedAndRemaingPayments(String Total, String DecimalValue,
			String CalculateTaxCondition,String OrderBase,String Weighttype,String Subtotal,String PromotionCoupon,
			String Addons,String DiscountcalculationfromSubTotal,
			String OrderAmountValue,String userordershippingorhandlingfee, String TotalPrice,
			String IsShippingTaxable, String Tax, String PriceAfterApplyingCoupon, String OrderType) throws InterruptedException {
		try{

			MouseAdjFunction();
			String ActualAppliedPayment;

			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
				ActualAppliedPayment = d.findElement(Property.AppliedPayment).getText();
			} else {
				ActualAppliedPayment = d.findElement(Property.AppliedPaymentB1).getText();
			}

			//String ActualRemainingPayment = d.findElement(Property.RemainingBalance).getText();	

			String ExpectedAppliedPayment = Config.Currency+Total;
			// String ExpectedRemainingPayment = Config.Currency+DecimalValue;	

			String[] TaxType = CalculateTaxCondition.split("_");

			if(TaxType[0].equals("Vertex"))
			{ 
				if(TaxType[1].equals("ON"))
				{	
					String ExpectedAppliedPayment1 =VertexTotal(ExpectedAppliedPayment);
					if(ExpectedAppliedPayment1.equals(ActualAppliedPayment))
					{
						System.out.println("Both Total values are same ");
						et.log(LogStatus.INFO,"Both Total values are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<----- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
						System.out.println("Actual Applied value is :"+ActualAppliedPayment);
						System.out.println("Expected Applied value is :"+ExpectedAppliedPayment1);
					}
				}
				else
				{
					String ExpectedAppliedPayment1 =ExpectedAppliedPayment;
					if(ExpectedAppliedPayment1.equals(ActualAppliedPayment))
					{
						System.out.println("Both Total values are same ");
						et.log(LogStatus.INFO,"Both Total values are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--- Ordercheckout Both applied payments values are different ---------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
						System.out.println("Actual Applied value is :"+ActualAppliedPayment);
						System.out.println("Expected Applied value is :"+ExpectedAppliedPayment1);
					}
				}
			}
			else
			{
				String OCappliedpayment = null;
				String OCmarkupgrandtotal = null;

				if(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
						userordershippingorhandlingfee.equals("0.0000"))
				{
					if(OrderBase.equals("Item")&& (Weighttype.equals("KGS")|| Weighttype.equals("LBS")))
					{
						String OC3rdpartytax = d.findElement(Property.OCTaxAmount).getText();
						String OC3rdpartytax2 = OC3rdpartytax.substring(1,OC3rdpartytax.length());
						String OC3rdpartyshippingprice = d.findElement(Property.OrderCheckoutGridPostage).getText();
						String OC3rdpartyshippingprice2 = OC3rdpartyshippingprice.substring(1,OC3rdpartyshippingprice.length());
						OCappliedpayment = OCgrandtotal(Subtotal,PriceAfterApplyingCoupon,OC3rdpartytax2,OC3rdpartyshippingprice2,
								Addons,DiscountcalculationfromSubTotal,OrderAmountValue, Total, IsShippingTaxable, Tax);
					}
				}
				else
				{
					if(OrderBase.equals("Item")&& Weighttype.equals("KGS")|| Weighttype.equals("LBS"))
					{
						String OCmarkupfee = d.findElement(Property.OCHandilingfee).getText();
						String OCmarkupfee1 = OCmarkupfee.substring(1,OCmarkupfee.length());
						String OC3rdpartytax = d.findElement(Property.OCTaxAmount).getText();
						String OC3rdpartytax1 = OC3rdpartytax.substring(1,OC3rdpartytax.length());
						String OC3rdpartyshippingprice = d.findElement(Property.OrderCheckoutGridPostage).getText();
						String OC3rdpartyshippingprice1 = OC3rdpartyshippingprice.substring(1,OC3rdpartyshippingprice.length());
						OCmarkupgrandtotal = Ocupdatedmarkupgrandtotal(Subtotal,PromotionCoupon,OCmarkupfee1,OC3rdpartytax1,OrderAmountValue,
								Addons,DiscountcalculationfromSubTotal,OC3rdpartyshippingprice1);

					}
				}
				if(Weighttype.equals("--Select--")&&(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
						userordershippingorhandlingfee.equals("0.0000")))
				{
					if(ExpectedAppliedPayment.equals(ActualAppliedPayment))
					{
						System.out.println("Both Total values are same ");
						et.log(LogStatus.INFO,"Both Total values are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Ordercheckout Both applied payments values are different -------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
						System.out.println("Actual Applied value is :"+ActualAppliedPayment);
						System.out.println("Expected Applied value is :"+ExpectedAppliedPayment);

					}
				}
				if(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
						userordershippingorhandlingfee.equals("0.0000"))
				{
					if(OrderBase.equals("Item")&& (Weighttype.equals("KGS")|| Weighttype.equals("LBS")))
					{
						if(OCappliedpayment.equals(ActualAppliedPayment))
						{
							System.out.println("Both Total values are same ");
							et.log(LogStatus.INFO,"Both Total values are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<--------- Ordercheckout Both applied payments values are different ---------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
							System.out.println("Actual Applied value is :"+ActualAppliedPayment);
							System.out.println("Expected Applied value is :"+OCappliedpayment);

						}
					}
				}
				else if(!(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
						userordershippingorhandlingfee.equals("0.0000")))
				{
					if(OrderBase.equals("Item")&&(Weighttype.equals("KGS")|| Weighttype.equals("LBS")))
					{
						if(OCmarkupgrandtotal.equals(ActualAppliedPayment))
						{
							System.out.println("Both Total values are same ");
							et.log(LogStatus.INFO,"Both Total values are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<-------- Ordercheckout Both applied payments values are different --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
							System.out.println("Actual Applied value is :"+ActualAppliedPayment);
							System.out.println("Expected Applied value is :"+OCmarkupgrandtotal);
						}
					}
					else
					{
						if(ExpectedAppliedPayment.equals(ActualAppliedPayment))
						{
							System.out.println("Both Total values are same ");
							et.log(LogStatus.INFO,"Both Total values are same");
						}
						else
						{
							//System.out.println("lavanya");
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<------- Ordercheckout Both applied payments values are different ---------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Ordercheckout Both applied payments values are different ------>"+ErrorNumber);
							System.out.println("Actual Applied value is :"+ActualAppliedPayment);
							System.out.println("Expected Applied value is :"+ExpectedAppliedPayment);
						}
					}
				}	
			}
		}
		catch (Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}
	/**
	 * Calculates tax including third-party shipping and coupon discount.
	 */
	public static String TaxValueWithThirdPartyShippingIsTaxable(String OrderbaseThirdPartyShippingPrice, String TotalPrice, 
			String PriceAfterApplyingCoupon, String Tax, String OrderAmountValue) throws InterruptedException, AWTException
	{
		MouseAdjFunction();
		String OrderbaseThirdPartyShippingPrice2 = OrderbaseThirdPartyShippingPrice.substring(1, OrderbaseThirdPartyShippingPrice.length());
		double OrderbaseThirdPartyShippingPrice3 = Double.valueOf(OrderbaseThirdPartyShippingPrice2).doubleValue();

		//System.out.println("TotalPrice :"+TotalPrice);
		//System.out.println("PriceAfterApplyingCoupon :"+PriceAfterApplyingCoupon);

		double TotalPrice2 = Double.valueOf(TotalPrice).doubleValue();
		double CouponCode2 = Double.valueOf(PriceAfterApplyingCoupon).doubleValue();
		double Taxpercentage = Double.valueOf(Tax).doubleValue();
		double ExpectedTaxVlaue2 = ((TotalPrice2+OrderbaseThirdPartyShippingPrice3-CouponCode2)*Taxpercentage)/100;
		String ExpectedTaxVlaue3 = String.valueOf(ExpectedTaxVlaue2);
		String ExpectedTaxCalculationValue = Decimalsetting(ExpectedTaxVlaue3, OrderAmountValue);
		//System.out.println("ExpectedTaxCalculationValue :"+ExpectedTaxCalculationValue);
		String  ExpectedTaxCalculatedValue4 = Config.Currency+ExpectedTaxCalculationValue;
		return ExpectedTaxCalculatedValue4;

	}
	/**
	 * Adds the displayed tax amount to expected total.
	 */
	public static String VertexTotal(String ExpectedTotal)
	{
		String ActualOCTaxAmount = d.findElement(Property.OCTaxAmount).getText();
		String ActualOCTaxAmount1 = ActualOCTaxAmount.substring(1, ActualOCTaxAmount.length());
		double ActualOCTaxAmount2 = Double.valueOf(ActualOCTaxAmount1).doubleValue();
		String ExpectedTotalAmount1 = ExpectedTotal.substring(1,ExpectedTotal.length());
		double ExpectedTotalAmount2 = Double.valueOf(ExpectedTotalAmount1).doubleValue();

		double Total = ActualOCTaxAmount2 + ExpectedTotalAmount2;
		String TotalValue = "$"+Total;

		return TotalValue;
	}

	/**
	 * Calculates the grand total including subtotal, taxes, shipping, addons, coupons, and discounts.
	 */

	public static String OCgrandtotal(String Subtotal,String PromotionCoupon,String OC3rdpartytax1,
			String OC3rdpartyshippingprice1,String Addons,String DiscountcalculationfromSubTotal,
			String OrderAmountValue, String Total, String IsShippingTaxable, String Tax) throws InterruptedException, AWTException 
	{
		String  ExpectedSubTotal = Config.Currency+Subtotal;
		String ActualsubTotal1 = ExpectedSubTotal.substring(1,ExpectedSubTotal.length());
		double ActualsubTotal2 = Double.valueOf(ActualsubTotal1).doubleValue();
		//double OCactualsubtotal2 = Double.valueOf(OCactualsubtotal1).doubleValue();
		//System.out.println("oc subtotal is"+OCactualsubtotal2);
		String Expecteddiscount = Config.Currency + DiscountcalculationfromSubTotal;
		String ActualDiscountPercentage1 = Expecteddiscount.substring(1,Expecteddiscount.length());
		double ActualDiscountPercentage2 = Double.valueOf(ActualDiscountPercentage1).doubleValue();
		String Expectedcouponvalue = Config.Currency+PromotionCoupon;
		//System.out.println("expected promotion is " + Expectedcouponvalue);
		String Actualcouponvalue1 = Expectedcouponvalue.substring(1,Expectedcouponvalue.length());
		double Actualcouponvalue2 = Double.valueOf(Actualcouponvalue1).doubleValue();
		double OC3rdpartytax2 = Double.valueOf(OC3rdpartytax1).doubleValue();
		double OC3rdpartyshippingprice2 = Double.valueOf(OC3rdpartyshippingprice1).doubleValue();
		String ExpectedAddons = Config.Currency+Addons;
		String Actualaddons1 = ExpectedAddons.substring(1,ExpectedAddons.length()); 
		double ActualAddons2 = Double.valueOf(Actualaddons1).doubleValue();

		double OCgrandtotal = ActualsubTotal2 + OC3rdpartytax2 + OC3rdpartyshippingprice2+ActualAddons2 - (Actualcouponvalue2+ActualDiscountPercentage2);

		String OCgrandtotal1 = "" + OCgrandtotal;
		String OCgrandtotal4 = Decimalsetting(OCgrandtotal1,OrderAmountValue);
		OCgrandtotal1 = OCgrandtotal4;
		String OCgrandtotal2 = Config.Currency+OCgrandtotal1;
		//System.out.println("ocgrand total"+OCgrandtotal2);
		return OCgrandtotal2;

	}

	/**
	 * Calculates updated grand total including markup fee along with subtotal, taxes, shipping, addons, coupons, and discounts.
	 */

	public static String Ocupdatedmarkupgrandtotal(String Subtotal,String PromotionCoupon,String OCmarkupfee1,
			String OC3rdpartytax1,String OrderAmountValue,String Addons,String DiscountcalculationfromSubTotal,
			String OC3rdpartyshippingprice1) throws InterruptedException, AWTException
	{
		String  ExpectedSubTotal = Config.Currency+Subtotal;
		String ActualsubTotal1 = ExpectedSubTotal.substring(1,ExpectedSubTotal.length());
		double ActualsubTotal2 = Double.valueOf(ActualsubTotal1).doubleValue();
		String Expectedcouponvalue = Config.Currency+PromotionCoupon;
		String Actualcouponvalue1 = Expectedcouponvalue.substring(1,Expectedcouponvalue.length());
		double Actualcouponvalue2 = Double.valueOf(Actualcouponvalue1).doubleValue();
		double OCmarkupfee2 = Double.valueOf(OCmarkupfee1).doubleValue();
		double OC3rdpartytax2 = Double.valueOf(OC3rdpartytax1).doubleValue();
		String ExpectedAddons = Config.Currency+Addons;
		String Actualaddons1 = ExpectedAddons.substring(1,ExpectedAddons.length()); 
		double ActualAddons2 = Double.valueOf(Actualaddons1).doubleValue();
		String Expecteddiscount = Config.Currency + DiscountcalculationfromSubTotal;
		String ActualDiscountPercentage1 = Expecteddiscount.substring(1,Expecteddiscount.length());
		double ActualDiscountPercentage2 = Double.valueOf(ActualDiscountPercentage1).doubleValue();
		double OC3rdpartyshippingprice2 = Double.valueOf(OC3rdpartyshippingprice1).doubleValue();


		double OCupdatedtotal = ActualsubTotal2 + OC3rdpartytax2 + OCmarkupfee2 + OC3rdpartyshippingprice2 + ActualAddons2
				- (Actualcouponvalue2+ActualDiscountPercentage2);

		String OCupdatedtotal1 = "" + OCupdatedtotal;
		String OCupdatedtotal2 = Decimalsetting(OCupdatedtotal1,OrderAmountValue);
		OCupdatedtotal1 = OCupdatedtotal2;
		String OCupdatedtotal3 = Config.Currency+OCupdatedtotal1;
		//System.out.println("handiling fee grand ttal is"+OCupdatedtotal3);

		return OCupdatedtotal3;	
	}
	/**
	 * Calculates updated tax including subtotal, 3rd party shipping, addons, handling fee, discounts, and coupons.
	 */

	public static String UpdatedtaxOCMarkup(String Subtotal,String PromotionCoupon,String OrderAmountValue,String Addons,String DiscountcalculationfromSubTotal,
			String OC3rdpartyshippingprice1,String Tax,String userordershippingorhandlingfee) throws InterruptedException, AWTException 
	{
		String  ExpectedSubTotal = Config.Currency+Subtotal;
		String ActualsubTotal1 = ExpectedSubTotal.substring(1,ExpectedSubTotal.length());
		double ActualsubTotal2 = Double.valueOf(ActualsubTotal1).doubleValue();
		String Expectedcouponvalue = Config.Currency+PromotionCoupon;
		String Actualcouponvalue1 = Expectedcouponvalue.substring(1,Expectedcouponvalue.length());
		double Actualcouponvalue2 = Double.valueOf(Actualcouponvalue1).doubleValue();
		String ExpectedAddons = Config.Currency+Addons;
		String Actualaddons1 = ExpectedAddons.substring(1,ExpectedAddons.length()); 
		double ActualAddons2 = Double.valueOf(Actualaddons1).doubleValue();
		String Expecteddiscount = Config.Currency + DiscountcalculationfromSubTotal;
		String ActualDiscountPercentage1 = Expecteddiscount.substring(1,Expecteddiscount.length());
		double ActualDiscountPercentage2 = Double.valueOf(ActualDiscountPercentage1).doubleValue();
		double OC3rdpartyshippingprice2 = Double.valueOf(OC3rdpartyshippingprice1).doubleValue();
		String ExpectedTax = Tax+Config.PercentageSymbol;
		String ActualUpdatedTax = ExpectedTax.substring(0,ExpectedTax.length()-1);
		double ActualUpdatedTax2 =Double.valueOf(ActualUpdatedTax).doubleValue();
		String ExpectedHandilingFee = Config.Currency+userordershippingorhandlingfee;
		String ActualHandilingFee = ExpectedHandilingFee.substring(1,ExpectedHandilingFee.length());
		double ActualHandilingFee1 = Double.valueOf(ActualHandilingFee).doubleValue();
		double UpdatedTax = (((ActualsubTotal2+OC3rdpartyshippingprice2+ActualAddons2+ActualHandilingFee1) - 
				(ActualDiscountPercentage2+Actualcouponvalue2))* ActualUpdatedTax2)/100;

		String UpdatedTax1 = ""+UpdatedTax;
		String UpdatedTax2 = Decimalsetting(UpdatedTax1,OrderAmountValue);
		String UpdatedTax3 = Config.Currency+UpdatedTax2;
		///System.out.println("updated markup tax is" +UpdatedTax3);
		return UpdatedTax3;

	}

	/**
	 * Calculates updated tax excluding markup fees, considering subtotal, 3rd party shipping, addons, discounts, and coupons.
	 */

	public static String UpdatedtaxOCWithOutMarkup(String Subtotal,String PromotionCoupon,String OrderAmountValue,String Addons,String DiscountcalculationfromSubTotal,
			String OC3rdpartyshippingprice1,String Tax) throws InterruptedException, AWTException 
	{
		String  ExpectedSubTotal = Config.Currency+Subtotal;
		String ActualsubTotal1 = ExpectedSubTotal.substring(1,ExpectedSubTotal.length());
		double ActualsubTotal2 = Double.valueOf(ActualsubTotal1).doubleValue();
		String Expectedcouponvalue = Config.Currency+PromotionCoupon;
		String Actualcouponvalue1 = Expectedcouponvalue.substring(1,Expectedcouponvalue.length());
		double Actualcouponvalue2 = Double.valueOf(Actualcouponvalue1).doubleValue();
		String ExpectedAddons = Config.Currency+Addons;
		String Actualaddons1 = ExpectedAddons.substring(1,ExpectedAddons.length()); 
		double ActualAddons2 = Double.valueOf(Actualaddons1).doubleValue();
		String Expecteddiscount = Config.Currency + DiscountcalculationfromSubTotal;
		String ActualDiscountPercentage1 = Expecteddiscount.substring(1,Expecteddiscount.length());
		double ActualDiscountPercentage2 = Double.valueOf(ActualDiscountPercentage1).doubleValue();
		double OC3rdpartyshippingprice2 = Double.valueOf(OC3rdpartyshippingprice1).doubleValue();
		String ExpectedTax = Tax+Config.PercentageSymbol;
		String ActualUpdatedTax = ExpectedTax.substring(0,ExpectedTax.length()-1);
		double ActualUpdatedTax2 =Double.valueOf(ActualUpdatedTax).doubleValue();
		//String ExpectedHandilingFee = Config.Currency+userordershippingorhandlingfee;
		//String ActualHandilingFee = ExpectedHandilingFee.substring(0,ExpectedHandilingFee.length());
		// double ActualHandilingFee1 = Double.valueOf(ActualHandilingFee).doubleValue();
		double UpdatedTax = (((ActualsubTotal2+OC3rdpartyshippingprice2+ActualAddons2) - 
				(ActualDiscountPercentage2+Actualcouponvalue2))* ActualUpdatedTax2)/100;

		String UpdatedTax1 = ""+UpdatedTax;
		String UpdatedTax2 = Decimalsetting(UpdatedTax1,OrderAmountValue);
		String UpdatedTax3 = Config.Currency+UpdatedTax2;
		//System.out.println("updated tax " +UpdatedTax3);
		return UpdatedTax3;

	}

	/**
	 * Verifies the entire price structure on the Order Checkout page including:
	 * quantities, item prices, discounts, taxes, shipping, handling fees, promotions,
	 * and the final grand total.
	 *
	 * Handles complex conditions based on:
	 * - Shipping method and type (e.g., weight-based)
	 * - Promotions and discounts (flat/percentage)
	 * - Tax exemptions
	 * - Vertex or standard tax calculations
	 * - Third-party shipping/markup handling</p>
	 *
	 * If mismatches are found during any verification, the method logs an error,
	 * captures a screenshot, and increments the global error counter.
	 */
	public static void OrderCheckOutPriceInformantion(String Quantity, String ItemPerPrice, String Discount, String Addons,
			String TotalPrice, String Total, String PromotionCode, 
			String PromotionCoupon, String Tax, String PriceAfterCalculatingTax, 
			String AddonPricePerPiece, String DiscountPercentage, String PromotionDiscountAfterSubtractingFromSubTotal, 
			String PromotionDiscountPercentage, String DiscountCalculationFromSubTotal, String OrderType,
			String TestStep, String Parameters, String ProdutType, String OrderBase,
			String OrderBaseShipping, String CalculateTaxCondition, String EnablePromotionsORDiscounts,String Weighttype,
			String DiscountcalculationfromSubTotal,String Subtotal,String OrderAmountValue,
			String userordershippingorhandlingfee,String Priceafterapplyingfulfillmentshippingmarkupfee,String IsShippingTaxable,
			String PriceAfterApplyingCoupon,String IsTaxExempt,String DecimalValue,
			String SubTotal,String ShippingPricePerPiece) throws InterruptedException  {
		try{
			MouseAdjFunction();

			if(IsTaxExempt.equalsIgnoreCase("yes")){
				PriceAfterCalculatingTax="0.00";
				String PriceAfterCalculatingTax1 = Decimalsetting(PriceAfterCalculatingTax, DecimalValue);
				PriceAfterCalculatingTax = PriceAfterCalculatingTax1;
			}
			Common.Wait.wait5Second();
			// Apply promotion code if value more than zero
			if(PromotionCoupon.equals("0") || PromotionCoupon.equals("0.00") || PromotionCoupon.equals("0.000") || PromotionCoupon.equals("0.0000"))
			{
				//!System.out.println("Promotion coupon value is empty");
			}
			else
			{
				if(EnablePromotionsORDiscounts.equals("ON"))
				{
					if(Subtotal.equals("0") || Subtotal.equals("0.00") || Subtotal.equals("0.000") || Subtotal.equals("0.0000")){
						System.out.println("if subtotal is  zero promotion text field will be in hidden");

					}else{
						Common.Wait.wait2Second();
						d.findElement(Property.PromotionCodeTextBox).sendKeys(PromotionCode);
						Common.Wait.wait2Second();
						d.findElement(Property.PromotionApplyButton).click();
					}
				}
			}

			Common.Wait.wait2Second();

			String ExpectedOrderBaseShippingPrice = Config.Currency+ShippingPricePerPiece;;
			String ExpectedQuantity = Quantity;
			int ExpectedQuantity1 = Double.valueOf(ExpectedQuantity).intValue();
			String ExpectedShippingHandilingPrice = Config.Currency+Priceafterapplyingfulfillmentshippingmarkupfee;

			String ExpectedItemPrice = null;
			// Generate expected Item price based on condition		
			if(OrderType.equals("DynShipTOMultipleLocations") || OrderType.equals("StaticShipTOMultipleLocations") 
					|| OrderType.equals("StaticInventoryShipTOMultipleLocations"))
			{
				ExpectedItemPrice = "-";
			}
			else
			{
				ExpectedItemPrice = Config.Currency+ItemPerPrice;
			}

			String ExpectedAmount=null;


			ExpectedAmount = Config.Currency+TotalPrice;

			String ExpectedPromotionDiscount = Config.Currency+PromotionCoupon;
			// Generate promotion discount percentage based on promotion discount amount or percentage
			if(PromotionDiscountPercentage.equals("N"))
			{
				if(Subtotal.equals("0") || Subtotal.equals("0.00") || Subtotal.equals("0.000") || Subtotal.equals("0.0000")){
					System.out.println("if subtotal is  zero promotion value should be zero");
					if(!OrderType.equals("Download")) {
						ExpectedPromotionDiscount = "-"+Config.Currency+PromotionDiscountAfterSubtractingFromSubTotal;
					}
					else {
						ExpectedPromotionDiscount = Config.Currency+PromotionDiscountAfterSubtractingFromSubTotal;	
					}
				}
				else{
					ExpectedPromotionDiscount = Config.Currency+PromotionCoupon;
				}
			}
			else
			{
				ExpectedPromotionDiscount = Config.Currency+PromotionDiscountAfterSubtractingFromSubTotal;
			}

			String ExpectedTaxpercentage = "Tax ("+Tax+"%)";
			String ExpectedTaxAmount = Config.Currency+PriceAfterCalculatingTax;
			String ExpectedTotal = Config.Currency+Total;
			// Generate Applied discount based on Discount amount or percentage
			if(DiscountPercentage.equals("N"))
			{	
			}
			else
			{	
			}

			// Generate applied addon price based on perpice or flat rate
			if(AddonPricePerPiece.equals(""))
			{
			}
			else
			{
			}

			// Click on Quantity value to scroll up screen (Screen shot purpose).
			Common.Wait.wait5Second();
			//d.findElement(Property.OrderCheckoutGridQuantity).click();
			String ActualQuantity1=null;

			ActualQuantity1 = d.findElement(Property.OrderCheckoutGridQuantity).getText();


			int ActualQuantity = Double.valueOf(ActualQuantity1).intValue();
			// Compare the expected quantity with actual quantity
			if(ExpectedQuantity1 == ActualQuantity)
			{
				System.out.println("Order Checkout Both Quantity are same");
				et.log(LogStatus.INFO, "Order Checkout Both Quantity are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<------- Order Checkout Both Quantitys are different ------> "+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Checkout Both Quantitys are differentt------>"+ErrorNumber);
				System.out.println("Actual Quantity is : "+ActualQuantity);
				System.out.println("Expected Quantity is : "+ExpectedQuantity);
			}
			// Get the grid item price based on order process type condition
			String ActualOCItemPrice = null;

			ActualOCItemPrice = d.findElement(Property.OrderCheckoutGridItemPrice).getText();


			// Compare grid expected item price with actual item price 
			if(ExpectedItemPrice.equals(ActualOCItemPrice))
			{
				System.out.println("Order Checkout Both Item prices are same ");
				et.log(LogStatus.INFO, "Order Checkout Both Item Prices are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<------- Order Checkout Both Item prices are different --------> "+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Checkout Both Item prices are different------>"+ErrorNumber);
				System.out.println("ActualOCItemPrice : "+ActualOCItemPrice);
				System.out.println("ExpectedItemPrice : "+ExpectedItemPrice);
			}

			// Get the actual discount based on order process type
			if(EnablePromotionsORDiscounts.equals("ON"))
			{
				//code removed beacuase rght know we are not showing dicount in grid at checkout page
				//In further release we can add the code
			}
			if((OrderBase.equals("Order")||OrderBase.equals("Split Ship")) && (OrderType.equals("ShipToMyAddress")))
			{

				d.findElement(Property.OrderCheckOutGridOrderBaseAmounto).getText();
			}
			else
			{
				d.findElement(Property.OrderCheckoutGridAmounto).getText();
			}

			//Click on total amount attribute to scroll down screen (Screen shot purpose).
			//d.findElement(Property.OCTotal).click();

			//Get order check out subtotal value 
			ExpectedAmount = Config.Currency+TotalPrice;
			String ActualOCSubTotal = d.findElement(Property.OCSubTotal).getText();
			// Generate order checkout subtotal if we have third party shipping 

			// Verify the subtotal value in order check out page 
			// verify subtotal in general conditio (with out third party shipping)
			if((ExpectedAmount.equals(ActualOCSubTotal)))
			{
				System.out.println("Order Checkout Both SubTotals are same");
				et.log(LogStatus.INFO, "Order Checkout Both SUbTotals are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<----------- Order Checkout Both SubTotals are different ---------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Checkout Both SubTotals are different ------>"+ErrorNumber);
				System.out.println("ActualOCSubTotal : "+ActualOCSubTotal);
				System.out.println("ExpectedAmount : "+ExpectedAmount);
			}


			// Get the discount value if enable promotions or discounts on conditions
			if(EnablePromotionsORDiscounts.equals("ON"))
			{
				String ActualOCPromotionDiscount = d.findElement(Property.OCPromotionDiscount).getText();
				if(ActualOCPromotionDiscount.equals("-")) {

					if(DecimalValue.equals("2.0"))
					{
						ActualOCPromotionDiscount = "$0.00";
					}
					else if(DecimalValue.equals("3.0"))
					{
						ActualOCPromotionDiscount = "$0.000";
					}
					else if(DecimalValue.equals("4.0"))
					{
						ActualOCPromotionDiscount = "$0.0000";
					}
					else if(DecimalValue.equals("0.0"))
					{
						ActualOCPromotionDiscount = "$0";
					}
				}

				// Compare expected promotion with actual promotion
				if(ExpectedPromotionDiscount.equals(ActualOCPromotionDiscount))
				{
					System.out.println("Order Checkout Both Promotion discounts are same");
					et.log(LogStatus.INFO, "Order Checkout Both Promotion discounts are same");
				}
				else
				{

					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Order Checkout Both Promotion discounts are different ---------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Order Checkout Both Promotion discounts are different ------>"+ErrorNumber);
					System.out.println("ActualOCPromotionDiscount : "+ActualOCPromotionDiscount);
					System.out.println("ExpectedPromotionDiscount : "+ExpectedPromotionDiscount);
				}
			}

			// This shipping price is available  with only order base 
			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download"))
			{
				String ActualOCShippingAmount = d.findElement(Property.OCTotalShippingPrice).getText();
				if(!(Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.00")||Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.000")
						||Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.0000") || Priceafterapplyingfulfillmentshippingmarkupfee.equals("0")))
				{//condition with shipping handling fee and mark up fee
					if(ExpectedShippingHandilingPrice.equals(ActualOCShippingAmount))
					{
						System.out.println("Order Checkout Both Shipping Amounts are same ");
						et.log(LogStatus.INFO, "Order Checkout Both Shipping Amounts are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<------------ Order Checkout Both Shipping Amounts are different -------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Order Checkout Both Shipping Amounts are different ------>"+ErrorNumber);
						System.out.println("ActualOCShippingAmount : "+ActualOCShippingAmount);
						System.out.println("ExpectedShippingAmount : "+ExpectedShippingHandilingPrice);
					}
				}// ends
				else if(((Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.00")||Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.000")
						||Priceafterapplyingfulfillmentshippingmarkupfee.equals("0.0000") ||
						Priceafterapplyingfulfillmentshippingmarkupfee.equals("0"))) && !(Weighttype.equals("KGS")|| Weighttype.equals("LBS")))
				{// condtion with out shipping handling fee and mark up fee
					if(ExpectedOrderBaseShippingPrice.equals(ActualOCShippingAmount))
					{
						System.out.println("Order Checkout Both Shipping Amounts are same ");
						et.log(LogStatus.INFO, "Order Checkout Both Shipping Amounts are same");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<------------ Order Checkout Both Shipping Amounts are different ---------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Order Checkout Both Shipping Amounts are different ------>"+ErrorNumber);
						System.out.println("ActualOCShippingAmount : "+ActualOCShippingAmount);
						System.out.println("ExpectedShippingAmount : "+ExpectedOrderBaseShippingPrice);
					}
				}// condition with out shipping handling fee and mark fee ends
			}// order base ends


			String[] CalculateTaxConditions = CalculateTaxCondition.split("_");
			if(IsTaxExempt.equalsIgnoreCase("yes")){}else{

				if(!CalculateTaxConditions[0].equals("---Select---") && CalculateTaxConditions[1].equals("ON"))
				{	
					String ActualOCTaxPercentage=null;
					if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
						ActualOCTaxPercentage = d.findElement(Property.OCTaxPercentage).getText();
					}
					else {
						ActualOCTaxPercentage = d.findElement(Property.OCTaxPercentageB1).getText();
					}

					//
					if(CalculateTaxConditions[0].equals("Vertex"))
					{
						if(ActualOCTaxPercentage.equals(""))
						{
							System.out.println("Order Checkout Both Tax percentages are saeme ");
							et.log(LogStatus.INFO, "Order Checkout Both Tax percentages are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<------------- Order Checkout Both Tax percentages are different ------------>"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Order Checkout Both Tax percentages are different ------>"+ErrorNumber);
							System.out.println("ActualOCTaxPercentage : "+ActualOCTaxPercentage);
							System.out.println("ExpectedTaxpercentage : "+"");
						}
					}
					else
					{
						//Due to 3rd party shipping methods tax changed according to the shipping method

						if(ExpectedTaxpercentage.equals(ActualOCTaxPercentage))
						{
							System.out.println("Both Tax percentages are saeme ");
							et.log(LogStatus.INFO, "Order Checkout Both Tax percentages are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<------------- Order Checkout Both Tax percentages are different ------------>"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Order Checkout Both Tax percentages are different ------>"+ErrorNumber);
							System.out.println("ActualOCTaxPercentage : "+ActualOCTaxPercentage);
							System.out.println("ExpectedTaxpercentage : "+ExpectedTaxpercentage);
						}
					}
				}
				if(CalculateTaxConditions[0].equals("Vertex"))
				{
					String ActualOCTaxAmount;

					if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
						ActualOCTaxAmount = d.findElement(Property.OCTaxAmount).getText();
					}
					else {
						ActualOCTaxAmount = d.findElement(Property.OCTaxAmountB1).getText();
					}
					if(ActualOCTaxAmount.matches("\\$\\d{1,3}\\.\\d{2,4}"))
					{
						System.out.println("Order Checkout Both Tax amounts are same");
						et.log(LogStatus.INFO, "Order Checkout Both Tax amounts are same");
						//!System.out.println("ActualOCTaxAmount : "+ActualOCTaxAmount);
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Order Checkout Both Tax amounts are different -------------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Order Checkout Both Tax amounts are different ------>"+ErrorNumber);
						System.out.println("ActualOCTaxAmount : "+ActualOCTaxAmount);
						System.out.println("ExpectedTaxAmount : "+"\\$\\d{1,3}\\.\\d{2,4}");
					}
				}
				else
				{
					String ActualOCTaxAmount;

					if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
						ActualOCTaxAmount = d.findElement(Property.OCTaxAmount).getText();
					}
					else {
						ActualOCTaxAmount = d.findElement(Property.OCTaxAmountB1).getText();
					}
					//System.out.println("expected tax is" +ExpectedTaxAmount);
					if (OrderBase.equals("Order")&& (Weighttype.equals("KGS")|| Weighttype.equals("LBS")) &&
							IsShippingTaxable.equals("Yes"))
					{
						String OrderbaseThirdPartyShippingPrice = d.findElement(Property.OCTotalShippingPrice).getText();

						String ExpectedTaxCalculatedValue4 = TaxValueWithThirdPartyShippingIsTaxable(OrderbaseThirdPartyShippingPrice, TotalPrice, PriceAfterApplyingCoupon,
								Tax, OrderAmountValue);
						if(ActualOCTaxAmount.equals(ExpectedTaxCalculatedValue4))
						{
							System.out.println("Order Checkout Both Tax Amount are Same ");
							et.log(LogStatus.INFO, "Order Checkout Both Tax amounts are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<------- Order Checkout Both Tax amounts are different ------------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Order Checkout Both Tax amounts are different ------>"+ErrorNumber);
							System.out.println("ActualOCTaxAmount : "+ActualOCTaxAmount);
							System.out.println("ExpectedTaxAmount : "+ExpectedTaxCalculatedValue4);
						}
					}
					else 
					{
						if((ExpectedTaxAmount.equals(ActualOCTaxAmount)))
						{
							//!System.out.println("Both Tax amounts are same which was changed by shipping methods");
							et.log(LogStatus.INFO, "Order Checkout Both Tax amounts are same");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<---------- Order Checkout Both Tax amounts are different ------------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Order Checkout Both Tax amounts are different ------>"+ErrorNumber);
							System.out.println("ActualOCTaxAmount : "+ActualOCTaxAmount);
							System.out.println("ExpectedTaxAmount : "+ExpectedTaxAmount);
						}
					}
				}
			}
			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
				String Expectedhandilingfee= Config.Currency+userordershippingorhandlingfee;
				String OChandilingfee = d.findElement(Property.OCHandilingfee).getText();

				if(userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
						userordershippingorhandlingfee.equals("0.0000")) 
				{ 	
				}
				else
				{
					if(OChandilingfee.equals(Expectedhandilingfee))
					{
						System.out.println("Order Checkout Both handiling fee are same ");
						et.log(LogStatus.INFO, "Order Checkout Both handiling fee are same");
					}
					else
					{
						//System.out.println("666666");
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Order Checkout Both handiling fee are different ----------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Order Checkout Both handiling fee are different ------>"+ErrorNumber);
						System.out.println("Actuahandilingfee : "+OChandilingfee);
						System.out.println("Expectedhandlingfee : "+Expectedhandilingfee);
					}
				}
			}

			String ActualOCTotal;
			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
				ActualOCTotal = d.findElement(Property.OCTotal).getText();
			}
			else {
				ActualOCTotal = d.findElement(Property.OCTotalB1).getText();
			}

			if(ExpectedTotal.equals(ActualOCTotal))
			{
				System.out.println("Order Checkout Both GrandTotal Amounts are same ");
				et.log(LogStatus.INFO, "Order Checkout Both GrandTotal Amounts are same");
			}
			else
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Order Checkout Both GrandTotal Amounts are different ----------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Checkout Both GrandTotal Amounts are different ------>"+ErrorNumber);
				System.out.println("ActualOCTotal : "+ActualOCTotal);
				System.out.println("ExpectedTotal : "+ExpectedTotal);
			}
		}
		catch (Exception e)
		{
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}		
	}

	/**
	 * Verifies order summary details (subtotal, shipping, handling fee, tax, total)
	 * on the page against expected values, logging mismatches and capturing screenshots on errors.
	 * Adjusts verification based on the order type (e.g., "Mailinglist" uses different locators).
	 */
	public static void OrderSummaryVerification(String SubTotal,String TotalPrice,String Total, String ShippingPricePerPiece,
			String Tax,String PriceAfterCalculatingTax, String userordershippingorhandlingfee, String OrderType) throws InterruptedException {
		try {	
			String ActualSubtotalOs;
			String ActualShippingOs = " ";
			String ActualHandlingFeeOs = " ";
			String ActualTaxPercentageOs;
			String ActualTaxAmountOs;
			String ActualTotalOs;

			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download"))
			{
				ActualSubtotalOs = d.findElement(Property.SubtotalOs).getText();
				ActualShippingOs = d.findElement(Property.ShippingOs).getText();
				ActualHandlingFeeOs = d.findElement(Property.HandlingFeeOs).getText();
				ActualTaxPercentageOs = d.findElement(Property.TaxPercentageOs).getText();
				ActualTaxAmountOs = d.findElement(Property.TaxAmountOs).getText();
				ActualTotalOs = d.findElement(Property.TotalOs).getText();
			}
			else {
				ActualSubtotalOs = d.findElement(Property.SubtotalOsB1).getText();
				ActualTaxPercentageOs = d.findElement(Property.TaxPercentageOsB1).getText();
				ActualTaxAmountOs = d.findElement(Property.TaxAmountOsB1).getText();
				ActualTotalOs = d.findElement(Property.TotalOsB1).getText();
			}

			String ExpectedSubtotalOs = Config.Currency+TotalPrice;
			String ExpectedShippingOs = Config.Currency+ShippingPricePerPiece;
			String ExpectedHandlingFeeOs = Config.Currency+userordershippingorhandlingfee;
			String ExpectedTaxpercentageOs = "Tax ("+Tax+"%)";
			String ExpectedTaxAmountOs = Config.Currency+PriceAfterCalculatingTax;
			String ExpectedTotalOs = Config.Currency+Total;

			if(ActualSubtotalOs.equals(ExpectedSubtotalOs)) 
			{
				System.out.println("Both SubTotal Amounts are same in order Summary Page");
				et.log(LogStatus.INFO, "Order Summary Both SubTotal Amounts are same");
			}
			else 
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Order Summary Page Both SubTotal Amounts are different ----------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Summary Page SubTotal Amounts are different ------>"+ErrorNumber);
				System.out.println("ActualSubtotalOs : "+ActualSubtotalOs);
				System.out.println("ExpectedSubtotalOs : "+ExpectedSubtotalOs);
			}
			if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
				if(ActualShippingOs.equals(ExpectedShippingOs)) 
				{
					System.out.println("Both Shipping price are same in order Summary Page");
					et.log(LogStatus.INFO, "Order Summary Both Shipping price are same");
				}
				else 
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Order Summary Page Both Shipping Amounts are different ----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Order Summary Page Shipping Amounts are different ------>"+ErrorNumber);
					System.out.println("ActualShippingOs : "+ActualShippingOs);
					System.out.println("ExpectedShippingOs : "+ExpectedShippingOs);
				}

				if(ActualHandlingFeeOs.equals(ExpectedHandlingFeeOs)) 
				{
					System.out.println("Both HandlingFee Amounts are same in order Summary Page");
					et.log(LogStatus.INFO, "Order Summary Both Handling Amount are same");

				}
				else 
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Order Summary Page Both HandlingFee Amounts are different ----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Order Summary Page HandlingFee Amounts are different ------>"+ErrorNumber);
					System.out.println("ActualHandlingFeeOs : "+ActualHandlingFeeOs);
					System.out.println("ExpectedHandlingFeeOs : "+ExpectedHandlingFeeOs);
				}
			}
			if(ActualTaxPercentageOs.equals(ExpectedTaxpercentageOs)) 
			{
				System.out.println("Both Tax Percentages are same in order Summary Page");
				et.log(LogStatus.INFO, "Order Summary Both Tax Percentages are same");
			}
			else 
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Order Summary Page Both Tax Percentage Amounts are different ----------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Summary Page Tax Percentage Amounts are different ------>"+ErrorNumber);
				System.out.println("ActualTaxPercentageOs : "+ActualTaxPercentageOs);
				System.out.println("ExpectedTaxpercentageOs : "+ExpectedTaxpercentageOs);
			}

			if(ActualTaxAmountOs.equals(ExpectedTaxAmountOs)) 
			{
				System.out.println("Both Tax Amounts are same in order Summary Page");
				et.log(LogStatus.INFO, "Order Summary Both Tax Amounts are same");
			}
			else 
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Order Summary Page Both Tax Amounts are different ----------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Summary Page Tax Amounts are different ------>"+ErrorNumber);
				System.out.println("ActualTaxAmountOs : "+ActualTaxAmountOs);
				System.out.println("ExpectedTaxAmountOs : "+ExpectedTaxAmountOs);
			}

			if(ActualTotalOs.equals(ExpectedTotalOs)) 
			{
				System.out.println("Both Total Amounts are same in order Summary Page");
				et.log(LogStatus.INFO, "Order Summary Both Total Amounts are same");
			}
			else 
			{
				ErrorNumber = ErrorNumber+1;
				captureScreenshot();
				System.out.println("<--------- Order Summary Page Both Total Amounts are different ----------->"+ErrorNumber);
				et.log(LogStatus.ERROR,"<---- Order Summary Page Total Amounts are different ------>"+ErrorNumber);
				System.out.println("ActualTotalOs : "+ActualTotalOs);
				System.out.println("ExpectedTotalOs : "+ExpectedTotalOs);
			}
		}
		catch(Exception e){
			ErrorNumber = ErrorNumber+1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Verifies pricing details in the "View Orders" (Order Info) page for different user roles or stores (User, Approver, PS).
	 * Compares actual values shown on the UI with expected values calculated from the backend/order context.
	 * Handles the following verifications:
	 *     Subtotal price verification
	 *     Add-on charges verification
	 *     Discount price verification (if promotions/discounts are enabled)
	 *     Total amount (parameter present but not yet validated)
	 *     Shipping price per piece (parameter present but not yet validated)
	 *     Tax amount (parameter present but not yet validated)
	 *     Final price after tax (parameter present but not yet validated)
	 *     Handling or shipping fees (parameter present but not yet validated)
	 *     Postage cost (parameter present but not yet validated)
	 *     Promotion discount (parameters present but not yet validated)
	 */
	public static void gen5viewordersOrderinfoVerification(String SubTotal,String Total,String Addons,String ShippingPricePerPiece,
			String Tax,String PriceAfterCalculatingTax, String userordershippingorhandlingfee, String Postage,
			String DiscountPercentage,String Discount,String DiscountCalculationFromSubTotal,String EnablePromotionsORDiscounts,String PromotionDiscountPercentage,
			String PromotionCoupon,String PromotionDiscountAfterSubtractingFromSubTotal, String PageName, String PriceType, String OrderType, String Store)throws InterruptedException 
	{

		try {
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			Common.Wait.wait2Second();

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);
			Common.Wait.wait2Second();			

			// Sub total price verification
			String ActualVOGSubtotal ;
			String ExpectedVOGSubTotal =null;
			ExpectedVOGSubTotal = Config.Currency + SubTotal;
			PriceType  = "Subtotal";

			if(Store.equalsIgnoreCase("user")) {
				WebElement Actual=	d.findElement(Property.OrderInfoSubTotal);
				ActualVOGSubtotal=	Actual.getText();
				if(ActualVOGSubtotal.equals(ExpectedVOGSubTotal))
				{
					System.out.println("Sub total prices are same In view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page Sub total prices are same");
					RW_File.WriteResult(ActualVOGSubtotal, ExpectedVOGSubTotal,PageName, PriceType, "PASS");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- subtotal prices are different In view orders(OrderInfo) page ------> "+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- subtotal prices are different In view orders(OrderInfo) page------>"+ErrorNumber);
					System.out.println("Expected value: "+ExpectedVOGSubTotal);
					System.out.println("Actual value :"+ActualVOGSubtotal);
					RW_File.WriteResult(ActualVOGSubtotal, ExpectedVOGSubTotal,PageName, PriceType, "FAIL");
				}
			}
			else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
				ActualVOGSubtotal=	d.findElement(Property.APOrderInfoSubTotal).getText();

				if(ActualVOGSubtotal.equals(ExpectedVOGSubTotal))
				{
					if(Store.equals("App")) {
						System.out.println("Sub total prices are same In Approver view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Sub total prices are same in Approver");
					}
					else if(Store.equals("PS")) {
						System.out.println("Sub total prices are same In PS view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Sub total prices are same in PS");
					}
				}
				else
				{
					if(Store.equals("App")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- subtotal prices are different In Approver view orders(OrderInfo) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- subtotal prices are different In Approver view orders(OrderInfo) page------>"+ErrorNumber);
						System.out.println("Expected value: "+ExpectedVOGSubTotal);
						System.out.println("Actual value :"+ActualVOGSubtotal);
					}
					else if(Store.equals("PS")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- subtotal prices are different In PS view orders(OrderInfo) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- subtotal prices are different In PS view orders(OrderInfo) page------>"+ErrorNumber);
						System.out.println("Expected value: "+ExpectedVOGSubTotal);
						System.out.println("Actual value :"+ActualVOGSubtotal);	
					}
				}
			}

			//Addon price verification
			PriceType  = "AddOns";
			String ExpectedAddonPrice = Config.Currency+Addons;
			String ActuaAddonPrice;

			if(Store.equalsIgnoreCase("user")) {
				ActuaAddonPrice = d.findElement(Property.OrderInfoAddOnPrice).getText();
				if(ActuaAddonPrice.equals(ExpectedAddonPrice))
				{
					System.out.println("Addon prices are same In view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page Addon prices are same");
					RW_File.WriteResult(ActuaAddonPrice, ExpectedAddonPrice,PageName, PriceType, "PASS");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------  Addon prices are different In view orders(OrderInfo) page--------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<----  Addon prices are different In view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("Expected Addon price In view orders(OrderInfo) page : "+ExpectedAddonPrice);
					System.out.println("Actual Addon price In view orders(OrderInfo) page: "+ActuaAddonPrice);
					RW_File.WriteResult(ActuaAddonPrice, ExpectedAddonPrice,PageName, PriceType, "FAIL");
				}
			}
			else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
				ActuaAddonPrice = d.findElement(Property.APOrderInfoAddOnPrice).getText();
				if(ActuaAddonPrice.equals(ExpectedAddonPrice))
				{
					if(Store.equals("App")) {
						System.out.println("Addon prices are same In Approver view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Addon prices are samein Approver");
					}
					else if(Store.equals("PS")) {
						System.out.println("Addon prices are same In PS view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Addon prices are same in PS");
					}
				}
				else
				{
					if(Store.equals("App")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------  Addon prices are different In Approver view orders(OrderInfo) page--------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Addon prices are different In Approver view orders(OrderInfo) page ------>"+ErrorNumber);
						System.out.println("Expected Addon price In Approver view orders(OrderInfo) page : "+ExpectedAddonPrice);
						System.out.println("Actual Addon price In Approver view orders(OrderInfo) page: "+ActuaAddonPrice);
					}
					else if(Store.equals("PS")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------  Addon prices are different In PS view orders(OrderInfo) page--------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Addon prices are different In PS view orders(OrderInfo) page ------>"+ErrorNumber);
						System.out.println("Expected Addon price In PS view orders(OrderInfo) page : "+ExpectedAddonPrice);
						System.out.println("Actual Addon price In PS view orders(OrderInfo) page: "+ActuaAddonPrice);
					}
				}
			}

			//Discount price verification
			String ExpectedOrderDiscount = null;
			if(DiscountPercentage.equals("N"))
			{
				ExpectedOrderDiscount = "-"+Config.Currency+Discount;
			}
			else if(Discount.equals("0")||Discount.equals("0.0")||Discount.equals("0.00")||
					Discount.equals("0.000")||Discount.equals("0.0000"))
			{
				ExpectedOrderDiscount = "-"+Config.Currency+DiscountCalculationFromSubTotal;
			}
			else
			{
				ExpectedOrderDiscount = "-"+Config.Currency+DiscountCalculationFromSubTotal;	
			}

			if(EnablePromotionsORDiscounts.equals("ON"))
			{  
				String ActualOrderDiscount="";
				PriceType  = "Discount";
				if(Store.equalsIgnoreCase("user")) {
					ActualOrderDiscount = d.findElement(Property.OrderInfoDiscount).getText();  
					if(ActualOrderDiscount.equals(ExpectedOrderDiscount))
					{
						System.out.println(" Discount prices are same In view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Discount prices are same");
						RW_File.WriteResult(ActualOrderDiscount, ExpectedOrderDiscount,PageName, PriceType, "PASS");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<-------- Discount prices are different In view orders(OrderInfo) page --------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Discount prices are different In view orders(OrderInfo) page------>"+ErrorNumber);
						System.out.println("Expected Discount price In view orders(OrderInfo) page : "+ExpectedOrderDiscount);
						System.out.println("Actual Discount price In view orders(OrderInfo) page : "+ActualOrderDiscount);
						RW_File.WriteResult(ActualOrderDiscount, ExpectedOrderDiscount,PageName, PriceType, "FAIL");
					}
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualOrderDiscount = d.findElement(Property.APOrderInfoDiscount).getText();  
					if(ActualOrderDiscount.equals(ExpectedOrderDiscount))
					{	
						if(Store.equals("App")) {
							System.out.println(" Discount prices are same In Approver view orders(OrderInfo) page");
							et.log(LogStatus.INFO, "Order Deatils Page Discount prices are same in Approver");
						}
						else if(Store.equals("PS")) {
							System.out.println(" Discount prices are same In PS view orders(OrderInfo) page");
							et.log(LogStatus.INFO, "Order Deatils Page Discount prices are samein PS");
						}
					}
					else
					{
						if(Store.equals("App")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<-------- Discount prices are different In Approver view orders(OrderInfo) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<----  Discount prices are different In Approver view orders(OrderInfo) page------>"+ErrorNumber);
							System.out.println("Expected Discount price In Approver view orders(OrderInfo) page : "+ExpectedOrderDiscount);
							System.out.println("Actual Discount price In Approver view orders(OrderInfo) page : "+ActualOrderDiscount);
						}
						else if(Store.equals("PS")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<-------- Discount prices are different In PS view orders(OrderInfo) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<----  Discount prices are different In PS view orders(OrderInfo) page------>"+ErrorNumber);
							System.out.println("Expected Discount price In PS view orders(OrderInfo) page : "+ExpectedOrderDiscount);
							System.out.println("Actual Discount price In PS view orders(OrderInfo) page : "+ActualOrderDiscount);
						}
					}
				}
				//Promotion price verification
				String ExpectedOrderPromotonDiscount = null;
				PriceType  = "Promotion Discount";
				if(PromotionDiscountPercentage.equals("N"))
				{
					ExpectedOrderPromotonDiscount = "-"+Config.Currency+PromotionCoupon;
				}
				else
				{
					ExpectedOrderPromotonDiscount = "-"+Config.Currency+PromotionDiscountAfterSubtractingFromSubTotal;
				}

				String ActualOrderPromotonDiscount = "";
				if(PromotionCoupon.equals("0") || PromotionCoupon.equals("0.00") || PromotionCoupon.equals("0.000") || PromotionCoupon.equals("0.0000"))
				{

				}
				else
				{
					if(EnablePromotionsORDiscounts.equals("ON"))
						if(SubTotal.equals("0") || SubTotal.equals("0.00") || SubTotal.equals("0.000") || SubTotal.equals("0.0000")){
							System.out.println("if subtotal is  zero promotion is not shown in  application in UserViewOrdersPage");
						}
						else{
							if(Store.equalsIgnoreCase("user")) {
								ActualOrderPromotonDiscount = d.findElement(Property.OrderInfoPromotionDiscount).getText();
							}
							else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
								ActualOrderPromotonDiscount = d.findElement(Property.APOrderInfoPromotionDiscount).getText();
							}
						}
				}

				if(PromotionCoupon.equals("0") || PromotionCoupon.equals("0.00") || PromotionCoupon.equals("0.000") || PromotionCoupon.equals("0.0000"))
				{ }
				else
				{
					if(SubTotal.equals("0") || SubTotal.equals("0.00") || SubTotal.equals("0.000") || SubTotal.equals("0.0000"))
					{
						System.out.println("if subtotal is  zero promotion value should be zero and not visible in View Orders ");
					}
					else{

						if(ActualOrderPromotonDiscount.equals(ExpectedOrderPromotonDiscount))
						{
							if(Store.equals("user")) {
								System.out.println("Promotion discount values are same In view orders(OrderInfo) page");
								et.log(LogStatus.INFO, "Order Deatils Page Promotion discount values are same");
								RW_File.WriteResult(ActualOrderPromotonDiscount, ExpectedOrderPromotonDiscount,PageName, PriceType, "PASS");
							}
							else if(Store.equals("App")) {
								System.out.println("Promotion discount values are same In Approver view orders(OrderInfo) page");
								et.log(LogStatus.INFO, "Order Deatils Page Promotion discount values are same in Approver");
							}
							else if(Store.equals("PS")) {
								System.out.println("Promotion discount values are same In PS view orders(OrderInfo) page");
								et.log(LogStatus.INFO, "Order Deatils Page Promotion discount values are same in PS");
							}

						}
						else
						{
							if(Store.equals("user")) {
								ErrorNumber = ErrorNumber+1;
								captureScreenshot();
								System.out.println("<-------- Promotion Discount prices are different In view orders(OrderInfo) page --------->"+ErrorNumber);
								et.log(LogStatus.ERROR,"<---- Promotion Discount prices are different In view orders(OrderInfo) page------>"+ErrorNumber);
								System.out.println("Expected Promotion Discount price : "+ExpectedOrderPromotonDiscount);
								System.out.println("Actual Promotion Discount price : "+ActualOrderPromotonDiscount);
								RW_File.WriteResult(ActualOrderPromotonDiscount, ExpectedOrderPromotonDiscount,PageName, PriceType, "FAIL");
							}
							else if(Store.equals("App")) {
								ErrorNumber = ErrorNumber+1;
								captureScreenshot();
								System.out.println("<-------- Promotion Discount prices are different In Approver view orders(OrderInfo) page --------->"+ErrorNumber);
								et.log(LogStatus.ERROR,"<---- Promotion Discount prices are different In Approver view orders(OrderInfo) page------>"+ErrorNumber);
								System.out.println("Expected Promotion Discount price : "+ExpectedOrderPromotonDiscount);
								System.out.println("Actual Promotion Discount price : "+ActualOrderPromotonDiscount);	
							}
							else if(Store.equals("PS")) {
								ErrorNumber = ErrorNumber+1;
								captureScreenshot();
								System.out.println("<-------- Promotion Discount prices are different In PS view orders(OrderInfo) page --------->"+ErrorNumber);
								et.log(LogStatus.ERROR,"<---- Promotion Discount prices are different In PS view orders(OrderInfo) page------>"+ErrorNumber);
								System.out.println("Expected Promotion Discount price : "+ExpectedOrderPromotonDiscount);
								System.out.println("Actual Promotion Discount price : "+ActualOrderPromotonDiscount);
							}
						}
					}
				}
			}

			//shipping price verification
			String ActualOrderShippingPrice= "";
			if(Store.equalsIgnoreCase("user")) {
				ActualOrderShippingPrice = d.findElement(Property.OrderInfoShipping).getText();
			}
			else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
				ActualOrderShippingPrice = d.findElement(Property.APOrderInfoShipping).getText();
			}
			String ExpectedOrderShippingPrice =null;

			ExpectedOrderShippingPrice = Config.Currency + ShippingPricePerPiece;
			PriceType  = "Shipping";


			if(ActualOrderShippingPrice.equals(ExpectedOrderShippingPrice))
			{
				if(Store.equals("user")) {
					System.out.println(" Shipping prices are same In view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page Shipping prices are same");
					RW_File.WriteResult(ActualOrderShippingPrice, ExpectedOrderShippingPrice,PageName, PriceType, "PASS");
				}
				else if(Store.equals("App")) {
					System.out.println(" Shipping prices are same In Approver view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page Shipping prices are same in Approver");
				}
				else if(Store.equals("PS")) {
					System.out.println(" Shipping prices are same In PS view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page Shipping prices are same in PS");
				}
			}
			else
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In view orders(OrderInfo) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Shipping prices are different In view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(OrderInfo) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(OrderInfo) page: "+ActualOrderShippingPrice);
					RW_File.WriteResult(ActualOrderShippingPrice, ExpectedOrderShippingPrice,PageName, PriceType, "FAIL");
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In Approver view orders(OrderInfo) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Shipping prices are different In Approver view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(OrderInfo) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(OrderInfo) page: "+ActualOrderShippingPrice);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In PS view orders(OrderInfo) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Shipping prices are different In PS view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(OrderInfo) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(OrderInfo) page: "+ActualOrderShippingPrice);
				}
			}

			//Handling fee
			String ActualHandlingFee="";
			if(Store.equalsIgnoreCase("user")) {
				ActualHandlingFee = d.findElement(Property.OrderInfoShippingHandling).getText();
			}
			else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
				ActualHandlingFee = d.findElement(Property.APOrderInfoShippingHandling).getText();
			}

			String ExpectedHandlingFee =null;

			ExpectedHandlingFee = Config.Currency + userordershippingorhandlingfee;
			PriceType  = "Handling Fee";

			if(ActualHandlingFee.equals(ExpectedHandlingFee)) 
			{
				if(Store.equals("user")) {
					System.out.println("HandlingFee Amounts are same in order Summary Page");
					et.log(LogStatus.INFO, "Order Deatils Page HandlingFee Amounts are same");
					RW_File.WriteResult(ActualHandlingFee, ExpectedHandlingFee,PageName, PriceType, "PASS");
				}
				else if(Store.equals("App")) {
					System.out.println("HandlingFee Amounts are same in Approver order Summary Page");
					et.log(LogStatus.INFO, "Order Deatils Page HandlingFee Amounts are same in Approver");
				}
				else if(Store.equals("PS")) {
					System.out.println("HandlingFee Amounts are same in PS order Summary Page");
					et.log(LogStatus.INFO, "Order Deatils Page HandlingFee Amounts are same in PS");
				}
			}
			else 
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  HandlingFee Amounts are different In view orders(OrderInfo) page  ----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- HandlingFee Amounts are different In view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("ActualHandlingFeeOs : "+ActualHandlingFee);
					System.out.println("ExpectedHandlingFeeOs : "+ExpectedHandlingFee);
					RW_File.WriteResult(ActualHandlingFee, ExpectedHandlingFee,PageName, PriceType, "Fail");
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  HandlingFee Amounts are different In Approver view orders(OrderInfo) page  ----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- HandlingFee Amounts are different In Approver view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("ActualHandlingFeeOs : "+ActualHandlingFee);
					System.out.println("ExpectedHandlingFeeOs : "+ExpectedHandlingFee);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  HandlingFee Amounts are different In PS view orders(OrderInfo) page  ----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- HandlingFee Amounts are different In PS view orders(OrderInfo) page  ------>"+ErrorNumber);
					System.out.println("ActualHandlingFeeOs : "+ActualHandlingFee);
					System.out.println("ExpectedHandlingFeeOs : "+ExpectedHandlingFee);
				}
			}

			//Postage
			if(OrderType.equals("Mailinglist")) {
				String ActualPostageAmount="";
				if(Store.equalsIgnoreCase("user")) {
					ActualPostageAmount = d.findElement(Property.OrderInfoPostage).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualPostageAmount = d.findElement(Property.APOrderInfoPostage).getText();
				}

				String ExpectedPostageAmount =null;

				ExpectedPostageAmount = Config.Currency + Postage;
				PriceType  = "Postage";
				if(ActualPostageAmount.equals(ExpectedPostageAmount)) 
				{
					if(Store.equals("user")) {
						System.out.println("Postage Amounts are same In view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Postage Amounts are same");
						RW_File.WriteResult(ActualPostageAmount, ExpectedPostageAmount,PageName, PriceType, "PASS");
					}
					else if(Store.equals("App")) {
						System.out.println("Postage Amounts are same In Approver view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Postage Amounts are same in Approver");
					}
					else if(Store.equals("PS")) {
						System.out.println("Postage Amounts are same In PS view orders(OrderInfo) page");
						et.log(LogStatus.INFO, "Order Deatils Page Postage Amounts are same in PS");
					}
				}
				else 
				{
					if(Store.equals("user")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<---------  Postage Amounts are different In view orders(OrderInfo) page----------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----Postage Amounts are different In view orders(OrderInfo) page ------>"+ErrorNumber);
						System.out.println("ActualPostageAmount : "+ActualPostageAmount);
						System.out.println("ExpectedPostageAmount : "+ExpectedPostageAmount);
						RW_File.WriteResult(ActualPostageAmount, ExpectedPostageAmount,PageName, PriceType, "FAIL");
					}
					else if(Store.equals("App")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<---------  Postage Amounts are different In Approver view orders(OrderInfo) page----------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----Postage Amounts are different In Approver view orders(OrderInfo) page ------>"+ErrorNumber);
						System.out.println("ActualPostageAmount : "+ActualPostageAmount);
						System.out.println("ExpectedPostageAmount : "+ExpectedPostageAmount);
					}
					else if(Store.equals("PS")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<---------  Postage Amounts are different In PS view orders(OrderInfo) page----------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----Postage Amounts are different In PS view orders(OrderInfo) page ------>"+ErrorNumber);
						System.out.println("ActualPostageAmount : "+ActualPostageAmount);
						System.out.println("ExpectedPostageAmount : "+ExpectedPostageAmount);
					}
				}
			}
			//Tax
			String ActualTaxAmount="";
			if(OrderType.equals("Mailinglist")) {
				if(Store.equalsIgnoreCase("user")) {
					ActualTaxAmount= d.findElement(Property.OrderInfoTax).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualTaxAmount= d.findElement(Property.APOrderInfoTax).getText();
				}
			}
			else {
				if(Store.equalsIgnoreCase("user")) {
					ActualTaxAmount= d.findElement(Property.OrderInfoTaxB1).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualTaxAmount= d.findElement(Property.APOrderInfoTax).getText();
				}
			}
			String ExpectedTaxAmount =null;

			ExpectedTaxAmount = Config.Currency + PriceAfterCalculatingTax;
			PriceType  = "Tax";
			if(ActualTaxAmount.equals(ExpectedTaxAmount)) 
			{
				if(Store.equals("user")) {
					System.out.println("Tax Amounts are same In view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page TAX Amounts are same");
					RW_File.WriteResult(ActualTaxAmount, ExpectedTaxAmount,PageName, PriceType, "PASS");
				}
				else if(Store.equals("App")) {
					System.out.println("Tax Amounts are same In Approver view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page TAX Amounts are same in Approver");
				}
				else if(Store.equals("PS")) {
					System.out.println("Tax Amounts are same In PS view orders(OrderInfo) page");
					et.log(LogStatus.INFO, "Order Deatils Page TAX Amounts are same in PS");
				}
			}
			else 
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  Tax Amounts are different In view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<----Tax Amounts are different In view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTaxAmount : "+ActualTaxAmount);
					System.out.println("ExpectedTaxAmount : "+ExpectedTaxAmount);
					RW_File.WriteResult(ActualTaxAmount, ExpectedTaxAmount,PageName, PriceType, "FAIL");
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  Tax Amounts are different In Approver view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<----Tax Amounts are different In Approver view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTaxAmount : "+ActualTaxAmount);
					System.out.println("ExpectedTaxAmount : "+ExpectedTaxAmount);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<---------  Tax Amounts are different In PS view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<----Tax Amounts are different In PS view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTaxAmount : "+ActualTaxAmount);
					System.out.println("ExpectedTaxAmount : "+ExpectedTaxAmount);
				}
			}

			//Total
			String ActualTotal="";
			if(OrderType.equals("Mailinglist")) {
				if(Store.equalsIgnoreCase("user")) {
					ActualTotal = d.findElement(Property.OrderInfoTotal).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualTotal = d.findElement(Property.APOrderInfoTotal).getText();
				}
			}
			else {
				if(Store.equalsIgnoreCase("user")) {
					ActualTotal = d.findElement(Property.OrderInfoTotalB1).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualTotal = d.findElement(Property.APOrderInfoTotal).getText();
				}
			}
			String ExpectedTotal =null;
			PriceType  = "Total";
			ExpectedTotal = Config.Currency + Total;
			if(ActualTotal.equals(ExpectedTotal)) 
			{
				if(Store.equals("user")) {
					System.out.println("Both Total Amounts are same");
					et.log(LogStatus.INFO, "Order Deatils Page Total Amounts are same");
					RW_File.WriteResult(ActualTotal, ExpectedTotal,PageName, PriceType, "PASS");
				}
				else if(Store.equals("App")) {
					System.out.println("Both Total Amounts are same in Approver");
					et.log(LogStatus.INFO, "Order Deatils Page Total Amounts are same in Approver");
				}
				else if(Store.equals("PS")) {
					System.out.println("Both Total Amounts are same in PS");
					et.log(LogStatus.INFO, "Order Deatils Page Total Amounts are same in PS");
				}
			}
			else 
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Total Amounts are different In view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Total Amounts are different In view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTotal : "+ActualTotal);
					System.out.println("ExpectedTotal : "+ExpectedTotal);
					RW_File.WriteResult(ActualTotal, ExpectedTotal,PageName, PriceType, "FAIL");
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Total Amounts are different In Approver view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Total Amounts are different In Approver view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTotal : "+ActualTotal);
					System.out.println("ExpectedTotal : "+ExpectedTotal);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Total Amounts are different In PS view orders(OrderInfo) page----------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Total Amounts are different In PS view orders(OrderInfo) page ------>"+ErrorNumber);
					System.out.println("ActualTotal : "+ActualTotal);
					System.out.println("ExpectedTotal : "+ExpectedTotal);
				}
			}
		}
		catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Verifies pricing and product-related details in the "View Orders" (Order Info) page for different stores (user, App, PS).
	 * Performs verification based on various parameters like subtotal, addons, discount, postage, etc.
	 *
	 * This method validates the following (if applicable):</p>
	 *   Delivery method (e.g., SHIPMENT WITH LIST, MAILING LIST, SHIP TO MY ADDRESS)
	 *   Subtotal price
	 *   Add-on charges
	 *   Discount (based on percentage or flat rate)
	 *   Quantity of products
	 *   Item price per product
	 *   Postage price (for Mailing List orders only)
	 *   Parameters such as Total, Tax, Promotions, Handling Fee, etc. are declared for future use
	 */
	public static void gen5viewordersProductsVerification(String SubTotal, String Addons, String Postage,
			String DiscountPercentage,String Discount,String DiscountCalculationFromSubTotal,String EnablePromotionsORDiscounts,
			String Quantity, String ItemPerPrice, String PageName, String OrderType, String Store)throws InterruptedException 
	{

		try {
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));
			Common.Wait.wait2Second();

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);
			Common.Wait.wait2Second();

			if(Store.equalsIgnoreCase("user")) {
				//Delivery Method
				if(OrderType.equals("ShipmentWithList")) {
					String Deliverymethod = "SHIPMENT WITH LIST";
					String ActualDeliverymethod = d.findElement(By.xpath("//p[normalize-space()='SHIPMENT WITH LIST']")).getText();
					if(Deliverymethod.equals(ActualDeliverymethod))
					{
						System.out.println("Delivery Method is 'SHIPMENT WITH LIST' ");
						et.log(LogStatus.INFO, "Order Deatils Page Delivery Method is SHIPMENT WITH LIST");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Both process type different In view orders(Products) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both proess type different In view orders(Products) page------>"+ErrorNumber);
						System.out.println("Expected value: "+Deliverymethod);
						System.out.println("Actual value :"+ActualDeliverymethod);
					}
				}
				else if(OrderType.equals("Mailinglist"))
				{
					String Deliverymethod = "MAILING LIST";
					String ActualDeliverymethod = d.findElement(By.xpath("//p[normalize-space()='MAILING LIST']")).getText();
					if(Deliverymethod.equals(ActualDeliverymethod))
					{
						System.out.println("Delivery Method is 'MAILING LIST' ");
						et.log(LogStatus.INFO, "Order Deatils Page Delivery Method is Mailing LIST");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Both process type different In view orders(Products) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both proess type different In view orders(Products) page------>"+ErrorNumber);
						System.out.println("Expected value: "+Deliverymethod);
						System.out.println("Actual value :"+ActualDeliverymethod);
					}
				}
				else if(OrderType.equals("Download"))
				{
					String Deliverymethod = "DOWNLOAD";
					String ActualDeliverymethod = d.findElement(By.xpath("//p[normalize-space()='DOWNLOAD']")).getText();
					if(Deliverymethod.equals(ActualDeliverymethod))
					{
						System.out.println("Delivery Method is 'Download' ");
						et.log(LogStatus.INFO, "Order Deatils Page Delivery Method is Download");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Both process type different In view orders(Products) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both proess type different In view orders(Products) page------>"+ErrorNumber);
						System.out.println("Expected value: "+Deliverymethod);
						System.out.println("Actual value :"+ActualDeliverymethod);
					}
				}
				else 
				{
					String Deliverymethod = "SHIP TO MY ADDRESS";

					String ActualDeliverymethod = d.findElement(By.xpath("//p[normalize-space()='SHIP TO MY ADDRESS']")).getText();

					if(Deliverymethod.equals(ActualDeliverymethod))
					{
						System.out.println("Delivery Method is 'SHIP TO MY ADDRESS' ");
						et.log(LogStatus.INFO, "Order Deatils Page Delivery Method is SHIP TO MY ADDRESS");
					}
					else
					{
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------- Both process type different In view orders(Products) page ------> "+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both proess type different In view orders(Products) page------>"+ErrorNumber);
						System.out.println("Expected value: "+Deliverymethod);
						System.out.println("Actual value :"+ActualDeliverymethod);
					}
				}
			}
			// Sub total price verification
			String ActualVOGSubtotal = " ";
			if(OrderType.equals("Mailinglist")) {
				if(Store.equalsIgnoreCase("user")) {
					ActualVOGSubtotal =	d.findElement(Property.ProductSubTotalB2).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualVOGSubtotal =	d.findElement(Property.APProductSubTotal).getText();
				}
			}
			else if(OrderType.equals("ShipmentWithList")) {
				if(Store.equalsIgnoreCase("user")) {
					ActualVOGSubtotal =	d.findElement(Property.ProductSubTotalB1).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualVOGSubtotal =	d.findElement(Property.APProductSubTotal).getText();
				}
			}
			else if(OrderType.equals("Download")) {
				if(Store.equalsIgnoreCase("user")) {
					ActualVOGSubtotal =	d.findElement(Property.ProductSubTotal).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {

					ActualVOGSubtotal =	d.findElement(Property.APProductSubTotal1).getText();
				}
			}
			else {
				if(Store.equalsIgnoreCase("user")) {
					ActualVOGSubtotal =	d.findElement(Property.ProductSubTotal).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {

					ActualVOGSubtotal =	d.findElement(Property.APProductSubTotal).getText();
				}
			}
			String ExpectedVOGSubTotal =null;
			ExpectedVOGSubTotal = Config.Currency + SubTotal;

			if(ActualVOGSubtotal.equals(ExpectedVOGSubTotal))
			{
				if(Store.equals("user")) {
					System.out.println("Both Sub total prices are same In view orders(OrderInfo) page");
				}
				else if(Store.equals("App")) {
					System.out.println("Both Sub total prices are same In Approver view orders(OrderInfo) page");
				}
				else if(Store.equals("PS")) {
					System.out.println("Both Sub total prices are same In PS view orders(OrderInfo) page");
				}
			}
			else
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Both View orders grid subtotal prices are different In view orders(Products) page ------> "+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both View orders grid subtotal prices are different In view orders(Products) page------>"+ErrorNumber);
					System.out.println("Expected value: "+ExpectedVOGSubTotal);
					System.out.println("Actual value :"+ActualVOGSubtotal);
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Both View orders grid subtotal prices are different In Approver view orders(Products) page ------> "+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both View orders grid subtotal prices are different In Approver view orders(Products) page------>"+ErrorNumber);
					System.out.println("Expected value: "+ExpectedVOGSubTotal);
					System.out.println("Actual value :"+ActualVOGSubtotal);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Both View orders grid subtotal prices are different In PS view orders(Products) page ------> "+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both View orders grid subtotal prices are different In PS view orders(Products) page------>"+ErrorNumber);
					System.out.println("Expected value: "+ExpectedVOGSubTotal);
					System.out.println("Actual value :"+ActualVOGSubtotal);
				}
			}

			if(!OrderType.equals("Download")) {
				//Addon price verification
				String ExpectedAddonPrice = Config.Currency+Addons;
				String ActuaAddonPrice="";
				if(OrderType.equals("Mailinglist")) {
					if(Store.equalsIgnoreCase("user")) {
						ActuaAddonPrice = d.findElement(Property.ProductAddOnPriceB2).getText();
					}
					else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
						ActuaAddonPrice = d.findElement(Property.APProductAddOnPrice).getText();
					}
				}
				else if(OrderType.equals("ShipmentWithList")) {
					if(Store.equalsIgnoreCase("user")) {
						ActuaAddonPrice = d.findElement(Property.ProductAddOnPriceB1).getText();
					}
					else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
						ActuaAddonPrice = d.findElement(Property.APProductAddOnPrice).getText();
					}
				}
				else {
					if(Store.equalsIgnoreCase("user")) {
						ActuaAddonPrice = d.findElement(Property.ProductAddOnPrice).getText();
					}
					else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
						ActuaAddonPrice = d.findElement(Property.APProductAddOnPrice).getText();
					}
				}
				if(ActuaAddonPrice.equals(ExpectedAddonPrice))
				{
					if(Store.equals("user")) {
						System.out.println("Both Addon prices are same In view orders(Products) page");
					}
					else if(Store.equals("App")) {
						System.out.println("Both Addon prices are same In Approver view orders(Products) page");
					}
					else if(Store.equals("PS")) {
						System.out.println("Both Addon prices are same In PS view orders(Products) page");
					}
				}
				else
				{
					if(Store.equals("user")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<-------- Both Addon prices are different In view orders(Products) page--------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both Addon prices are different In view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected Addon price In view orders(Products) page : "+ExpectedAddonPrice);
						System.out.println("Actual Addon price In view orders(Products) page: "+ActuaAddonPrice);
					}
					else if(Store.equals("App")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<-------- Both Addon prices are different In Approver view orders(Products) page--------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both Addon prices are different In Approver view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected Addon price In view orders(Products) page : "+ExpectedAddonPrice);
						System.out.println("Actual Addon price In view orders(Products) page: "+ActuaAddonPrice);
					}
					else if(Store.equals("PS")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<-------- Both Addon prices are different In PS view orders(Products) page--------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<---- Both Addon prices are different In PS view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected Addon price In view orders(Products) page : "+ExpectedAddonPrice);
						System.out.println("Actual Addon price In view orders(Products) page: "+ActuaAddonPrice);
					}
				}
			}
			//Discount price verification
			String ExpectedOrderDiscount = null;
			if(DiscountPercentage.equals("N"))
			{
				ExpectedOrderDiscount = "-"+Config.Currency+Discount;
			}
			else if(Discount.equals("0")||Discount.equals("0.0")||Discount.equals("0.00")||
					Discount.equals("0.000")||Discount.equals("0.0000"))
			{
				ExpectedOrderDiscount = "-"+Config.Currency+DiscountCalculationFromSubTotal;
			}
			else
			{
				ExpectedOrderDiscount = "-"+Config.Currency+DiscountCalculationFromSubTotal;	
			}

			if(EnablePromotionsORDiscounts.equals("ON"))
			{  
				String ActualOrderDiscount="";
				if(OrderType.equals("Mailinglist")) {
					if(Store.equalsIgnoreCase("user")) {
						ActualOrderDiscount = d.findElement(Property.ProductDiscountB2).getText();  
					}
				}
				else if(OrderType.equals("ShipmentWithList")) {
					if(Store.equalsIgnoreCase("user")) {
						ActualOrderDiscount = d.findElement(Property.ProductDiscountB1).getText(); 
					}
				}
				else {
					if(Store.equalsIgnoreCase("user")) {
						ActualOrderDiscount = d.findElement(Property.ProductDiscount).getText();
					}
				}
				if(ActualOrderDiscount.equals(ExpectedOrderDiscount))
				{
					if(Store.equals("user")) {
						System.out.println("Both Discount prices are same In view orders(Products) page");
					}
				}
				else
				{
					if(Store.equals("user")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<-------- Both Discount prices are different In view orders(Products) page --------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Both Discount prices are different In view orders(Products) page------>"+ErrorNumber);
						System.out.println("Expected Discount price In view orders(OrderInfo) page : "+ExpectedOrderDiscount);
						System.out.println("Actual Discount price In view orders(OrderInfo) page : "+ActualOrderDiscount);
					}
				}

				//Quantity
				String ExpectedVOGPOPUPProdQuantity = Quantity;
				Thread.sleep(5000);

				String ActualQuantity="";

				if(Store.equalsIgnoreCase("user")) {
					ActualQuantity=d.findElement(Property.ProductQuantity).getText();
				}
				else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
					ActualQuantity=d.findElement(Property.APProductQuantity).getText();
				}

				System.out.println(" Quantiy in View orders Products -->  " + ActualQuantity);

				int ActualQuantity1 = Double.valueOf(ActualQuantity).intValue();
				int Quantity1 = Double.valueOf(Quantity).intValue();
				if(Quantity1==ActualQuantity1){
					if(Store.equals("user")) {
						System.out.println("Both quantitys are same");
					}
					else if(Store.equals("App")) {
						System.out.println("Both quantitys are same in Approver");
					}
					else if(Store.equals("PS")) {
						System.out.println("Both quantitys are same in PS");
					}
				}else
				{
					if(Store.equals("user")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------  Quantitys are different In view orders(Products) page --------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Quantitys are different In view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected  Quantitys : "+ExpectedVOGPOPUPProdQuantity);
						System.out.println("Actual  Quantitys: "+ActualQuantity);
					}
					else if(Store.equals("App")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------  Quantitys are different In Approver view orders(Products) page --------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Quantitys are different In Approver view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected  Quantitys : "+ExpectedVOGPOPUPProdQuantity);
						System.out.println("Actual  Quantitys: "+ActualQuantity);
					}
					else if(Store.equals("PS")) {
						ErrorNumber = ErrorNumber+1;
						captureScreenshot();
						System.out.println("<--------  Quantitys are different In PS view orders(Products) page --------->"+ErrorNumber);
						et.log(LogStatus.ERROR,"<----  Quantitys are different In PS view orders(Products) page ------>"+ErrorNumber);
						System.out.println("Expected  Quantitys : "+ExpectedVOGPOPUPProdQuantity);
						System.out.println("Actual  Quantitys: "+ActualQuantity);
					}
				}

				//Item Price
				if(!OrderType.equals("Download")) {
					String ExpectedProdItemPrice=null;

					ExpectedProdItemPrice = Config.Currency+ItemPerPrice;
					String ActualProdItemPrice="";
					if(OrderType.equals("Mailinglist"))  {
						if(Store.equalsIgnoreCase("user")) {
							ActualProdItemPrice = d.findElement(Property.VoProductItemPriceB2).getText();
						}
						else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
							ActualProdItemPrice = d.findElement(Property.APVoProductItemPrice).getText();
						}
					}
					else if(OrderType.equals("ShipmentWithList")) {
						if(Store.equalsIgnoreCase("user")) {
							ActualProdItemPrice = d.findElement(Property.VoProductItemPriceB1).getText();
						}
						else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
							ActualProdItemPrice = d.findElement(Property.APVoProductItemPrice).getText();
						}
					}
					else {
						if(Store.equalsIgnoreCase("user")) {
							ActualProdItemPrice = d.findElement(Property.VoProductItemPrice).getText();
						}
						else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
							ActualProdItemPrice = d.findElement(Property.APVoProductItemPrice).getText();
						}
					}
					if(ActualProdItemPrice.equals(ExpectedProdItemPrice))
					{
						if(Store.equals("user")) {
							System.out.println(" Item prices are same");
						}
						else if(Store.equals("App")) {
							System.out.println(" Item prices are same in Approver");
						}
						else if(Store.equals("PS")) {
							System.out.println(" Item prices are same in PS");
						}
					}
					else
					{
						if(Store.equals("user")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<--------  Item Prices are different In view orders(Products) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Item Prices are different In view orders(Products) page ------>"+ErrorNumber);
							System.out.println("Expected  Item Prices price : "+ExpectedProdItemPrice);
							System.out.println("Actual Item price : "+ActualProdItemPrice);
						}
						else if(Store.equals("App")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<--------  Item Prices are different In Approver view orders(Products) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Item Prices are different In Approver view orders(Products) page ------>"+ErrorNumber);
							System.out.println("Expected  Item Prices price : "+ExpectedProdItemPrice);
							System.out.println("Actual Item price : "+ActualProdItemPrice);
						}
						else if(Store.equals("PS")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<--------  Item Prices are different In PS view orders(Products) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Item Prices are different In PS view orders(Products) page ------>"+ErrorNumber);
							System.out.println("Expected  Item Prices price : "+ExpectedProdItemPrice);
							System.out.println("Actual Item price : "+ActualProdItemPrice);
						}

					}
				}
				//Postage Price
				if(OrderType.equals("Mailinglist"))  {
					String ExpectedProdPostagePrice = Config.Currency+Postage;
					String ActualProdPostagePrice="";

					if(Store.equalsIgnoreCase("user")) {
						ActualProdPostagePrice = d.findElement(Property.VoProductPostagePrice).getText();
					}
					if(ActualProdPostagePrice.equals(ExpectedProdPostagePrice))
					{
						if(Store.equals("user")) {
							System.out.println(" Postage prices are same");
						}
					}
					else
					{
						if(Store.equals("user")) {
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<--------  Postage Prices are different In view orders(Products) page --------->"+ErrorNumber);
							et.log(LogStatus.ERROR,"<---- Postage Prices are different In view orders(Products) page ------>"+ErrorNumber);
							System.out.println("Expected  Postage Prices price : "+ExpectedProdPostagePrice);
							System.out.println("Actual Postage price : "+ActualProdPostagePrice);
						}
					}
				}
			}
		}
		catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Verifies shipping-related details in the "View Orders" Shipping tab for different user roles (user, App, PS).
	 * Specifically checks for correct shipping provider and shipping price based on expected input.
	 *
	 * This method validates the following (if applicable):
	 *   Shipping provider name (only validated for user store)
	 *   Shipping price per piece for all roles (user, App, PS)
	 */
	public static void gen5viewordersOrderShippingVerification(String OrderBaseShipping, String ShippingPricePerPiece, String PageName, String Store) throws InterruptedException
	{
		try {
			if(Store.equalsIgnoreCase("user")) {
				//Shipping provider name
				String ShippingProviderName = "Custom Shipping" ;
				String ActualShippingProviderName = d.findElement(By.xpath("//p[normalize-space()='Custom Shipping']")).getText();

				if(ShippingProviderName.equals(ActualShippingProviderName))
				{
					System.out.println("Shipping Provider Name is 'Custom Shipping' ");
					et.log(LogStatus.INFO, "Order Deatils Page Shipping Provider Name is 'Custom Shipping'");
				}
				else
				{
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<--------- Both process type different In view orders(Products) page ------> "+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both proess type different In view orders(Products) page------>"+ErrorNumber);
					System.out.println("Expected value: "+ShippingProviderName);
					System.out.println("Actual value :"+ActualShippingProviderName);
				}
			}
			//shipping price verification
			String ActualOrderShippingPrice="";

			if(Store.equalsIgnoreCase("user")) {
				ActualOrderShippingPrice= d.findElement(Property.voshipping).getText();
			}
			else if(Store.equalsIgnoreCase("App") || Store.equalsIgnoreCase("PS")) {
				ActualOrderShippingPrice= d.findElement(Property.APvoshipping).getText();
			}
			String ExpectedOrderShippingPrice =null;

			ExpectedOrderShippingPrice = Config.Currency + ShippingPricePerPiece;


			if(ActualOrderShippingPrice.equals(ExpectedOrderShippingPrice))
			{
				if(Store.equals("user")) {
					System.out.println("Both Shipping prices are same In view orders(ShippingTab) page");
				}
				else if(Store.equals("App")) {
					System.out.println("Both Shipping prices are same In Approver view orders(ShippingTab) page");
				}
				else if(Store.equals("PS")) {
					System.out.println("Both Shipping prices are same In PS view orders(ShippingTab) page");
				}
			}
			else
			{
				if(Store.equals("user")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In view orders(ShippingTab) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both Shipping prices are different In view orders(ShippingTab) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(Shipping) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(Shipping) page: "+ActualOrderShippingPrice);
				}
				else if(Store.equals("App")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In Approver view orders(ShippingTab) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both Shipping prices are different In Approver view orders(ShippingTab) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(shipping) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(shipping) page: "+ActualOrderShippingPrice);
				}
				else if(Store.equals("PS")) {
					ErrorNumber = ErrorNumber+1;
					captureScreenshot();
					System.out.println("<-------- Both Shipping prices are different In PS view orders(ShippingTab) page --------->"+ErrorNumber);
					et.log(LogStatus.ERROR,"<---- Both Shipping prices are different In PS view orders(ShippingTab) page  ------>"+ErrorNumber);
					System.out.println("Expected Shipping price In view orders(shipping) page : "+ExpectedOrderShippingPrice);
					System.out.println("Actual Shipping price In view orders(shipping) page: "+ActualOrderShippingPrice);
				}	
			}
		}

		catch (Exception e){
			ErrorNumber = ErrorNumber +1;
			captureScreenshot();
			e.printStackTrace();
		}
	}
	/**
	 * Logs in the Approver by navigating to the Admin URL,
	 * entering the username and password, and clicking the login button.
	 * Logs success or failure of the login process.
	 */
	public static void ApproverLogin()
			throws InterruptedException, IOException {
		try{

			MouseAdjFunction();
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			d.get(Config.Adminurl);
			et.log(LogStatus.INFO, "Navigated to User URL: " + Config.Adminurl);
			Wait.wait2Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.username));
			//d.findElement(Property.username).isDisplayed();
			Common.Wait.wait2Second();

			waitfl.until(new Function<WebDriver, WebElement>() 
			{
				public WebElement apply(WebDriver driver) {
					return driver.findElement(Property.UserName);
				}
			});
			Wait.wait5Second();
			d.findElement(Property.username).sendKeys(Config.ApproverNamel1);
			d.findElement(Property.Password).sendKeys(Config.ApproverPwdl1);


			d.findElement(Property.LoginButton).click();
			et.log(LogStatus.PASS, "Approver login successful.");
			Common.Wait.wait5Second();
		}
		catch(Exception e)
		{
			captureScreenshot();
			e.printStackTrace();
			et.log(LogStatus.FAIL, "Approver login failed. Error #" + ErrorNumber);
		}
	}
	/**
	 * Logs in the PS by navigating to the Admin URL,
	 * entering the username and password, and clicking the login button.
	 * Logs success or failure of the login process.
	 */
	public static void PSLogin()
			throws InterruptedException, IOException {
		try{
			MouseAdjFunction();
			new WebDriverWait(d, Duration.ofSeconds(Config.ElementWaitTime));

			FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
			waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
			waitfl.pollingEvery(Duration.ofSeconds(5));
			waitfl.ignoring(NoSuchElementException.class);
			waitfl.ignoring(StaleElementReferenceException.class);

			et.log(LogStatus.INFO, "Navigated to User URL: " + Config.Adminurl);
			Common.Wait.wait25Second();
			waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.username));
			waitfl.until(ExpectedConditions.elementToBeClickable(Property.username));

			waitfl.until(new Function<WebDriver, WebElement>() 
			{
				public WebElement apply(WebDriver driver) 
				{
					return driver.findElement(Property.UserName);
				}
			});
			Thread.sleep(15000);
			d.findElement(Property.username).sendKeys(Config.PSNamel1);
			d.findElement(Property.Password).sendKeys(Config.PSPwdl1);

			d.findElement(Property.LoginButton).click();
			et.log(LogStatus.PASS, "PS login successful.");
			Common.Wait.wait5Second();
		}
		catch(Exception e)
		{
			captureScreenshot();
			e.printStackTrace();
			et.log(LogStatus.FAIL, "PS login failed. Error #" + ErrorNumber);
		}
	}
	public static void copyExcelFile(String sourcePath, String destinationFilePath) throws IOException {
	    File sourceFile = new File(sourcePath);
	    File destFile = new File(destinationFilePath);

	    // Ensure parent directory exists
	    File parentDir = destFile.getParentFile();
	    if (!parentDir.exists()) {
	        parentDir.mkdirs();
	    }

	    Files.copy(sourceFile.toPath(), destFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

	    //System.out.println("File copied to: " + destFile.getAbsolutePath());
	}

	public static void addStatusColumn(String filePath) throws IOException {
	    FileInputStream fis = new FileInputStream(filePath);
	    Workbook workbook = new XSSFWorkbook(fis);

	    Sheet sheet = workbook.getSheet("TestData"); // Access "TestData" sheet

	    if (sheet == null) {
	        System.out.println("Sheet 'TestData' not found.");
	        fis.close();
	        return;
	    }

	    Row headerRow = sheet.getRow(0);
	    if (headerRow == null) {
	        headerRow = sheet.createRow(0);
	    }

	    int lastColumnIndex = headerRow.getLastCellNum();
	    if (lastColumnIndex < 0) lastColumnIndex = 0;

	    // Add "Status" column header
	    Cell statusHeader = headerRow.createCell(lastColumnIndex);
	    statusHeader.setCellValue("Status");

	    // Add empty "Status" cells to each row
	    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row == null) continue;

	        Cell statusCell = row.createCell(lastColumnIndex);
	        statusCell.setCellValue(""); // Placeholder
	    }

	    fis.close();

	    FileOutputStream fos = new FileOutputStream(filePath);
	    workbook.write(fos);
	    fos.close();

	    // NO workbook.close() for older POI versions

	   // System.out.println("Status column added to 'TestData' sheet.");
	}
	public static void updateStatus(String filePath, int rowNumberToUpdate, String result) throws IOException {
	    FileInputStream fis = new FileInputStream(filePath);
	    Workbook workbook = new XSSFWorkbook(fis);
	    Sheet sheet = workbook.getSheet("TestData"); // Use the correct sheet name

	    if (sheet == null) {
	        System.out.println("Sheet 'TestData' not found.");
	        fis.close();
	        return;
	    }

	    Row headerRow = sheet.getRow(0);
	    if (headerRow == null) {
	        System.out.println("Header row not found.");
	        fis.close();
	        return;
	    }

	    int statusColumnIndex = headerRow.getLastCellNum() - 1; // Assuming last column is "Status"
	    int tdColumnIndex = 0; // TD# is in column A (index 0)

	    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row == null) continue;

	        Cell tdCell = row.getCell(tdColumnIndex);
	        if (tdCell == null || tdCell.getCellType() != Cell.CELL_TYPE_NUMERIC) continue;

	        int tdValue = (int) tdCell.getNumericCellValue();

	        Cell statusCell = row.getCell(statusColumnIndex);
	        if (statusCell == null) {
	            statusCell = row.createCell(statusColumnIndex);
	        }

	        if (tdValue == rowNumberToUpdate) {
	            statusCell.setCellValue(result);
	            System.out.println("✅ Updated TD# " + tdValue + " at Excel row " + (i + 1) + " with result: " + result);
	        } else {
	            // Only update NE if empty
	            String current = "";
	            if (statusCell.getCellType() == Cell.CELL_TYPE_STRING) {
	                current = statusCell.getStringCellValue().trim();
	            }
	            if (current.isEmpty()) {
	                statusCell.setCellValue("NE");
	            }
	        }
	    }

	    fis.close();
	    FileOutputStream fos = new FileOutputStream(filePath);
	    workbook.write(fos);
	    fos.close();

	    System.out.println("Status updated in TestData: TD " + rowNumberToUpdate + " -> " + result + ", others -> NE");
	}
	/**
	 * Sends an HTML email with the automation report attached.
	 *
	 * Uses Gmail SMTP (SSL) to send the Extent Report located at
	 * "ExtentReports/Log_{Execution_Time}.html" to predefined recipients.
	 */
	public static void mail() throws InterruptedException, EmailException {
		// Build the attachment
		EmailAttachment attachment = new EmailAttachment();
		String reportPath = System.getProperty("user.dir") + "\\ExtentReports\\Log_" + Execution_Time + ".html";
		attachment.setPath(reportPath);
		attachment.setDisposition(EmailAttachment.ATTACHMENT);
		attachment.setDescription("Automation Results");
		attachment.setName("Log_" + Execution_Time + "_Report.html");

		// Common email setup
		HtmlEmail email = new HtmlEmail();
		email.setSSLOnConnect(true);
		email.attach(attachment);

		if (Config.emailType.equalsIgnoreCase("Gmail")) {
			email.setHostName("smtp.gmail.com");
			email.setSmtpPort(465);
			email.setAuthenticator(new DefaultAuthenticator("markupautotest@gmail.com", "ohiv vyif uqda nfyo")); // Use app password if needed
			email.setFrom("markupautotest@gmail.com");
			email.addTo("Santhoshpothamshetty@gmail.com");
			email.addTo("vasavi.konidena@tecra.com");
			email.addTo("sneha.gittaiah@tecra.com");
			email.addTo("manoj.narra@tecra.com");
			email.addTo("gopisyama.tunikipati@tecra.com");
		} else {
			throw new EmailException("Unsupported email type: " + Config.emailType);
		}

		// Email subject and HTML message
		email.setSubject("Dynamic Xmpie Pricing Automation Script Report");

		String htmlMsg = "<html><body style='font-family:Verdana,sans-serif; font-size:12px;'>"
				+ "Hi Team,<br><br>"
				+ "Please find the attached report of the Dynamic XMPie Pricing Automation script.<br><br>"
				+ "Thanks & Regards,<br>"
				+ "ACGen5 QA Team"
				+ "</body></html>";

		email.setHtmlMsg(htmlMsg);

		email.send();
		System.out.println("Email sent successfully.");
	}
}