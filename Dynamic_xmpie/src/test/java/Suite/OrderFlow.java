package Suite;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.mail.EmailException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import com.google.common.base.Function;
import com.relevantcodes.extentreports.LogStatus;
import Common.Commonclass;
import Common.Wait;
import Config.Config;
import Config.Property;
import ExcelFiles.Xls_Reader;
import Utility.RW_File;
import Utility.Testutil;

/**
 * OrderFlow class manages the execution of pricing flow test cases.
 * Inherits common functionality from Commonclass.
 */
public class OrderFlow extends Commonclass 
{
	// Stores the order number generated in the Order Summary page, used across multiple test views.
	public static String OrderNumber = null;

	// Used to track current test data set in the result Excel file.
	public static String TestData1= null;
	public static String Parameters1= null;
		
	
	// Counters for test result status.
	public static int  pass=0;
	public static int fail=0;
	public static int error=0;

	// Stores user choice if shipping address is different from billing.
	public static String ShipAddSameAsBillAddSub = null;

	// Counter to increment error screenshot image names; resets for each test row.
	public static int ImageNumber= 0;
	public static String CustomRole="";

	// Variables for validating base pricing logic.
	public static int BasePriceIncrementValue = 0;
	public static int BasePriceDownload = 0;
	public static boolean IsBaseDiscountZero = false;
	
	int TestDatavalue1  = 0;
	
	// Define Excel file path
	String excelPath = System.getProperty("user.dir") + "/Testdata/ACGen5_Dynamic_XMPie_OrderFlow.xlsx";

	// Prepare destination folder and filename with timestamp
	DateFormat dateFormat1 = new SimpleDateFormat("_yyyy-MM-dd_hh-mm-ss_a");
	Date date1 = new Date();
	String destFolder = "Output";
	String copiedFilePath = destFolder + "/"+Config.Script+"_Output" + dateFormat1.format(date1) + ".xlsx";

	@BeforeClass
	// Runs once before all test methods in the class 
	public void root() throws InterruptedException{
	}

	@BeforeTest
	// Placeholder method to create required folders before the test suite starts.
	public void CreateFolder() throws IOException
	{
		// Copy Excel file first to new file with timestamped name
		copyExcelFile(excelPath, copiedFilePath);

		// Add "Status" column to copied file if it does not exist
		addStatusColumn(copiedFilePath);
	}

	// Conditionally skips the test if marked as "Skip" in the Excel sheet.
	public void beforeTest() throws IOException 
	{
		if (Testutil.isSkip(datatable_suite1, "TestData")) 
		{
			//Assume.assumeTrue(false); // Used to skip test
		}
	}

	/**
	 * Executes the full pricing flow based on test data.
	 * Each row from the Excel sheet is treated as a separate test case.
	 * Validates calculations for pricing, discounts, tax, shipping, etc.
	 */
	@SuppressWarnings("deprecation")
	@Test(dataProvider = "testParameterData")	
	public void AccuPrice(String TestData, String Parameters, String ProdutType, String OrderType,  
			String OrderBase, String PaymentType, String PaymentSubOpt, String CostCenter, String ShipAddSameAsBillAdd, 
			String WeightPerPackage, String CalculateTaxCondition, String EnablePromotionsORDiscounts,
			String EnableZeroAmountOrder,String BasePrice, String Quantity, String ItemPerPrice, 
			String FlatRate, String DownloadPrice, String SubTotal, String DiscountPercentage, String Discount,
			String DiscountCalculationFromSubTotal,
			String Addons, String Postage, String TotalPrice, 
			String PromotionDiscountPercentage, String PromotionCoupon, String DiscountPrice,  
			String Tax, String PriceAfterApplyingCoupon, String PriceAfterCalculatingTax,String OrderBaseShipping, 
			String Total, String DecimalValue,String Weightdecimalvalue,String Weighttype,String Weightdecimaltext,
			String OrderAmountValue, String IsShippingTaxable, String IsPostageTaxable, 
			String PromotionDiscountAfterSubtractingFromSubTotal, 
			String ShippingPricePerPiece, String AddonPricePerPiece, String PostagePriceperpiece, String Payment1Price, String AmtAfterPayment1, 
			String Payment2Price, String AmtAfterPayment2, String Payment3Price,
			String AmtAfterPayment3, String Payment4Price, String AmtAfterPayment4, String Payment5Price,
			String userordershippingorhandlingfee, String Priceafterapplyinghandilingfee,
			String FullfilmentShippingOrHandlingFeeTypeAmountPercent,
			String FullfilmentShippingMarkupFeeAmountPercentTypeAmountPercent,String FullfilmentShippingOrHandlingFee,
			String FullfilmentShippingMarkupFee,String Priceafterapplyingfulfillmentshippingmarkupfee,String IfShippingaddressIseditble,
			String IsTaxExempt, String ZeroAmountorder, String OrderTypeSplitShip,String Customrole)throws Exception, 
	NoSuchElementException, InterruptedException, NumberFormatException, NullPointerException
	{
		try
		{
			// Index for internal counting.
			int i=1;

			// Assigning current test's custom role.
			CustomRole=Customrole;

			// Assigning identifiers for result tracking.
			TestData1 = TestData;
			Parameters1 = Parameters; 
			i=i+1;

			// Reads promotion code from config.
			String PromotionCode = Config.PromotionCode;

			// Parses and validates numeric test data.
			
			if (TestData != null && !TestData.trim().isEmpty() && TestData.trim().matches("\\d+(\\.\\d+)?")) {
				String TestDatavalue = TestData.trim();
				TestDatavalue1 = Double.valueOf(TestDatavalue).intValue();

				// Decimal conversion of all monetary fields to ensure precision in validations.
				String ItemPerPrice1 = Decimalsetting(ItemPerPrice, DecimalValue);
				ItemPerPrice = ItemPerPrice1; 

				String BasePrice1 = Decimalsetting(BasePrice, DecimalValue);
				BasePrice = BasePrice1; 

				String Discount1 = Decimalsetting(Discount, DecimalValue);
				Discount = Discount1;

				String FlatRate1 = Decimalsetting(FlatRate, DecimalValue);
				FlatRate = FlatRate1;
				
				String DownloadPrice1 =Decimalsetting(DownloadPrice, DecimalValue);
				DownloadPrice = DownloadPrice1;

				String Weightdecimaltext1= Decimalsetting(Weightdecimaltext, DecimalValue);
				Weightdecimaltext = Weightdecimaltext1;

				String userordershippingorhandlingfee1 = Decimalsetting(userordershippingorhandlingfee,DecimalValue);
				userordershippingorhandlingfee = userordershippingorhandlingfee1;

				String FullfilmentShippingOrHandlingFee1 = Decimalsetting(FullfilmentShippingOrHandlingFee,DecimalValue);
				FullfilmentShippingOrHandlingFee = FullfilmentShippingOrHandlingFee1;

				String FullfilmentShippingMarkupFee1 = Decimalsetting(FullfilmentShippingMarkupFee,DecimalValue);
				FullfilmentShippingMarkupFee = FullfilmentShippingMarkupFee1;

				String Priceafterapplyingfulfillmentshippingmarkupfee1 = Decimalsetting(Priceafterapplyingfulfillmentshippingmarkupfee,DecimalValue);
				Priceafterapplyingfulfillmentshippingmarkupfee = Priceafterapplyingfulfillmentshippingmarkupfee1;

				String SubTotal1 = Decimalsetting(SubTotal, DecimalValue);
				SubTotal = SubTotal1;

				String Addons1 = Decimalsetting(Addons, DecimalValue);
				Addons = Addons1;

				String Postage1 = Decimalsetting(Postage, DecimalValue);
				Postage = Postage1;

				String PromotionCoupon1 = Decimalsetting(PromotionCoupon, DecimalValue);
				PromotionCoupon = PromotionCoupon1;

				String DiscountPrice1 = Decimalsetting(DiscountPrice, DecimalValue);
				DiscountPrice = DiscountPrice1;

				String Tax1 = Decimalsetting(Tax, DecimalValue);
				Tax = Tax1;

				String OrderBaseShipping1 = Decimalsetting(OrderBaseShipping, DecimalValue);
				OrderBaseShipping = OrderBaseShipping1;

				String PromotionDiscountAfterSubtractingFromSubTotal1 = Decimalsetting(PromotionDiscountAfterSubtractingFromSubTotal, DecimalValue);
				PromotionDiscountAfterSubtractingFromSubTotal =PromotionDiscountAfterSubtractingFromSubTotal1;

				String ShippingPricePerPiece1 = Decimalsetting(ShippingPricePerPiece, DecimalValue);
				ShippingPricePerPiece = ShippingPricePerPiece1;

				String AddonPricePerPiece1 = Decimalsetting(AddonPricePerPiece, DecimalValue);
				AddonPricePerPiece =AddonPricePerPiece1;

				String PostagePricePerPiece1 = Decimalsetting(PostagePriceperpiece, DecimalValue);
				PostagePriceperpiece = PostagePricePerPiece1;

				String Payment1Price1 = Decimalsetting(Payment1Price, DecimalValue);
				Payment1Price = Payment1Price1;

				String AmtAfterPayment11 = Decimalsetting(AmtAfterPayment1, DecimalValue);
				AmtAfterPayment1 = AmtAfterPayment11;

				String Payment2Price1 = Decimalsetting(Payment2Price, DecimalValue);
				Payment2Price = Payment2Price1;

				String AmtAfterPayment21 = Decimalsetting(AmtAfterPayment2, DecimalValue);
				AmtAfterPayment2 = AmtAfterPayment21;

				String Payment3Price1 = Decimalsetting(Payment3Price, DecimalValue);
				Payment3Price = Payment3Price1;
			}

			// Convert execution range from config values (string) to integers
			int StartRowValue = Double.valueOf(Config.ExecutionStartRow).intValue();
			int EndRowVlaue = 0;

			// If 'n' is given in config, get the last row from Excel sheet. 
			if(Config.ExecutionEndRow.equals("n"))
			{
				EndRowVlaue = Double.valueOf(Xls_Reader.SheetRowcount).intValue();
			}
			else
			{
				EndRowVlaue = Double.valueOf(Config.ExecutionEndRow).intValue();
			}

			// Get number of rows manually selected for execution (if any)
			int SelectedRowsArrayCount = Config.SelectedRows.length;

			// Check whether to execute current test row based on config-defined start/end rows or specific row selection
			if((TestDatavalue1 >= StartRowValue && TestDatavalue1 <= EndRowVlaue && SelectedRowsArrayCount == 0) ||
					(Arrays.asList(Config.SelectedRows).contains((TestDatavalue1))))
			{
				// Proceed only if TestData is valid and numeric
				if (TestData != null && !TestData.trim().isEmpty() && TestData.trim().matches("\\d+(\\.\\d+)?")) 
				{
					//************** Admin settings starts here ****************************

					// Get current timestamp for logging and report filenames
					DateFormat dateFormat = new SimpleDateFormat("_yyyy-MMM-dd_h-mm-ss_a");
					Date date = new Date();
					System.out.println("Time Stamp : "+dateFormat.format(date));

					// Close any previously opened browsers or drivers (Windows-specific)
					String os = System.getProperty("os.name");
					if(os.contains("Windows")) 
					{
						if(Config.browser.equals("GC"))
						{
							// Runtime.getRuntime().exec("taskkill /F /IM chrome.exe");
							System.out.println(" Don't kill browser");
						}
						else if(Config.browser.equals("IE"))
						{
							Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
						}
						else if(Config.browser.equals("FF"))
						{
							Runtime.getRuntime().exec("taskkill /F /IM firefox.exe");
						}
						// Always kill chromedriver to prevent hanging sessions
						Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe");
					}

					// If configured, redirect all console output and errors to a log file for debugging
					if(Config.IsConsoleErrorSave.equalsIgnoreCase("Yes"))
					{
						System.setErr(new PrintStream(new FileOutputStream(System.getProperty("user.dir")+"\\ErrorLog\\"+Commonclass.SheetNameErrorLog+".txt", true), true));
						System.setOut(new PrintStream(new FileOutputStream(System.getProperty("user.dir")+"\\ErrorLog\\"+Commonclass.SheetNameErrorLog+".txt", true), true));
					}

					// Start the test logging in the report
					et.log(LogStatus.INFO, "Test Started: Dynamic XMPie Pricing Automation");
					Common.Wait.wait5Second();        	 

					// Check if discount is zero, which may affect how prices are calculated
					if(Discount.equals("0")||Discount.equals("0.0")|| Discount.equals("0.00")||
							Discount.equals("0.000")||	Discount.equals("0.0000"))
					{	
						IsBaseDiscountZero = true;
					}

					System.out.println("****** ******"+" Test Data Number --> "+TestData+"******* *******");
					et.log(LogStatus.INFO, "****** ******"+" Test Data Number --> "+TestData+"******* *******");

					// Reset image counter (used for error screenshots)
					ImageNumber= 0;

					StartBrowser();  // Launch browser locally

					adminLogin();  // Perform admin login

					System.out.println("**********ADMIN**********");
					System.out.println("Admin Login done");
					et.log(LogStatus.INFO, "Admin Login done");

		 			// If DecimalValue is empty, log info; else, apply all necessary decimal-related settings
					if(DecimalValue.isEmpty()) 
					{
						System.out.println("********** Decimal value is Empty ***********");
						et.log(LogStatus.INFO, "Decimal value is empty.");
					}
					else  
					{
						DecimalvalueSetting(DecimalValue, Tax, IsShippingTaxable, OrderAmountValue,Weightdecimalvalue,Weighttype,
								userordershippingorhandlingfee,PaymentSubOpt,PaymentType, CalculateTaxCondition, 
								EnablePromotionsORDiscounts,FullfilmentShippingOrHandlingFee,FullfilmentShippingMarkupFee,OrderBase,
								EnableZeroAmountOrder, TestData, CostCenter, ShipAddSameAsBillAdd, WeightPerPackage,OrderType);

					}	

					// Set per item price for the product
					ItemPerPrice(ItemPerPrice, FlatRate,Weighttype,Weightdecimaltext, Quantity);

					// Set base price and tax exemption     
					BasePriceSetting(BasePrice, ProdutType, DownloadPrice, IsTaxExempt);

					// set up the discounts price
					Discount(Discount, DiscountPercentage, EnablePromotionsORDiscounts);

					// Set up the Add-on price 
					AddonPrice(Addons, AddonPricePerPiece);

					// Setup the Coupon code price 			
					CouponCodePrice(PromotionDiscountPercentage, PromotionCoupon, EnablePromotionsORDiscounts);

					if(OrderType.equals("Mailinglist")) {
						// Setup the Postage for mailing list orders
						PostageSetting(Postage, PostagePriceperpiece);
					}

					// Setup the Tax price
					if(Tax.isEmpty())
					{
						System.out.println("Tax value is empty ");
						et.log(LogStatus.INFO,"********** Tax value is empty ***********");
					}
					else
					{
						String[] CalculateTaxConditions = CalculateTaxCondition.split("_");
						if(CalculateTaxConditions[0].equals("---Select---") || CalculateTaxConditions[0].equals("Vertex"))
						{
							//No Taxes
						}
						else
						{
							// Apply tax if not exempt or skipped
							TaxSettings(Tax);
						}	
					}

					// Setup the shipping setting
					if(OrderBaseShipping.equals("0") || OrderBaseShipping.equals("0.00") || OrderBaseShipping.equals("0.000") || 
							OrderBaseShipping.equals("0.0000"))
					{
						//NO Shipping Price
					}
					else 
					{
						ShippingPriceSetting(OrderBaseShipping,ShippingPricePerPiece);
					}

					//Setup the shipping/handling fee in the fulfillment location
					if((userordershippingorhandlingfee.equals("0.00")||userordershippingorhandlingfee.equals("0.000")||
							userordershippingorhandlingfee.equals("0.0000"))&&(!(FullfilmentShippingOrHandlingFee.equals("0.00")
									||FullfilmentShippingOrHandlingFee.equals("0.000")||FullfilmentShippingOrHandlingFee.equals("0.0000"))))

					{
						// Add fulfillment fee details
						fullfillmentdetails(FullfilmentShippingOrHandlingFee,FullfilmentShippingMarkupFee,FullfilmentShippingOrHandlingFeeTypeAmountPercent,
								FullfilmentShippingMarkupFeeAmountPercentTypeAmountPercent);
					}
					else
					{
						// Remove fulfillment fees if not applicable
						deleteingfullfillmentdetails(FullfilmentShippingOrHandlingFee,FullfilmentShippingMarkupFee,FullfilmentShippingOrHandlingFeeTypeAmountPercent,
								FullfilmentShippingMarkupFeeAmountPercentTypeAmountPercent);
					} 

					// Update OrderBaseShipping if needed based on ShippingPricePerPiece
					if(ShippingPricePerPiece.equals("0") || ShippingPricePerPiece.equals("0.00") || ShippingPricePerPiece.equals("0.000") || 
							ShippingPricePerPiece.equals("0.0000"))
					{
						//System.out.println("No Need to assign value");
					}
					else
					{
						if(OrderBaseShipping.equals("0") || OrderBaseShipping.equals("0.00") || OrderBaseShipping.equals("0.000") ||
									OrderBaseShipping.equals("0.0000"))
						{

						}
						else
						{
							OrderBaseShipping = ShippingPricePerPiece;
						}
					}

					// If ItemPerPrice is zero, fall back to FlatRate
					double ItemPrice1 = Double.valueOf(ItemPerPrice).doubleValue();

					if(ItemPrice1 == 0)
					{
						ItemPerPrice = FlatRate;
					}

					// Log out admin after setup is complete 
					Common.Wait.wait5Second();
					d.findElement(Property.MyAccountMenu).click();
					Common.Wait.wait2Second();
					d.findElement(Property.Alogout).click();

					et.log(LogStatus.INFO, " Admin Setup is Done so logging out");
					System.out.println("Admin Setup is Done");
					Common.Wait.wait5Second();

					// Setup fluent wait to handle dynamic loading
					FluentWait<WebDriver> waitfl = new FluentWait<WebDriver>(d);
					waitfl.withTimeout(Duration.ofSeconds(Config.ElementWaitTime));
					waitfl.pollingEvery(Duration.ofSeconds(5));
					waitfl.ignoring(NoSuchElementException.class);
					waitfl.ignoring(StaleElementReferenceException.class);

					// Handle mouse actions or UI adjustments
					MouseAdjFunction();					

					// User logs in to the system
					userLogin();
					System.out.println("**********USER**********");
					et.log(LogStatus.INFO, "User Homepage Navigating to the Shopping cart");

					// Adjust mouse (UI-specific fix)
					MouseAdjFunction();

					// Wait until Shopping Cart link is clickable
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.ShoppingCartLinkC));

					// Navigate to Shopping Cart page
					Wait.wait2Second();
					d.findElement(Property.ShoppingCartLinkC).click();
					Common.Wait.wait2Second();

					try 
					{	
						// If cart is already empty
						if(d.findElement(Property.ErrorMsgInShoppingCart).isDisplayed())
						{
							d.findElement(Property.OrgunitName).click(); // Go back to org unit/home
						}
						et.log(LogStatus.PASS, "Shopping Cart is empty");
					}
					catch(NoSuchElementException e) 
					{
						// If cart has items, remove them
						Common.Wait.wait2Second();
						d.findElement(Property.EmptyCartLinkC).click();
						Common.Wait.wait2Second();
						d.findElement(Property.EmptyCartConfirmOKButton).click();

						et.log(LogStatus.PASS, "Shopping Cart is empty");

						// Wait until the empty message reappears
						waitfl.until(ExpectedConditions.elementToBeClickable(Property.ErrorMsgInShoppingCart));
						waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ErrorMsgInShoppingCart));
						waitfl.until(new Function<WebDriver, WebElement>() 
						{
							public WebElement apply(WebDriver driver) {
								return driver.findElement(Property.ErrorMsgInShoppingCart);
							}
						});

						Common.Wait.wait5Second();
						d.findElement(Property.OrgunitName).click();	// Return to org-unit
					}
					MouseAdjFunction();

					waitfl.until(ExpectedConditions.elementToBeClickable(Property.MenuBar));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.MenuBar));
					Common.Wait.wait2Second();

					// Open menu
					d.findElement(Property.MenuBar).click();
					Common.Wait.wait10Second();

					// Navigate through category selection
					d.findElement(Property.Categories).click();
					Common.Wait.wait5Second();
					d.findElement(Property.DynamicCategories).click();

					et.log(LogStatus.PASS, "Navigated to Select template page");
					
					//click on grid view
					d.findElement(By.xpath("//*[@id='productList']/div[2]/div/div/div[1]/div/div/div/button[2]/span")).click();

					MouseAdjFunction();
					if(ProdutType.equals("Dynamic xmpie_AT"))
					{
						waitfl.until(ExpectedConditions.presenceOfElementLocated(Property. DynamicPrintNamel1));
						Common.Wait.wait5Second();
						d.findElement(Property.DynamicPrintNamel1).click(); // Select product
						et.log(LogStatus.PASS, "Selected the Product:"+ProdutType);
					}


					// Format and update all monetary values based on OrderAmountValue

					String FlatRate2 = Decimalsetting(FlatRate, OrderAmountValue);
					FlatRate = FlatRate2;
					
					String DownloadPrice2 =Decimalsetting(DownloadPrice, OrderAmountValue);
					DownloadPrice = DownloadPrice2;

					String SubTotal2 = Decimalsetting(SubTotal, OrderAmountValue);
					SubTotal = SubTotal2;

					String DiscountCalculationFromSubTotal2 = Decimalsetting(DiscountCalculationFromSubTotal, OrderAmountValue);
					DiscountCalculationFromSubTotal = DiscountCalculationFromSubTotal2;

					String TotalPrice2 = Decimalsetting(TotalPrice, OrderAmountValue);
					TotalPrice = TotalPrice2;

					String PromotionCoupon2 = Decimalsetting(PromotionCoupon, OrderAmountValue);
					PromotionCoupon = PromotionCoupon2;

					String DiscountPrice2 = Decimalsetting(DiscountPrice, OrderAmountValue);
					DiscountPrice = DiscountPrice2;

					String Tax2 = Decimalsetting(Tax, OrderAmountValue);
					Tax = Tax2;

					String PriceAfterApplyingCoupon2 = Decimalsetting(PriceAfterApplyingCoupon, OrderAmountValue);
					PriceAfterApplyingCoupon = PriceAfterApplyingCoupon2;

					String PriceAfterCalculatingTax2 = Decimalsetting(PriceAfterCalculatingTax, OrderAmountValue);
					PriceAfterCalculatingTax = PriceAfterCalculatingTax2;

					String OrderBaseShipping2 = Decimalsetting(OrderBaseShipping, OrderAmountValue);
					OrderBaseShipping = OrderBaseShipping2;

					String Totala2 = Decimalsetting(Total, OrderAmountValue);
					Total = Totala2;

					String PromotionDiscountAfterSubtractingFromSubTotal2 = Decimalsetting(PromotionDiscountAfterSubtractingFromSubTotal, OrderAmountValue);
					PromotionDiscountAfterSubtractingFromSubTotal =PromotionDiscountAfterSubtractingFromSubTotal2;

					String ShippingPricePerPiece2 = Decimalsetting(ShippingPricePerPiece, OrderAmountValue);
					ShippingPricePerPiece = ShippingPricePerPiece2;

					String AddonPricePerPiece2 = Decimalsetting(AddonPricePerPiece, OrderAmountValue);
					AddonPricePerPiece =AddonPricePerPiece2;

					String PostagePriceperpiece2 = Decimalsetting(PostagePriceperpiece, OrderAmountValue);
					PostagePriceperpiece =PostagePriceperpiece2;

					String Payment1Price2 = Decimalsetting(Payment1Price, OrderAmountValue);
					Payment1Price = Payment1Price2;

					String AmtAfterPayment12 = Decimalsetting(AmtAfterPayment1, OrderAmountValue);
					AmtAfterPayment1 = AmtAfterPayment12;

					String Payment2Price2 = Decimalsetting(Payment2Price, OrderAmountValue);
					Payment2Price = Payment2Price2;

					String AmtAfterPayment22 = Decimalsetting(AmtAfterPayment2, OrderAmountValue);
					AmtAfterPayment2 = AmtAfterPayment22;

					String Payment3Price2 = Decimalsetting(Payment3Price, OrderAmountValue);
					Payment3Price = Payment3Price2;

					Common.Wait.wait2Second();

					String Quantity1 = String.format("%.0f", new BigDecimal(Quantity)); // Round off quantity to integer

					// Proceed to Personalize Page
					Common.Wait.wait5Second();
					et.log(LogStatus.PASS, "Navigated to Personalize page");
					d.findElement(Property.Next1).click(); // Go to next step


					//Select Quantity page and select Based on Order Type
					et.log(LogStatus.PASS, "Navigated to Select Quantity/List page");
					Common.Wait.wait5Second();
					if(OrderBase.equalsIgnoreCase("Split Ship") && OrderType.equals("ShipToMyAddress") 
							|| OrderType.equals("ShipToMultipleAddress"))
					{
						d.findElement(Property.ShipToMyAddress).click();
						et.log(LogStatus.PASS, "OrderType: ShipToMyAddress");
						Common.Wait.wait2Second();
						String Quantity11 = String.format("%.0f", new BigDecimal(Quantity));
						d.findElement(Property.Quantity).clear();

						d.findElement(Property.Quantity).sendKeys(""+Quantity11);
						et.log(LogStatus.INFO, "Quantity:" +Quantity11);
					}
					else if(OrderType.equals("ShipToMyAddress"))
					{
						d.findElement(Property.ShipToMyAddress).click();
						Common.Wait.wait5Second();
						d.findElement(Property.Quantity).clear();
						d.findElement(Property.Quantity).sendKeys(Quantity1);
						et.log(LogStatus.INFO, "Quantity:" +Quantity1);
						Common.Wait.wait5Second();	
					}
					else if(OrderType.equals("Download")) {
						d.findElement(Property.Download).click();
						Common.Wait.wait5Second();
					}
					else
					{
						// For list-based shipping (mailing or shipment with list)
						WebDriverWait wait = new WebDriverWait(d, Duration.ofSeconds(2));
						JavascriptExecutor js = (JavascriptExecutor) d;
						if(OrderType.equals("Mailinglist")) {
							et.log(LogStatus.PASS, "OrderType: Mailinglist");
							WebElement mailingListTab = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[normalize-space()='Mailing list']")));
							mailingListTab.click();
						}
						else if(OrderType.equals("ShipmentWithList")){
							et.log(LogStatus.PASS, "OrderType: ShipmentWithList");
							WebElement mailingListTab = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[normalize-space()='Shipment with list']")));
							mailingListTab.click();
						}

						// Open the search bar
						try {
							WebElement searchIcon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@class, 'sc-eCVdqJ') and contains(@class, 'eeDYvB')]")));
							js.executeScript("arguments[0].click();", searchIcon);
						} catch (Exception e) {
							System.out.println("Search icon not found or already open.");
						}

						// Type in the appropriate list based on quantity
						WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input[type='text']")));
						js.executeScript("arguments[0].click();", searchInput);
						if (Quantity.equals("5") || Quantity.equals("5.0") || Quantity.equals("5.00")) {
							searchInput.sendKeys(Config.RecordsList5);
							et.log(LogStatus.INFO, "Quantity: 5");
						} else if (Quantity1.equals("10") || Quantity.equals("10.0") || Quantity.equals("10.00")) {
							searchInput.sendKeys(Config.RecordsList10);
							et.log(LogStatus.INFO, "Quantity: 10");
						} else if (Quantity1.equals("2") || Quantity.equals("2.0") || Quantity.equals("2.00")) {
							searchInput.sendKeys(Config.RecordsList2);
							et.log(LogStatus.INFO, "Quantity: 2");
						} else if (Quantity1.equals("7") || Quantity.equals("7.0") || Quantity.equals("7.00")) {
							searchInput.sendKeys(Config.RecordsList7);
							et.log(LogStatus.INFO, "Quantity: 7");
						} else {
							System.out.println("No matching record list for given quantity.");
							et.log(LogStatus.FAIL, "No matching record list for given quantity");
						}
						Common.Wait.wait5Second();
						d.findElement(Property.Listcheckbox).click(); // Select the matching list
					}

					// Proceed to View Summary page
					waitfl.until(ExpectedConditions.elementToBeClickable(Property.ContinueSelectList));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ContinueSelectList));
					d.findElement(Property.ContinueSelectList).click();

					et.log(LogStatus.PASS, "Navigated to View Summary Page");
					Common.Wait.wait5Second();

					// Check if the BasePrice is zero (in various possible formats)
					if(BasePrice.equals("0")|| BasePrice.equals("0.0") || BasePrice.equals("0.00") ||
							BasePrice.equals("0.000") || BasePrice.equals("0.0000"))
					{
						// If base price is 0, no special logic required
					}
					else
					{
						// If base price is NOT zero, perform calculations
						BasePriceIncrementValue = 0;

						// Calculate subtotal including base price logic
						String subtotal = SubTotalCalculation(ItemPerPrice,FlatRate,Quantity1,DiscountCalculationFromSubTotal,
								OrderAmountValue, BasePrice, OrderType);
						SubTotal = subtotal;

						// If the BasePrice increment logic triggered a flag (possibly indicating a downloadable item)
						if(BasePriceIncrementValue == 1)
						{
							// Update item price to match sub total
							ItemPerPrice = ""+subtotal;

							// Reset discount-related values
							DiscountCalculationFromSubTotal = "0.00";
							Discount = "0.00";

							// Format these reset values based on order amount settings
							String DiscountCalculationFromSubTotal1 = Decimalsetting(DiscountCalculationFromSubTotal, OrderAmountValue);
							DiscountCalculationFromSubTotal = DiscountCalculationFromSubTotal1;

							String Discount11 = Decimalsetting(Discount, OrderAmountValue);
							Discount = Discount11;
						}
					}
					//IF Order type as not a download below will execute
					if(!OrderType.equals("Download")) {
						// Check whether Addon checkbox is already selected
						boolean AddonStatus = d.findElement(Property.userAddonCheckbox).isSelected();

						if (AddonStatus) {
							System.out.print("Addon is already selected: " + AddonStatus);
						} else {
							// If not selected, click it to enable addon
							d.findElement(Property.userAddonCheckbox).click();
							System.out.print("Addon was not selected, now clicked.");
						}

						Common.Wait.wait2Second();
						// If Addon price is non-zero, validate pricing logic
						if(Addons.equals("0") || Addons.equals("0.00") || Addons.equals("0.000") || Addons.equals("0.0000"))
						{
							// No addon selected
						}
						else
						{
							// Validate addon pricing
							AddonPriceverify(Addons, AddonPricePerPiece, TestData,Parameters, OrderType);
						}

						// If Addon price per piece is non-zero, override Addons value
						if(AddonPricePerPiece.equals("0") || AddonPricePerPiece.equals("0.00") || AddonPricePerPiece.equals("0.000") || AddonPricePerPiece.equals("0.0000"))
						{
							// Do nothing
						}
						else
						{
							Addons = AddonPricePerPiece;
						}

						// Re-format Discount and Addons based on order amount
						String Discount2 = Decimalsetting(Discount, OrderAmountValue);
						Discount = Discount2;

						String Addons2 = Decimalsetting(Addons, OrderAmountValue);
						Addons = Addons2;

						Postage = PostagePriceperpiece; // Assign final postage value
					}

					// Validate and log all prices on the View Summary page
					ViewSummaryPriceInformation(Quantity, ItemPerPrice, SubTotal, Discount, Addons, Postage, DownloadPrice, TotalPrice, DiscountPercentage,
							DiscountCalculationFromSubTotal, OrderType, TestData, Parameters, 
							ProdutType, OrderBase, EnablePromotionsORDiscounts,Weighttype,DiscountCalculationFromSubTotal,
							Priceafterapplyingfulfillmentshippingmarkupfee,OrderAmountValue,DecimalValue, BasePriceIncrementValue);

					d.findElement(Property.AddToCart).click(); // Add product to cart

					et.log(LogStatus.PASS, "Navigated to Shopping Cart");
					et.log(LogStatus.PASS, "Product is Added to Shopping Cart Sucussfully");

					//Shopping cart page
					waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.Checkout));
					waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.Checkout));
					Common.Wait.wait2Second();

					// Validate Shopping Cart pricing again before proceeding
					ShoppingCartPriceInformation(Quantity,SubTotal, ItemPerPrice, Discount, Addons, TotalPrice, DiscountPercentage,
							OrderType, TestData, Parameters, ProdutType,OrderBase,Weighttype, EnablePromotionsORDiscounts,DiscountCalculationFromSubTotal,
							Priceafterapplyingfulfillmentshippingmarkupfee,OrderAmountValue,DiscountCalculationFromSubTotal,DecimalValue);

					// Proceed to checkout
					Common.Wait.wait10Second();
					d.findElement(Property.Checkout).click();
					et.log(LogStatus.PASS, "Navigated to CheckOut Page");
					Common.Wait.wait5Second();			 

					if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {

						// If it's a Split Ship order with "Ship items to multiple addresses", initiate the multiple shipping process
						if(OrderBase.equalsIgnoreCase("Split Ship") && OrderTypeSplitShip.equalsIgnoreCase("Ship items to multiple adresses")){

							// Click to enter multiple shipping address screen
							d.findElement(Property.MultipleShipping).click() ;
							Common.Wait.wait5Second();

							// Determine split quantity: 5 if total quantity > 10, otherwise use the full quantity
							int Quantity2 = Integer.parseInt(Quantity1);
							int Quantity21 =  0;
							//int Quantity22 = 0;
							if(Quantity2 > 10)
							{
								Quantity21 = Quantity2 -5;
								//Quantity22 = 5;
							}
							else
							{
								Quantity21 = Quantity2;
							}
							Common.Wait.wait5Second();

							// Enter the quantity for the first address
							d.findElement(Property.splitquantity).click();
							d.findElement(Property.splitquantity).clear();
							Common.Wait.wait5Second();
							d.findElement(Property.splitquantity).sendKeys(""+Quantity21);
							Common.Wait.wait5Second();

							// Select address from dropdown and add a new shipping address
							d.findElement(Property.shippingdropdown).click() ;
							Common.Wait.wait5Second();
							d.findElement(Property.btnSelectedAddresses).click() ;
							Common.Wait.wait5Second();
							d.findElement(Property.AddNewSpiltShipAddress).click() ;
							Common.Wait.wait5Second();

							// Choose shipping contact and confirm address selection
							d.findElement(Property.rdbtn_ShippingContact).click() ;
							Common.Wait.wait5Second();
							d.findElement(Property.SelectedAddresses).click() ;
							Common.Wait.wait5Second();
							d.findElement(Property.shippingdropdown).click() ;
							Common.Wait.wait5Second();
							d.findElement(Property.btnSelectedAddresses).click() ;				
						}

						// Skip if the order amount is zero
						if(ZeroAmountorder.equalsIgnoreCase("NO")){

							// Check that shipping address is not editable for certain order types
							if((OrderBase.equalsIgnoreCase("Order")||OrderBase.equalsIgnoreCase("Split Ship")
									|| OrderType.equals("ShipToMyAddress")) && IfShippingaddressIseditble.equalsIgnoreCase("NO"))
							{ 
								// Read shipping address value displayed on checkout page
								String ShippingAddressValue = d.findElement(Property.ShippingAddressValue).getText();
								String ExpectedShippAdd1 = "38345 W.Ten Mile Rd"+"\nFarmington Hills, Michigan, 78795";
								String ExpectedShippAdd1b = null;

								ExpectedShippAdd1b = "38345 W.Ten Mile Rd"+"\nFarmington Hills, Michigan, 78795";
								Common.Wait.wait2Second();

								// Click shipping step and proceed to billing details
								waitfl.until(ExpectedConditions.elementToBeClickable(Property.ordershipping));
								waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ordershipping));
								d.findElement(Property.ordershipping).click();
								Common.Wait.wait5Second();

								//Checkout page->2.Billing deatils Page
								d.findElement(Property.ShippingDetailsToNext).click();
								Common.Wait.wait5Second();

								// Parse the "Same as billing address" flag from test data
								String[] ShipAddSameAsBillAddArrary = ShipAddSameAsBillAdd.split("_");	
								ShipAddSameAsBillAdd = 	ShipAddSameAsBillAddArrary[0];
								int ShipAddSameAsBillAddSize = 	ShipAddSameAsBillAddArrary.length;

								// Check whether checkbox is enabled/selected and handle accordingly
								boolean SameAsBillAddStatus = d.findElement(Property.SameAsBillAddStatus).isEnabled();
								if(ShipAddSameAsBillAdd.equals("YES"))
								{
									ShipAddSameAsBillAddSub = " ";
									if(SameAsBillAddStatus == true)
									{
										System.out.println("As expected Same as Billing Add check box selected successfully");
										et.log(LogStatus.INFO,"As expected Same as Billing Add check box selected successfully");
									}
									else
									{
										ErrorNumber = ErrorNumber+1;
										captureScreenshot();
										System.out.println("<----- In Ordercheckout page Same as Billing Add check not selected ------>"+ErrorNumber);
										et.log(LogStatus.ERROR, "<----- In Ordercheckout page Same as Billing Add check not selected ------>"+ErrorNumber);
										System.out.println("Actual value is :"+"Same as Billing address check box NOT selected to user");
										System.out.println("Expected value is :"+"Same as Billing address check box selected to user");
									}
								}
								else if(ShipAddSameAsBillAdd.equals("NO"))
								{
									d.findElement(Property.ShipAddSameAsBillunCheckbox).click();
									ShipAddSameAsBillAddSub = " ";
									if(SameAsBillAddStatus == true)
									{
										System.out.println("As expected Same as Billing Add check box NOT selected successfully");
									}
									else
									{
										ErrorNumber = ErrorNumber+1;
										captureScreenshot();
										System.out.println("<----- In Ordercheckout page Same as Billing Add check selected ------>"+ErrorNumber);
										et.log(LogStatus.ERROR,"<----- In Ordercheckout page Same as Billing Add check selected ------>"+ErrorNumber);
										System.out.println("Actual value is :"+"Same as Billing address check box selected to user");
										System.out.println("Expected value is :"+"Same as Billing address check box NOT selected to user");
									}
								}
								
								String BillingAddressValue = "";

								if (ShipAddSameAsBillAdd.equalsIgnoreCase("Yes")) {
									d.findElement(By.xpath("//label[@title='Same as the shipping address']//input[@type='checkbox']")).click();
									Common.Wait.wait2Second();
									d.findElement(Property.BillingAddressValue).getText();
									Common.Wait.wait2Second();
									d.findElement(By.xpath("//label[@title='Same as the shipping address']//input[@type='checkbox']")).click();
								} else if (ShipAddSameAsBillAdd.equalsIgnoreCase("No")) {
									WebElement labelElement = d.findElement(By.xpath("//label[@title='Same as the shipping address']"));
									String classAttr = labelElement.getAttribute("class");
									if (classAttr.contains("checked")) {  // or "selected" or similar, depends on the actual HTML
										labelElement.click();  // Click the visual checkbox to uncheck
									}
									else {
										d.findElement(By.xpath("//label[@title='Same as the shipping address']//input[@type='checkbox']")).click();
									}
									d.findElement(Property.BillingAddressValue).getText();
								}
							
								/*
								
								String BillingAddressValue = "";

								// Locate checkbox input once
								WebElement checkbox = d.findElement(
								        By.xpath("//label[@title='Same as the shipping address']//input[@type='checkbox']")
								);

								if (ShipAddSameAsBillAdd.equalsIgnoreCase("Yes")) {

								    // If not selected → select it
								    if (!checkbox.isSelected()) {
								        checkbox.click();
								    }

								    Common.Wait.wait2Second();

								    BillingAddressValue = d.findElement(Property.BillingAddressValue).getText();

								} else if (ShipAddSameAsBillAdd.equalsIgnoreCase("No")) {

								    // If selected → unselect it
									if (checkbox.isSelected()) {
									    checkbox.click();
									}

									// Wait until old DOM refresh completes
									waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.BillingAddressValue));

									// Re-locate element AFTER DOM update
									WebElement billingElement = d.findElement(Property.BillingAddressValue);
									BillingAddressValue = billingElement.getText();
								}
*/
								// Verify Both Shipping and Billing address values --First time 
								String ExpectedBillAdd1 = "38345 W.Ten Mile Rd"+"\nFarmington Hills, Michigan, 78795";

								if(ShipAddSameAsBillAdd.equals("YES"))
								{
									if(BillingAddressValue.equals(ExpectedBillAdd1) && 
											(ShippingAddressValue.equals(ExpectedShippAdd1)))
									{
										System.out.println("Both Billing, Shipping Address are same");
									}
									else
									{
										System.out.println("Actual Billg value is :"+BillingAddressValue);
										System.out.println("Actual Shipping value is :"+ShippingAddressValue);
										System.out.println("ExpectedBillAdd1 :"+ExpectedBillAdd1);
										System.out.println("ExpectedShippAdd1 :"+ExpectedShippAdd1);
									}
								}
								else
								{
									if(BillingAddressValue.equals(ExpectedBillAdd1) && 
											(ShippingAddressValue.equals(ExpectedShippAdd1b)))
									{
										System.out.println("Both Billing, Shipping Address are same");
									}
									else
									{
										System.out.println("Actual Billg value is :"+BillingAddressValue);
										System.out.println("Actual Shipping value is :"+ShippingAddressValue);
										System.out.println("Expected Billing Add :"+ExpectedBillAdd1);
										System.out.println("Expected Shippint Add :"+ExpectedShippAdd1b);
									}
								}

								if (ShipAddSameAsBillAdd.equalsIgnoreCase("Yes")) {
									d.findElement(By.xpath("//label[@title='Same as the shipping address']//input[@type='checkbox']")).click();
									Common.Wait.wait2Second();
									d.findElement(Property.BillingEditLink).click();
								} else if (ShipAddSameAsBillAdd.equalsIgnoreCase("No")) {
									WebElement labelElement = d.findElement(By.xpath("//label[@title='Same as the shipping address']"));
									String classAttr = labelElement.getAttribute("class");
									if (classAttr.contains("checked")) {  // or "selected" or similar, depends on the actual HTML
										labelElement.click();  // Click the visual checkbox to uncheck
									}
									d.findElement(Property.BillingEditLink).click();
								}

								MouseAdjFunction();

								Common.Wait.wait5Second();
								waitfl.until(ExpectedConditions.elementToBeClickable(Property.BillingPopAdd1TextBoxL1));
								Common.Wait.wait2Second();
								d.findElement(Property.BillingPopAdd1TextBoxL1).clear();
								d.findElement(Property.BillingPopAdd1TextBoxL1).sendKeys("38345 W.Ten Mile Rd");
								d.findElement(Property.BillingPopAdd3TextBoxL1).clear();
								d.findElement(Property.BillingPopAdd3TextBoxL1).sendKeys("Farmington Hills");
								d.findElement(Property.BillingPopSaveButton).click();

								String ExpectedBillAdd2  = "38345 W.Ten Mile Rd"+"\nFarmington Hills, Michigan, 78795";

								String BillingAddressValue2 = d.findElement(Property.BillingAddressValue).getText();
								Common.Wait.wait2Second();
								d.findElement(Property.ShipAddSameAsBillunCheckbox).click();

								if(ShipAddSameAsBillAdd.equals("YES"))
								{
									if(BillingAddressValue2.equals(ExpectedBillAdd2) && 
											(ShippingAddressValue.equals(ExpectedShippAdd1)))
									{
										//System.out.println("Both Billing, Shipping Address are same and Shipping address do not have Edit | Select Links");
									}
									else
									{
										System.out.println("Actual Billg value is :"+BillingAddressValue2);
										System.out.println("Actual Shipping value is :"+ShippingAddressValue);
										System.out.println("ExpectedBillAdd2 :"+ExpectedBillAdd2);
										System.out.println("ExpectedShippAdd2 :"+ExpectedShippAdd1);
									}
								}
								else
								{
									d.findElement(Property.ShipAddSameAsBillunCheckbox).click();
									ShipAddSameAsBillAddSub = " ";
									if(BillingAddressValue2.equals(ExpectedBillAdd2) && 
											(ShippingAddressValue.equals(ExpectedShippAdd1b)))
									{
										//System.out.println("Both Billing, Shipping Address are same and Shipping address do not have Edit | Select Links");
									}
									else
									{
										System.out.println("Actual Billg value is :"+BillingAddressValue2);
										System.out.println("Actual Shipping value is :"+ShippingAddressValue);
										System.out.println("Expected Billing Value :"+ExpectedBillAdd2);
										System.out.println("Expected Shipping Value :"+ExpectedShippAdd1b);
									}
								}

								MouseAdjFunction();

								//Below code related to click on Shipping address same as Billing address check
								//box in user flow when admin Shipping Address same as Billing is "NO"

								if(ShipAddSameAsBillAdd.equals("NO") && (ShipAddSameAsBillAddSize == 2))
								{
									ShipAddSameAsBillAddSub = ShipAddSameAsBillAddArrary[1];

									if(ShipAddSameAsBillAddSub.equals("UserChkBoxYES"))
									{
										d.findElement(Property.ShipAddSameAsBillunCheckbox).click();
										ShipAddSameAsBillAddSub = ShipAddSameAsBillAddArrary[1];
										d.findElement(Property.SameAsBillAddStatus).click();
										Common.Wait.wait5Second();
										BillingAddressValue2 = d.findElement(Property.BillingAddressValue).getText();

										if(BillingAddressValue2.equals(ExpectedBillAdd2) && 
												(ShippingAddressValue.equals(ExpectedShippAdd1)))
										{
											//System.out.println("Both Billing, Shipping Address are same and Shipping address do not have Edit | Select Links");
										}
										else
										{		
											System.out.println("Actual Billg value is :"+BillingAddressValue2);
											System.out.println("Actual Shipping value is :"+ExpectedShippAdd1);
										}
									}
								}
							}
						}


						if((Weighttype.equalsIgnoreCase("KGS")||Weighttype.equalsIgnoreCase("LBS"))&&OrderBase.equalsIgnoreCase("Order") )
						{
							waitfl.until(ExpectedConditions.elementToBeClickable(Property.ordershipping));
							waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ordershipping));
							waitfl.until(new Function<WebDriver, WebElement>() 
							{
								public WebElement apply(WebDriver driver) {
									return driver.findElement(Property.ordershipping);
								}
							});
							d.findElement(Property.ordershipping).click();
							Common.Wait.wait5Second();

							d.findElement(Property.ShippingDetailsToNext).click();
							Common.Wait.wait2Second();
						}

						if(ZeroAmountorder.equalsIgnoreCase("YES")){
							d.findElement(By.xpath("//*[@id='chkSameAsBilling']")).click();
						}

						// If the shipping address is editable, update the fields and verify tax
						if(OrderBase.equalsIgnoreCase("Order") && ShipAddSameAsBillAdd.equalsIgnoreCase("NO") && IsShippingTaxable.equalsIgnoreCase("NO") && IfShippingaddressIseditble.equalsIgnoreCase("YES"))
						{

							d.findElement(Property.ShippingEditLink).click();
							Common.Wait.wait5Second();

							// Fill shipping address and select state/country to trigger tax logic
							d.findElement(Property.ShippingPopAdd1TextBoxL1).clear();
							d.findElement(Property.ShippingPopAdd1TextBoxL1).sendKeys("38345 W.Ten Mile Rd");
							d.findElement(Property.ShippingPopAdd3TextBoxL1).clear();
							d.findElement(Property.ShippingPopAdd3TextBoxL1).sendKeys("Farmington Hills");
							d.findElement(Property.ShippingPopZipTextBoxL1).clear();
							d.findElement(Property.ShippingPopZipTextBoxL1).sendKeys("78795");
							Common.Wait.wait2Second();

							Actions kb = new Actions(d);
							d.findElement(Property.ShippingPopStateDropdowL1).click();
							Common.Wait.wait2Second();
							kb.sendKeys("Michigan").perform();
							kb.sendKeys(Keys.ENTER).perform();
							Common.Wait.wait2Second();
							d.findElement(Property.ShippingPopCountryDropdowL1).click();
							Common.Wait.wait2Second();
							kb.sendKeys("USA").perform();
							kb.sendKeys(Keys.ENTER).perform();
							Common.Wait.wait2Second();


							d.findElement(Property.ShippingPopSaveButton).click();
							Common.Wait.wait10Second();

							waitfl.until(ExpectedConditions.elementToBeClickable(Property.ordershipping));
							waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ordershipping));
							waitfl.until(new Function<WebDriver, WebElement>() 
							{
								public WebElement apply(WebDriver driver) {
									return driver.findElement(Property.ordershipping);
								}
							});
							d.findElement(Property.ordershipping).click();
							Common.Wait.wait2Second();

							// Compare expected and actual tax percentage
							String ExpectedTaxPercentage="("+Tax+"%)";
							String ActualTaxPercentage=d.findElement(Property.OCTaxPercentage).getText();
							if(ExpectedTaxPercentage.equals(ActualTaxPercentage))
							{
								System.out.println("Both Tax percentages are same ");
							}
							else
							{
								ErrorNumber = ErrorNumber+1;
								captureScreenshot();
								System.out.println("<------------- Order Checkout Both Tax percentages are different ------------>"+ErrorNumber);
								et.log(LogStatus.ERROR,"<------------- Order Checkout Both Tax percentages are different ------------>"+ErrorNumber);
								System.out.println("ActualOCTaxPercentage : "+ActualTaxPercentage);
								System.out.println("ExpectedTaxpercentage : "+ExpectedTaxPercentage);
							}

							Common.Wait.wait2Second();
							d.findElement(Property.ShippingDetailsToNext).click();
							et.log(LogStatus.PASS, "CheckOut > Shipping Section is Done");
							Common.Wait.wait2Second();	

						}	

					}
					else {

						//Mailing list and Download there is no shipping details page
					}

					Common.Wait.wait5Second();
					int paymentLength = 0;
	/*				
					
					// Skip payment if total is zero
					if(Total.equals("0.00") || Total.equals("0.000") || Total.equals("0.0000") || Total.equals("0"))
					{
						//No need to apply payment methods whenever we have total amount is zero
					}
					else
					{
						boolean paymentsListBoxCount =false;
						String[] PaymentsTypes = null;

						// Handle multiple payment methods if comma-separated
						if(PaymentType.contains(","))
						{
							PaymentsTypes = PaymentType.split(",");
							paymentLength = PaymentsTypes.length;
							System.out.println("paymentLength : "+paymentLength);
							et.log(LogStatus.INFO, "paymentLength : "+paymentLength);

							for(int s=0; s <paymentLength; s++)
							{
								if(paymentLength == (s+1))
								{
									String BalanceAmount =  d.findElement(Property.OCRemainingBalance).getText();
									String BalanceAmount1 = BalanceAmount.substring(1,BalanceAmount.length());
									if(s == 0)
									{
										Payment1Price = BalanceAmount1;
									}
									else if(s ==1)
									{
										Payment2Price = BalanceAmount1;
									}
									else if(s == 2)
									{
										Payment3Price = BalanceAmount1;
									}
									else if(s == 3)
									{
										Payment4Price = BalanceAmount1;
									}
								}

								System.out.println("Order flow s value: "+s);
								et.log(LogStatus.INFO,"Order flow s value: "+s);
								switch(s)
								{
								case 0: //System.out.println("First case");
									SelectPaymentTypeInCheckOutPage(OrderBase, PaymentsTypes[s], TestData, Payment1Price,
											PaymentSubOpt, paymentsListBoxCount, s, paymentLength);	
									break;
								case 1: //System.out.println("Second Case");
									SelectPaymentTypeInCheckOutPage(OrderBase, PaymentsTypes[s], TestData, Payment2Price, 
											PaymentSubOpt, paymentsListBoxCount, s, paymentLength);	
									break;
								case 2: //System.out.println("Third Case");
									SelectPaymentTypeInCheckOutPage(OrderBase, PaymentsTypes[s], TestData, Payment3Price, 
											PaymentSubOpt,paymentsListBoxCount, s, paymentLength);	
									break;
								case 3: //System.out.println("Fourth Case");
									SelectPaymentTypeInCheckOutPage(OrderBase, PaymentsTypes[s], TestData, Payment4Price, 
											PaymentSubOpt, paymentsListBoxCount, s,paymentLength);	
									break;
								case 4: //System.out.println("Fifth Case");
									SelectPaymentTypeInCheckOutPage(OrderBase, PaymentsTypes[s], TestData, Payment5Price,
											PaymentSubOpt, paymentsListBoxCount, s, paymentLength);	
									break;
								}

								if(PaymentsTypes[s].equals("Credit Card"))
								{
									paymentsListBoxCount= true;
								}	

							}
						}
						else
						{
							int x = 9;
							SelectPaymentTypeInCheckOutPage(OrderBase, PaymentType, TestData, Total,
									PaymentSubOpt, paymentsListBoxCount, x, paymentLength);	
						}

						int DecimalValue1= Double.valueOf(OrderAmountValue).intValue();
						String DecimalValue2 = "0.00";
						if(DecimalValue1 == 2)
							DecimalValue2 = "0.00";
						else if(DecimalValue1 == 3)
							DecimalValue2 = "0.000";
						else if(DecimalValue1 == 4)
							DecimalValue2 = "0.0000";

						Common.Wait.wait5Second();
						// Verify that applied and remaining payment amounts are correct
						if(PaymentType.contains(","))
						{
							VerifyMultiAppliedAndRemaingPayments(Payment1Price, Payment2Price, Payment3Price, 
									Payment4Price, Payment5Price, DecimalValue2, 
									PaymentType, paymentLength);
						}
						else
						{
							VerifyAppliedAndRemaingPayments(Total, DecimalValue2, CalculateTaxCondition,OrderBase,Weighttype,
									SubTotal,PromotionCoupon,Addons,DiscountCalculationFromSubTotal,OrderAmountValue,
									userordershippingorhandlingfee, TotalPrice, IsShippingTaxable, Tax, PriceAfterApplyingCoupon, OrderType);
						}
					}
*/
	
					// Skip payment if total is zero
					if (Total.equals("0.00") || Total.equals("0.000") || 
					    Total.equals("0.0000") || Total.equals("0")) {

					    // Do nothing

					} else {

					    boolean paymentsListBoxCount = false;

					    // Check if multi-payment
					    boolean isMultiPayment = (PaymentSubOpt != null && PaymentSubOpt.startsWith(","));

					    // ---------------------------------------
					    // ✅ MULTI PAYMENT LOGIC
					    // ---------------------------------------
					    if (isMultiPayment) {

					        String[] PaymentsTypes = PaymentType.split(",");

					        PaymentSubOpt = PaymentSubOpt.substring(1);
					        String[] temp = PaymentSubOpt.split(",");

					        List<String> validSubOptions = new ArrayList<>();
					        for (String opt : temp) {
					            if (opt != null && !opt.trim().isEmpty()) {
					                validSubOptions.add(opt.trim());
					            }
					        }

					        String[] subOptions = validSubOptions.toArray(new String[0]);

					        int totalPayments = PaymentsTypes.length + subOptions.length;

					        String[] paymentAmounts = {
					                Payment1Price,
					                Payment2Price,
					                Payment3Price,
					                Payment4Price,
					                Payment5Price
					        };

					        for (int s = 0; s < totalPayments; s++) {

					            String currentPaymentMethod;
					            String currentAmount;

					            if (s == 0) {
					                currentPaymentMethod = PaymentsTypes[0].trim();
					                currentAmount = paymentAmounts[0];
					            } else {
					                currentPaymentMethod = subOptions[s - 1];
					                currentAmount = paymentAmounts[s];
					            }

					            SelectPaymentTypeInCheckOutPage(
					                    OrderBase,
					                    currentPaymentMethod,
					                    TestData,
					                    currentAmount,
					                    null,
					                    paymentsListBoxCount,
					                    s,
					                    totalPayments
					            );

					            if ("Credit Card".equalsIgnoreCase(currentPaymentMethod)) {
					                paymentsListBoxCount = true;
					            }
					        }

					    }
					    // ---------------------------------------
					    // ✅ SINGLE PAYMENT LOGIC (FIXED)
					    // ---------------------------------------
					    else {

					        // Wait for balance to be visible
					        Common.Wait.wait2Second();

					        String BalanceAmount = d.findElement(Property.OCRemainingBalance)
					                                .getText()
					                                .trim();

					        // Clean currency symbols
					        BalanceAmount = BalanceAmount.replaceAll("[^0-9.]", "");

					        System.out.println("Remaining Balance From UI = " + BalanceAmount);

					        SelectPaymentTypeInCheckOutPage(
					                OrderBase,
					                PaymentType,
					                TestData,
					                BalanceAmount,   // ✅ send UI remaining balance
					                null,
					                paymentsListBoxCount,
					                0,
					                1
					        );

					        VerifyAppliedAndRemaingPayments(
					                BalanceAmount,   // ✅ use balance
					                "0.00",
					                CalculateTaxCondition,
					                OrderBase,
					                Weighttype,
					                SubTotal,
					                PromotionCoupon,
					                Addons,
					                DiscountCalculationFromSubTotal,
					                OrderAmountValue,
					                userordershippingorhandlingfee,
					                TotalPrice,
					                IsShippingTaxable,
					                Tax,
					                PriceAfterApplyingCoupon,
					                OrderType
					        );
					    }

					    Common.Wait.wait2Second();
					}
					Common.Wait.wait2Second();

					// Click Next to go to Review Order Page
					d.findElement(Property.BillingDetailsToNext).click();
					et.log(LogStatus.PASS, "CheckOut > Billing Section is Done");

					// Perform price verification on Review Order Page
					OrderCheckOutPriceInformantion(Quantity,ItemPerPrice, Discount, Addons, TotalPrice, Total, PromotionCode, 
							PromotionCoupon, Tax, PriceAfterCalculatingTax, AddonPricePerPiece, DiscountPercentage,
							PromotionDiscountAfterSubtractingFromSubTotal, PromotionDiscountPercentage, DiscountCalculationFromSubTotal,
							OrderType, TestData,Parameters, ProdutType,OrderBase,OrderBaseShipping,CalculateTaxCondition,
							EnablePromotionsORDiscounts,Weighttype,DiscountCalculationFromSubTotal,SubTotal,OrderAmountValue,userordershippingorhandlingfee,
							Priceafterapplyingfulfillmentshippingmarkupfee,IsShippingTaxable, PriceAfterApplyingCoupon,IsTaxExempt,DecimalValue,
							SubTotal,ShippingPricePerPiece);


					// Cost Center related code
					Common.Wait.wait2Second();
					boolean IsCostcenterDisplayed = false ;

					//System.out.println("IsCostcenterDisplayed :"+IsCostcenterDisplayed);
					if(CostCenter.equals("YES"))
					{
						IsCostcenterDisplayed = d.findElement(Property.CostCenterTag).isDisplayed();

						if((IsCostcenterDisplayed == true))
						{
							// Log message
							System.out.println("Cost Center option displayed successfully");
							et.log(LogStatus.INFO, "Order Checkout Cost Center option is Displayed");
							// Wait for the input to be clickable
							WebDriverWait wait = new WebDriverWait(d, Duration.ofSeconds(5));
							WebElement dropdownInput = wait.until(ExpectedConditions.elementToBeClickable(
									By.xpath("//div[@class='dropdown__input-container css-19bb58m']")));

							// Click to activate dropdown
							dropdownInput.click();

							// Find input box (the actual text input inside the custom dropdown)
							WebElement inputBox = d.switchTo().activeElement();
							inputBox.sendKeys("Cost Center-1");
							inputBox.sendKeys(Keys.ENTER);							
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<----- In Ordercheckout page cost center not displayed ------>"+ErrorNumber);
							et.log(LogStatus.ERROR,"<----- In Ordercheckout page cost center not displayed ------>"+ErrorNumber);
							System.out.println("Actual value is :"+"Cost center option NOT displayed to user");
							System.out.println("Expected value is :"+"Cost center option displayed to user");
						}
						if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
							String Shippingaddressrevieworder  = d.findElement(Property.ShippingAddressValueRO1).getText();

							String BillingAddressrevieworder = d.findElement(Property.BillingAddressValueRO1).getText();

							if(Shippingaddressrevieworder.equals(BillingAddressrevieworder))
							{
								System.out.println("Both Billing, Shipping Address are same in Review order page");
								et.log(LogStatus.INFO, "Order Checkout Both Billing and Shipping Address are same");
							}
							else
							{
								System.out.println("Both Billing, Shipping Address are different in review order page");
								et.log(LogStatus.INFO, "Order Checkout Both Billing and Shipping Address are different");
							}
						}
					}
					else if(CostCenter.equals("NO"))
					{
						if(IsCostcenterDisplayed == false)
						{
							System.out.println("Cost Center option NOT displayed to user successfully");
							et.log(LogStatus.INFO, "Order Checkout Cost Center option is Not Displayed");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<----- In Ordercheckout page cost center displayed ------>"+ErrorNumber);
							et.log(LogStatus.ERROR,"<----- In Ordercheckout page cost center displayed ------>"+ErrorNumber);
							System.out.println("Actual value is :"+"Cost center option displayed to user");
							System.out.println("Expected value is :"+"Cost center option NOT displayed to user");
						}
						if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download")) {
							String Shippingaddressrevieworder1  = d.findElement(Property.ShippingAddressValueRO).getText();
							String BillingAddressrevieworder1 = d.findElement(Property.BillingAddressValueRO).getText();

							if(Shippingaddressrevieworder1.equals(BillingAddressrevieworder1))
							{
								System.out.println("Both Billing, Shipping Address are same in Review order page");
								et.log(LogStatus.INFO, "Order Checkout Both Billing and Shipping Address are same");
							}
							else
							{
								System.out.println("Both Billing, Shipping Address are different in review order page");
								et.log(LogStatus.INFO, "Order Checkout Both Billing and Shipping Address are different");
							}
						}
					}
					d.findElement(Property.AgreementCheck).click();
					Common.Wait.wait5Second();
					d.findElement(Property.SubmitOrder).click();
					et.log(LogStatus.PASS, "CheckOut > Review Order Section is Done");
					et.log(LogStatus.PASS, "Order Placed Sucussfully");
					Common.Wait.wait5Second();

					MouseAdjFunction();
					String[] EnableZeroOrder = EnableZeroAmountOrder.split("_");

					if((Total.equals("0.00") || Total.equals("0.000") || Total.equals("0.0000")) && EnableZeroOrder[0].equals("NO"))
					{
						waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ErrorMsg));
						waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ErrorMsg));
						waitfl.until(new Function<WebDriver, WebElement>() 
						{
							public WebElement apply(WebDriver driver) {
								return driver.findElement(Property.ErrorMsg);
							}
						});
						Common.Wait.wait2Second();
						String ErrorMessage = d.findElement(Property.ErrorMsg).getText();
						System.out.println("ErrorMessage :"+ErrorMessage);
						et.log(LogStatus.ERROR,"ErrorMessage :"+ErrorMessage);

						boolean ErrorMessageStatus = d.findElement(Property.ErrorMsg).isEnabled();
						if(ErrorMessage.equals("Amount should be greater than zero") && ErrorMessageStatus == true)
						{
							// Expected error msg displayed
							//!System.out.println("Expected error msg displayed successfully");
						}
						else
						{
							ErrorNumber = ErrorNumber+1;
							captureScreenshot();
							System.out.println("<----- Ordercheckout Both Error messages are different ------>"+ErrorNumber);
							et.log(LogStatus.ERROR,"<----- Ordercheckout Both Error messages are different ------>"+ErrorNumber);
							System.out.println("Actual Error msg is :"+ErrorMessage);
							et.log(LogStatus.ERROR,"Actual Error msg is :"+ErrorMessage);
							System.out.println("Expected Error Msg is :"+"Amount should be greater than zero");
							et.log(LogStatus.ERROR,"Expected Error Msg is :"+"Amount should be greater than zero");
						}
					}

					if((Total.equals("0.00") || Total.equals("0.000") || Total.equals("0.0000")) && EnableZeroOrder[0].equals("NO"))
					{
						//No need to below code when total values is zero
					}
					else
					{

						Common.Wait.wait2Second();

						MouseAdjFunction();
						WebElement element = d.findElement(Property.GetOrderNumber);
						String textWithSpecialCharacter = element.getText();

						// Use regular expression to remove the special character
						String regex = "[^0-9]"; // This regex removes anything except digits
						Pattern pattern = Pattern.compile(regex);
						Matcher matcher = pattern.matcher(textWithSpecialCharacter);
						String textWithoutSpecialCharacter = matcher.replaceAll("");


						// Store the value in the static variable
						OrderFlow.OrderNumber = textWithoutSpecialCharacter;

						System.out.println("OrderNumber : " + OrderFlow.OrderNumber);
						et.log(LogStatus.INFO, "OrderNumber : " + OrderFlow.OrderNumber);

						OrderSummaryVerification(SubTotal, TotalPrice, Total, ShippingPricePerPiece, Tax, 
								PriceAfterCalculatingTax,userordershippingorhandlingfee, OrderType);

						d.findElement(Property.UserBackToHomeLayout1).click();	
						Common.Wait.wait2Second();
						d.findElement(Property.MenuBar).click();
						Common.Wait.wait2Second();
						d.findElement(Property.UserReportsLinkLayout1).click();
						waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.UserViewOrderImageIconLayout1));
						waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.UserViewOrderImageIconLayout1));
						d.findElement(Property.UserViewOrderImageIconLayout1).click();

						Common.Wait.wait5Second();
						d.findElement(Property.OrderNumberIcon).click();
						Common.Wait.wait2Second();
						d.findElement(Property.OrderNumberIcon).sendKeys(OrderNumber);
						Common.Wait.wait5Second();

						d.findElement(Property.vieworderdeatil).click();
						et.log(LogStatus.PASS, "Navigated to View orders Page");

						String PageName = "Order Details Page";
						String PriceType = null;
						String Store = "user";

						gen5viewordersOrderinfoVerification(SubTotal, Total, Addons, ShippingPricePerPiece,
								Tax, PriceAfterCalculatingTax,  userordershippingorhandlingfee, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts, PromotionDiscountPercentage,
								PromotionCoupon,  PromotionDiscountAfterSubtractingFromSubTotal, PageName, PriceType, OrderType, Store);

						Common.Wait.wait5Second();
						d.findElement(By.xpath("//span[normalize-space()='Products']")).click();

						gen5viewordersProductsVerification(SubTotal, Addons, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts,
								Quantity, ItemPerPrice, PageName, OrderType, Store);

						if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download"))
						{
							Common.Wait.wait5Second();
							d.findElement(By.xpath("//span[normalize-space()='Shipping']")).click();
							gen5viewordersOrderShippingVerification(OrderBaseShipping, ShippingPricePerPiece, PageName, Store);
							Common.Wait.wait2Second();
						}
						else {
							System.out.println("NO shipping Tab in View order");
						}

						//close order Details page
						d.findElement(By.xpath("//button[normalize-space()='Close']")).click();
						et.log(LogStatus.PASS, "Completed the Price verification for the order");
						Common.Wait.wait2Second();
						d.findElement(Property.OrgunitName).click();	
						Common.Wait.wait5Second();
						d.findElement(Property.Gen5logout).click();
						et.log(LogStatus.PASS, "User Logout is Done");
						Common.Wait.wait2Second();


						//Approver login
						ApproverLogin();
						System.out.println("**********APPROVER**********");

						MouseAdjFunction();
						Common.Wait.wait2Second();

						et.log(LogStatus.INFO, "Navigating to Orders > Overview in Approver");
						d.findElement(Property.ApproverOrders).click();
						Common.Wait.wait2Second();
						d.findElement(Property.ApproverOverview).click();
						Common.Wait.wait10Second();
						waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ApproverSearchOrder));
						waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ApproverSearchOrder));
						d.findElement(Property.ApproverSearchOrder).click();
						Common.Wait.wait5Second();
						d.findElement(Property.ApproverSearchOrder).sendKeys(OrderNumber);
						Common.Wait.wait2Second();
						d.findElement(Property.APvieworderdeatil).click();
						et.log(LogStatus.INFO, "Navigated to Approver View orders Page");

						Store = "App";
						// Verify the price values in Approver view order page grid
						gen5viewordersOrderinfoVerification(SubTotal, Total, Addons, OrderBaseShipping,
								Tax, PriceAfterCalculatingTax,  userordershippingorhandlingfee, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts, PromotionDiscountPercentage,
								PromotionCoupon,  PromotionDiscountAfterSubtractingFromSubTotal, PageName, PriceType, OrderType, Store);

						Common.Wait.wait5Second();
						d.findElement(By.xpath("//span[normalize-space()='Products']")).click();

						gen5viewordersProductsVerification(SubTotal, Addons, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts,
								Quantity, ItemPerPrice, PageName, OrderType, Store);

						if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download"))
						{
							Common.Wait.wait5Second();
							d.findElement(By.xpath("//span[normalize-space()='Shipping']")).click();
							gen5viewordersOrderShippingVerification(OrderBaseShipping, ShippingPricePerPiece, PageName, Store);
							Common.Wait.wait2Second();
						}
						else {
							System.out.println("NO shipping Tab in View order");
						}
						Common.Wait.wait2Second();

						//close order Details page
						d.findElement(By.xpath("//button[normalize-space()='Cancel']")).click();
						et.log(LogStatus.PASS, "Completed the Price verification for the order in Approver");
						Common.Wait.wait2Second();
						d.findElement(Property.MyAccountMenu).click();
						Common.Wait.wait2Second();
						d.findElement(Property.Alogout).click();
						et.log(LogStatus.PASS, "Appover Logout is Done");

						//PRODUCTION SUPERVISOR login
						PSLogin();
						System.out.println("**********PRODUCTION SUPERVISOR**********");
						et.log(LogStatus.INFO, "Navigating to Orders > Overview in PS");

						d.findElement(Property.ApproverOrders).click();
						Common.Wait.wait2Second();
						d.findElement(Property.ApproverOverview).click();
						Common.Wait.wait5Second();

						waitfl.until(ExpectedConditions.presenceOfElementLocated(Property.ApproverSearchOrder));
						waitfl.until(ExpectedConditions.visibilityOfElementLocated(Property.ApproverSearchOrder));
						d.findElement(Property.ApproverSearchOrder).click();
						Common.Wait.wait2Second();
						d.findElement(Property.ApproverSearchOrder).sendKeys(OrderNumber);

						d.findElement(Property.APvieworderdeatil).click();
						et.log(LogStatus.PASS, "Navigated to PS View orders Page");

						Store = "PS";
						// Verify the price values in PS view order page grid
						gen5viewordersOrderinfoVerification(SubTotal, Total, Addons, OrderBaseShipping,
								Tax, PriceAfterCalculatingTax,  userordershippingorhandlingfee, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts, PromotionDiscountPercentage,
								PromotionCoupon,  PromotionDiscountAfterSubtractingFromSubTotal, PageName, PriceType, OrderType, Store);

						Common.Wait.wait5Second();
						d.findElement(By.xpath("//span[normalize-space()='Products']")).click();

						gen5viewordersProductsVerification(SubTotal, Addons, Postage,
								DiscountPercentage, Discount, DiscountCalculationFromSubTotal, EnablePromotionsORDiscounts,
								Quantity, ItemPerPrice, PageName, OrderType, Store);

						if(!OrderType.equals("Mailinglist") && !OrderType.equals("Download"))
						{
							Common.Wait.wait5Second();
							d.findElement(By.xpath("//span[normalize-space()='Shipping']")).click();
							gen5viewordersOrderShippingVerification(OrderBaseShipping, ShippingPricePerPiece, PageName, Store);
							Common.Wait.wait2Second();
						}
						else {
							System.out.println("NO shipping Tab in View order");
						}
						Common.Wait.wait2Second();
						
						//close order Details page
						d.findElement(By.xpath("//button[normalize-space()='Cancel']")).click();
						et.log(LogStatus.PASS, "Completed the Price verification for the order in PS");
						Common.Wait.wait2Second();
						d.findElement(Property.MyAccountMenu).click();
						Common.Wait.wait2Second();
						d.findElement(Property.Alogout).click();
						et.log(LogStatus.PASS, "PS Logout is Done"); 
						Common.Wait.wait5Second(); //*/
					} 

					// Check if the order number is generated (i.e., order was placed successfully)
					if(OrderNumber!=null)
					{
						// Log success message to console and ExtentReports
						System.out.println("**** *** **Order Has been Placed successfully*** *** ***");
						et.log(LogStatus.PASS,"**** *** **Order Has been Placed successfully*** *** ***");
					}
					else
					{
						// Log failure if no order number is returned, and display the number of errors
						System.out.println("** ** **Order Has not placed** ** **");
						et.log(LogStatus.FAIL,"** ** **Order Has not placed** ** **");
						System.out.println("Number of errors : "+ErrorNumber);
					}

					System.out.println("___ ___ ___ ____ _ ____ ____ ____ ______ ____ ___ ____ ____ ____ ____ ____ ____ _____ ____ _____ ____ _");
					// Log result and update status on the copied file
					if (ErrorNumber == 0) {
					    System.out.println("*** All the Expected values and Actual values are same, Test PASS ***");
					    et.log(LogStatus.PASS, "*** All the Expected values and Actual values are same, Test PASS ***");
					    updateStatus(copiedFilePath, TestDatavalue1, "Pass");  // Write result on copied file
					    pass++;
					} else {
					    System.out.println("*** Actual and Expected values are not same, Test FAIL ***");
					    et.log(LogStatus.FAIL, "*** Actual and Expected values are not same, Test FAIL ***");
					    System.out.println("Number of errors : " + ErrorNumber);
					    updateStatus(copiedFilePath, TestDatavalue1, "Fail");
					    error++;
					}

					// Stop the WebDriver and reset order number
					stopDriver();
					OrderNumber = null;
				}		
			}
		}
		catch (Exception e)
		{
			// Catch any unexpected exceptions, increment error count, and print stack trace for debugging
			ErrorNumber = ErrorNumber+1;
			error++;
			e.printStackTrace();
			updateStatus(copiedFilePath, TestDatavalue1, "Fail");
		}
	}

	// After all tests, finalize and summarize execution
	@AfterTest
	public void stop() throws IOException, NullPointerException, InterruptedException, EmailException {

		// Calculate total test data rows processed		
		int i=Integer.parseInt(Config.ExecutionEndRow);
		int j=Integer.parseInt(Config.ExecutionStartRow);
		int k=1;

		// Log the number of test cases executed based on row selection
		if((Config.SelectedRows.length==0)){
			System.out.println("Total NUmber of Test Data Executed  : "+(i-j+k));
			et.log(LogStatus.INFO,"Total NUmber of Test Data Executed : "+(i-j+k));

		}
		else
		{	
			System.out.println("Total Number of Test Data Executed  :1 " );
			et.log(LogStatus.INFO,"Total Number of Test Data Executed  :1 " );
		}

		// Output pass/fail summary
		System.out.println("Total Passed Test Data : "+pass);
		et.log(LogStatus.INFO,"Total Passed Test Data : "+pass);

		System.out.println("Total Failed Test Data : "+(fail+error));
		et.log(LogStatus.INFO,"Total Failed Test Data : "+(fail+error));
		RW_File.Closefile();

		// Save and finalize the ExtentReport
		er.flush(); //reports saved
		er.endTest(et);  //reports closed
		System.out.println("Report are saved and closed");
		if(Config.sendMail.equalsIgnoreCase("Yes"))
		{
			//mail(); // Optionally send report via email
		}
	}


	@Test
	@DataProvider(name = "testParameterData", parallel = false)
	// Provides test data from external source to the test method
	public static Object[][] testParameterData(Method method) throws Exception 
	{
		initialize();  // Initialize resources/configs
		Object data[][] = Testutil.getData(datatable_suite1, "TestData");
		return data;
	} 

}