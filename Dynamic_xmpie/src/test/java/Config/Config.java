package Config;

public class Config 
{
	public static String Script = "ACGen5_Dynamic_XMPie_OrderFlow";  // Release patch Number
	public static String Adminurl = "http://qa-acg5.accuconnect.com/admin";  
	public static String Userurl = "http://qa-acg5.accuconnect.com/";	
	public static String browser = "GC";  // FF GC IE																																						
	public static String ReleaseNo = "ACGen5_Sprint_21stJul-1stAug2025";  // Release patch Number
	public static String Currency = "$";  // Currency symbol Classic
	public static String PercentageSymbol = "%";  // Percentage symbol
	public static int  ElementWaitTime = 1 ; 			   // Explicit wait time in seconds and reuse it for Fluent waits conditions

	public static String TakeScreenShot = "Yes";         // YES yes Yes NO (other than Yes it is not taken screen shots)
	public static String IsConsoleErrorSave = "yes";  // Yes yes YES and NO
	public static String IsAdjustMOuse = "Yes"; 		   // YES yes no NO
	public static String ExecutionStartRow = "1";      // Any number (it must be equal or less than the Execution End row value)
	public static String ExecutionEndRow = "2";      // Any number or n (IF we enter n, it indicates last row of the Data sheet) 
	public static Integer[] SelectedRows ={4};            // 89,90,94,95,96,101,103,104,105 to execute for selected test cases, if it empty above ranges will work  

	public static String ProductPriceCode = "Price_AT";   // Product Price code
	public static String AddonPriceCode = "Addon_AT";	    // Add-on price code
	public static String ShippingPriceCode = "shipping_AT";
	public static String PostagePriceCode = "Postage_AT";
	public static String sendMail="yes";        // Yes YES NO no
	public static String emailType="Gmail";     // Gmail
	// Shipping Price code
	public static String DiscountName = "Discount_AT1";		         // Discount Name
	public static String DiscountName2 = "Discount_AT2";		     // Discount Name
	public static String CouponCodeName = "Coupon_AT1"; // Coupon code Name
	public static String PromotionCode = "570";
	//List
	public static String RecordsList10 = "Dynamic_10";            // 10 records list name
	public static String RecordsList5 = "Dynamic_05";	                // 7 Records list name
	public static String RecordsList2 = "Dynamic_02";
	public static String RecordsList7 = "Dynamic_07"; 

	// users Details
	public static String UserNamel1 = "Lt_user";
	public static String UserPwdl1 = "welcome";
	public static String AdminNamel1 = "Lt_Admin";	
	public static String AdminPwdl1 = "welcome";
	public static String ApproverNamel1 = "LT_Approver";
	public static String ApproverPwdl1 = "welcome";
	public static String PSNamel1 = "LT_Ps";
	public static String PSPwdl1 = "welcome";
}
