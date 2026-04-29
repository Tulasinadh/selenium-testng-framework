package Config;
import org.openqa.selenium.By;

public class Property {


	//******************  Common Details ****************
	public static By UserName = By.xpath("//*[@id='username']");
	public static By Password = By.xpath("//*[@id='password']");
	public static By LoginButton = By.xpath("//button[@class='login-submit']"); 
	public static By ForgetPassword = By.id("lnkForgotPassword");
	public static By AdminHomeLink = By.xpath("//a[@href='/admin']");


	// ******************** Links under Settings ********************
	public static By Settings = By.xpath("//a[@href='/admin/settings']");
	public static By GeneralSettings = By.xpath("//h3[normalize-space()='General Settings']");
	public static By ShippingSttings = By.xpath("//h3[normalize-space()='Shipping']");
	public static By PaymentSettings = By.xpath("//h3[normalize-space()='Payment Settings']");
	public static By TaxSettings = By.xpath("//h3[normalize-space()='Taxes']");
	public static By FulfillmentLocationsSettings = By.xpath("//h3[normalize-space()='Fulfillment Locations']");
	public static By CollaseALL =By.xpath("//button[normalize-space()='Collapse All']");

	//**Save and Back**
	public static By SettingsSave = By.xpath("//button[normalize-space()='Save']");
	public static By SettingsBack = By.xpath("//a[normalize-space()='Settings']");

	//****Accounting******

	public static By SettingAccounting =By.xpath("//h3[normalize-space()='Accounting']");
	public static By Taxenabled=By.xpath("(//input[@type='checkbox'])[3]");
	public static By TaxProviderDropDown = By.xpath("(//*[name()='svg'][@class='css-8mmkcg'])[4]");
	public static By TaxProviderList = By.xpath("//div[@role='listbox']//div[@role='option']//span");
	public static By SelectedTaxProvider = By.xpath("(//div[@class='css-19bb58m'])[1]");
	public static By EnablePromotionsORDiscounts = By.xpath("(//input[@type='checkbox'])[5]");
	public static By CurrencyDecimaldropdown = By.xpath("(//div[@class='css-7w0czw'])[1]");
	public static By weightdecimaldropdown = By.xpath("(//div[@class='css-7w0czw'])[2]");
	public static By OrderAmoutDecimal = By.xpath("(//div[@class='css-7w0czw'])[3]");
	public static By ShippingisTaxable = By.xpath("(//input[@type='checkbox'])[1]");

	//****Order Management******
	public static By SettingOrderManagement =By.xpath("//h3[normalize-space()='Order Management']");
	public static By EnableZeroAmountOrder = By.xpath("(//input[@type='checkbox'])[4]");
	public static By ShowBillingAddressToZeroAmount = By.xpath("(//input[@type='checkbox'])[5]");


	//************ shipping for enable weight settings**********//


	public static By Enableweightpackage = By.xpath("(//input[@type='checkbox'])[1]");
	public static By ordershippinghandlingfee = By.xpath("(//input[@type='checkbox'])[2]");
	public static By feeentertextbox = By.xpath("//input[@placeholder='Enter fee amount']");

	//***UPS***
	public static By upsdropdown = By.xpath("(//*[name()='svg'][@class='lucide lucide-chevron-right'])[2]");
	public static By upsnextdayselectbox = By.xpath("(//input[@type='checkbox'])[7]");
	public static By Upsshippingselctbox = By.xpath("//input[@title='Select All']");

	//***USPS****
	public static By Uspsdropdown = By.xpath("(//*[name()='svg'][@class='lucide lucide-chevron-right'])[3]");
	public static By Uspsshippingselctbox = By.xpath("//input[@title='Select All']");
	public static By UspsPriorityselectbox = By.xpath("(//input[@type='checkbox'])[7]");	

	//***Shipping Basis***
	public static By shippingbasis = By.xpath("(//div[@aria-hidden='true'])[1]");
	public static By Shippingpicebasis = By.xpath("(//div[@aria-hidden='true'])[2]");
	public static By Defaultlocation = By.xpath("(//div[@aria-hidden='true'])[3]");
	public static By DefaultPackageWeight = By.xpath("(//div[@aria-hidden='true'])[4]");


	//***************** Payment Settings ************************

	//**Billing**
	public static By BillingCheckBox = By.xpath("(//input[@type='checkbox'])[1]");
	public static By BillingSubOption = By.xpath("//input[@name='BILLING']");

	//**COOP Fund**
	public static By CoOpFund = By.xpath("(//input[@type='checkbox'])[5]");
	public static By CoopFundSubOption = By.name("//input[@name='COOP_FUND' and @value='399618']");
	public static By MoneyOnAccountSQL = By.xpath("//input[@name='COOP_FUND' and @value='400420']");
	public static By CFFIFO = By.xpath("(//input[@name='coopfundUsageType'])[1]");
	public static By CFExpiryDate = By.xpath("(//input[@name='coopfundUsageType'])[2]");

	//**Credit Card**
	public static By CreditCardStatus = By.xpath("(//input[@type='checkbox'])[2]");

	public static By CCAutoNet = By.xpath("//span[normalize-space()='Authorize.Net']");
	public static By CCBrainTree = By.xpath("//span[normalize-space()='Braintree']");
	public static By CCPayPal = By.xpath("//span[normalize-space()='PayPal']");

	public static By CCAonly = By.xpath("//span[normalize-space()='Authorize Only']");
	public static By CCACap = By.xpath("//span[normalize-space()='Authorize And Capture']");
	public static By CCAChaLat = By.xpath("//span[normalize-space()='Authorize And Charge Later']");

	//**Gift Card**
	public static By GiftcardStatus = By.xpath("(//input[@type='checkbox'])[3]");
	public static By GiftcardSQLRadioButton = By.xpath("//span[normalize-space()='Gift Card SQL']");			

	//**Cost Center**//
	public static By CostCenterStatus = By.xpath("(//input[@type='checkbox'])[8]");
	public static By DisplayAllCostCenterYes = By.xpath("(//input[@type='checkbox'])[9]");

	//**Ship and bill both same
	public static By ShipAddSameAsBill = By.xpath("(//input[@type='checkbox'])[7]");

	// ***************Link Under Products****************
	public static By Product = By.xpath("//span[normalize-space()='Products']");
	public static By PromotionsIconL1 = By.xpath("//a[normalize-space()='Promotions']");

	//**Product**
	public static By ProductsOverview= By.xpath("//a[normalize-space()='Overview']");
	public static By ProductsSearchBox = By.xpath("//input[@placeholder='Search Product...']");
	public static By EditLink = By.xpath("//button[@title='Edit']//*[name()='svg']");
	public static By ProductStatus = By.xpath("//div[@class='status-control']//label//input[@type='checkbox']");
	public static By ProductAlertOK =  By.xpath("//button[normalize-space()='Yes']");
	public static By ProdPricing = By.xpath("//h3[normalize-space()='Pricing']");
	public static By ProductCodeTextBox = By.xpath("//input[@placeholder='Code']");
	public static By BasePriceTextBoxDynamic = By.xpath("//input[@placeholder='Minimum Price']");
	public static By ProductInfoSave = By.xpath("//span[normalize-space()='Save']");
	public static By TaxExemptCheckBox =By.xpath("(//input[@type='checkbox'])[2]");
	public static By DownloadPrice =By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[3]/div[2]/div[6]/div[2]/div/div[1]/div[3]/div/div[2]/input");

	//**Pricing**
	public static By PricingIcon = By.xpath("//a[normalize-space()='Pricing']");
	public static By PricingSearchBox = By.xpath("//input[@placeholder='Search Pricing...']");
	public static By priceDetailsLink = By.xpath("(//*[name()='svg'][@class='lucide lucide-list icon-sm'])[1]");
	public static By PriceEditLink = By.xpath("//button[@class='text-primary focus-ring mr-4']//*[name()='svg']");
	public static By PriceEntertextBox = By.xpath("//input[@placeholder='Unit Price']");
	public static By PriceMinimumQuantity = By.xpath("//input[@placeholder='Minimum Quantity']");
	public static By PriceTypeDropDown = By.xpath("(//div[@class='css-oj4vjh-indicatorContainer'])[1]");
	public static By Weightentertextbox = By.xpath("//input[@placeholder='Shipping Weight']");
	public static By shippingtypedropdown = By.xpath("(//div[@class='css-oj4vjh-indicatorContainer'])[2]");
	public static By selectshippingdropdoen = By.xpath("(//div[@class='css-19bb58m'])[2]");
	public static By PriceDetailsSaveButton = By.xpath("//button[normalize-space()='Save']");


	// **************** Postage ************************
	public static By PostageIcon = By.xpath("//a[normalize-space()='PostagePricing']");
	public static By PostageSearch = By.xpath("//input[@placeholder='Search Postage Pricing...']");
	public static By PosatgeEditLink = By.xpath("//button[@title='Edit Field']");
	public static By PostagePrice = By.xpath("//input[@placeholder='Per Piece Price']");
	public static By MinimumQuantity = By.xpath("//div[3]//input[1]");
	public static By PostagePriceSave = By.xpath("//button[normalize-space()='Save']");
	// **************** Promotions details ************************

	public static By PromotionsSearch = By.xpath("//input[@placeholder='Search Promotion...']");
	public static By PromotionsEdit = By.xpath("//button[@title='Edit']//*[name()='svg']");
	public static By PromotionDiscountToggle = By.xpath("//input[@id='Discount']");
	public static By PromotionDiscountactive = By.xpath("(//input[@type='checkbox'])[2]"); 
	public static By PromotionValue = By.xpath("//input[@placeholder='value']");
	public static By PromotionName = By.xpath("//input[@name='promotionName']");
	public static By PromotionSave = By.xpath("//button[normalize-space()='Save']");
	public static By PromotionCouponValue = By.xpath("//input[@type='number' and contains(@class,'custom-input')]");

	// ******************* Taxes *****************************

	public static By TaxEditLink = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[9]/div/button[1]");
	public static By TaxValue = By.xpath("//input[@placeholder='Tax Percentage']");
	public static By TaxSaveButton = By.xpath("//button[@title='Save Tax Rate']//*[name()='svg']");

	// ********************  Fulfillment  locations ************

	public static By FulFillmentEdit = By.xpath("//button[@title='Edit']");
	public static By ShiipingHandilingEdit = By.xpath("//input[@placeholder='Shipping/Handling Fee']");
	public static By ShippingMarkupEdit = By.xpath("//input[@placeholder='Shipping Markup Fee']");
	public static By ShippingDropDown = By.xpath("(//div[@aria-hidden='true'])[1]");
	public static By MarkupDropDown = By.xpath("(//div[@aria-hidden='true'])[2]");
	public static By FullfillmentSave = By.xpath("//button[normalize-space()='Save']");

	//Login Page
	public static By username = By.xpath("//input[@id='username']");
	public static By pwd = By.xpath("//input[@id='password']");
	public static By Loginbtn = By.id("btnLogin");
	public static By UserLoginButton = By.xpath("//button[normalize-space()='Login']");

	public static By MenuBar = By.xpath("//i[@class='las la-2x la-bars']");
	public static By Categories = By.xpath("//p[normalize-space()='Categories']");
	public static By DynamicCategories = By.xpath("//p[normalize-space()='Xmpie_Ship']");  

	public static By HomeImageL1 = By.xpath("//img[contains(@src,'image_a.jpg')]");
	public static By DynamicPrintNamel1 = By.xpath("//h3[normalize-space()='Dynamic xmpie_AT']");  
	public static By DynamicEmailNamel1 = By.xpath("//a[@title='Dynamic Email']");
	public static By StaticPrintNamel1 = By.xpath("//a[@title='Static Print']");
	public static By StaticInventoryNamel1 = By.xpath("//h3[normalize-space()='Static Inventory']");
	public static By BroadcastNamel1 = By.xpath("//a[@title='Broadcast']");

	public static By DynamicPrintSelect = By.xpath("(//button[contains(text(),'Select')])[21]"); 
	public static By ProductClose = By.xpath("//div[@id='divQuickView']/div[3]/a");
	public static By Next1 = By.xpath("//button[normalize-space()='Next']");

	// Ship to my address
	public static By ShipToMyAddress = By.xpath("//button[normalize-space()='Ship to My address']");
	public static By ShipToMyAddressDownload = By.id("rdbShipToItemProcessType_2");
	public static By ShipToMyAddressDownLoadPrint = By.id("rdbShipToItemProcessType_3");
	public static By Quantity = By.xpath("//input[@value='1']");
	public static By ShippingRefresh = By.xpath("//*[@aria-owns='ddlShipping_listbox']/span/span");
	public static By PhoneTextBox = By.xpath("//input[starts-with(@id, 'SHIP_CF_PHONE-')]");
	public static By ShippingMethodsLink = By.id("lnkShippingMethods");
	public static By MultiShippingMethodsLink = By.id("lnkMultiShipMethods");
	public static By ShippingSelect = By.xpath("//*[@id='ddlShipping_listbox']/li[2]");
	public static By SelectedShippingValue = By.xpath("//*[@aria-owns='ddlShipping_listbox']/span/span");
	public static By OrderBaseSelectedShippingValue = By.xpath("//*[@id='shippingFieldSet']/div/span/span/span");
	public static By OrderBaseSelectedShippingValueL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/p[2]");
	public static By Upsshippingmethod = By.xpath("//*[@id='ddlShipping_listbox']/li[5]");
	public static By UspsShippingmethod = By.xpath("//*[@id='ddlShipping_listbox']/li[5]");
	
	//downlaod
	public static By Download = By.xpath("//button[normalize-space()='Download']");

	//Select list details
	public static By Listsearchbox = By.xpath("//button[@class='sc-eCYdqJ sc-jSMfEi eeDYvB eEKBuT']");
	public static By Listcheckbox = By.xpath("//input[@value='0']");
	public static By DropshipmentWithList = By.id("DeliveryMode_SHIPMENT_WITH_LIST");
	public static By ContinueSelectList = By.xpath("//button[normalize-space()='Next']");

	//view summary page
	public static By Addonprice = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[5]/p[2]");
	public static By AddonpriceB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[5]/p[2]");
	public static By userAddonCheckbox = By.xpath("//span[normalize-space()='Addon 1']"); 
	public static By UserAddonCheckBoxDSWL = By.xpath("//input[starts-with(@id, 'chkAttributeId-')]");
	public static By userAddonSelectList = By.xpath("//div[@id='divAccuItemDetails']/li[6]/select");
	public static By VSTQuantity = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[1]/p[2]");
	public static By VSTQuantityB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[1]/p[2]");
	public static By VSTQuantityB2 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[1]/div[2]/p[2]");
	public static By Layout2VSQuantity = By.id("lblTotalQuantity");

	public static By VSItemPrice = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[2]/p[2]");
	public static By VSItemPriceB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[2]/p[2]");
	public static By VSSubTotal = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[3]/p[2]");
	public static By VSSubTotalB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[3]/p[2]");
	public static By VSSubTotalB2 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[2]/div[2]/p[2]");
	public static By VSDownload = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[2]/div[3]/p[2]");
	public static By VSDiscountApplied = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[4]/p[2]");
	public static By VSDiscountAppliedB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[4]/p[2]");
	public static By VSDiscountAppliedB2 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[2]/div[4]/p[2]");
	public static By VSDiscountPercent = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[4]/p[2]");
	public static By Layout2VSDiscountPercent = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[4]/p[2]");
	public static By Layout2VSDiscountPercentB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[2]/div[4]/p[2]");
	public static By Layout2VSDiscountPercent1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[4]/p[2]");

	public static By VSAddonos = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[5]/p[2]");
	public static By VSAddonosB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[5]/p[2]");
	public static By VSPostagePrice = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[6]/p[2]");
	public static By VSDownloadPrice = By.xpath("//small[starts-with(@id, 'lblDownloadPrice')]");
	public static By VSLayout2DownloadPrice = By.xpath("(//label[@id='lblDownloadPrice'])[2]");
	public static By VSLayout1DownloadPrice = By.xpath("//label[starts-with(@id, 'lblDownloadPrice')]");
	public static By VSTotalPrice = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[3]/div[7]/p[2]");
	public static By VSTotalPriceB1 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[4]/div[8]/p[2]");
	public static By VSTotalPriceB2 = By.xpath("//*[@id='view-summary-step-right-content-wrapper']/div[2]/div[6]/p[2]");
	public static By ViewSummaryData = By.xpath("//label[@id='lblDiscountValue']/preceding-sibling::span");

	//shopping cart
	public static By SCartQuantity = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div/table/tbody/tr[1]/td[2]");
	public static By SCartItemPrice = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div/table/tbody/tr[1]/td[3]"); 
	public static By SCartAmount = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div/table/tbody/tr[1]/td[3]");
	public static By SCartTotal = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/p[2]");

	// Enable promotions or Discounts off Condition
	public static By SCartDownLoadAmountOFF = By.xpath("//table/tbody/tr/td[8]");
	public static By SCartAmountOFF = By.xpath("//table/tbody/tr/td[9]");
	public static By SCartDownLoadAmountOFF1 = By.xpath("//td[8]/div");
	public static By SCartAmountOFF1 = By.xpath("//td[9]/div");
	public static By SCartLandingDownLoadAmountOFF1 = By.xpath("//td[8]/div[2]");
	public static By SCartLandingAmountOFF1 = By.xpath("//td[9]/div[2]");
	public static By AddToCart = By.xpath("//button[normalize-space()='Add to cart']");
	public static By Checkout = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/button");
	public static By VSPageProof = By.id("frmProof");

	public static By ShoppingCartLinkC = By.xpath("//i[@class='las la-2x la-shopping-cart']");
	public static By ContinueShoppingLinkC = By.xpath("//a[starts-with(@id, 'lnkContinueShopping')]");
	public static By ErrorMsgInShoppingCart = By.xpath("//p[normalize-space()='Sorry, the shopping cart is empty!']");
	public static By OrgunitName = By.xpath("//a[@class='sc-jdAMXn geUmcJ']//img[@alt='No image']");

	public static By EmptyCartLinkC = By.xpath("//button[normalize-space()='Empty cart']");
	public static By EmptyCartConfirmOKButton = By.xpath("//button[normalize-space()='Confirm']");
	
	public static By ShippingAddressValueRO = By.xpath("//div[@class='sc-hhgfTD lecFD']//div[2]//div[2]//p[1]");
	public static By BillingAddressValueRO = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/p");
	public static By ShippingAddressValueRO1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/p[1]");  
	public static By BillingAddressValueRO1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div[4]/div[2]/p");

	public static By OrderCheckoutGridQuantity = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/p[1]");
	public static By OrderCheckoutGridItemPrice = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/p[2]");
	public static By OrderCheckoutGridDiscount = By.xpath("//table/tbody/tr/td[8]");//6
	public static By OrderCheckoutGridPostage = By.xpath("//table/tbody/tr/td[9]");//8
	public static By OrderCheckoutGridAmount = By.xpath("//table/tbody/tr/td[10]");//changed index9 to 10
	public static By OrderCheckOutGridDownloadItemPrice = By.xpath("//table/tbody/tr/td[4]");
	public static By OrderCheckOutGridDownloadDiscount = By.xpath("//table/tbody/tr/td[6]");	
	public static By OrderCheckOutGridShippingorPostageBroadCast = By.xpath("//table/tbody/tr/td[7]");
	public static By OrderCheckOutGridDownloadAmount = By.xpath("//table/tbody/tr/td[10]");//8
	public static By OrderCheckOutGridDownloadAmountS = By.xpath("//table/tbody/tr/td[10]");//8

	public static By OrderCheckOutGridOrderBaseAmounto = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/p[3]");
	public static By OrderCheckoutGridAmounto = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/p[3]");
	public static By OrderCheckoutGridQuantitys = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/p[1]");

	//checkout page
	public static By Savedshiptomyaddress = By.xpath("//a[@href='#SavedShipToMyAddress']");	
	public static By MultipleShipping= By.xpath("//a[@id='a_MultipleShipping']");
	public static By AddNewSpiltShipAddress= By.xpath("//a[@href='#SavedShipToMyAddress']");
	public static By rdbtn_ShippingContact= By.xpath("//input[@name='rdbtn_ShippingContact'][1]");


	public static By SameAsBillAddStatus = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/label/div");
	public static By ShipAddSameAsBillunCheckbox = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/label/div/input");
	public static By BillingAddressValue = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div");
	public static By ShippingAddressValue = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div/div");
	public static By ShippingAddressSelectLink = By.id("lnkSelectShippingAddress");

	public static By SplitshipAddress1 = By.xpath("//input[@id='SplitSHIPINFO_ADDRESS1']");
	public static By SplitShipCity = By.xpath("//input[@id='SplitSHIPINFO_CITY']");
	public static By SplitShipZip = By.xpath("//input[@id='SplitSHIPINFO_ZIP']");
	public static By SplitShipCountryDDL = By.xpath(".//*[@id='Contact_4']/div[2]/span/span/span[1]");
	public static By SplitShipStateDDL = By.xpath(".//*[@id='Contact_3']/div[2]/span/span/span[1]");
	public static By SplitShipaddresssave = By.xpath("//input[@name='chkSplitShipToMyaddress']");
	public static By Saveaddress=By.xpath("//button[@name='btnAddNewSave']");
	public static By splitquantity=By.xpath("//input[@placeholder='Quantity']");
	public static By shippingdropdown=By.xpath("//*[@id='divSplitShipEdit']/div[1]/div[1]/span/span/span[1]");
	public static By btnSelectedAddresses=By.xpath("//button[@id='btnSelectedAddresses']");
	public static By SelectedAddresses=By.id("btnGlobalShipAddSave");

	public static By UserViewOrdersShippingDetailsLink = By.xpath("//a[contains(text(),'Shipping Details')]");
	public static By UserViewOrdersShippingDetailsPopUpOK = By.id("btnOrderShippingOk");
	public static By UserViewOrdersBillingAddress = By.id("orderDetailsSection");
	public static By UserViewOrdersShippinAddress = By.xpath("//div[@id='divOrderShipping']/div[2]");
	public static By UserViewOrdersShippinAddressL1 = By.xpath("//div[@id='divOrderShipping']/div/div/div/div[2] ");

	public static By BillingEditLink = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div/button");
	public static By BillingPopAdd1TextBox = By.xpath("//div[@id='Contact_0']/input");
	public static By BillingPopAdd2TextBox = By.xpath("//div[@id='Contact_1']/input");
	public static By BillingPopAdd1TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/form/div[1]/div[1]/div[2]/input");
	public static By BillingPopAdd2TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/form/div[1]/div[2]/div[2]/input");
	public static By BillingPopAdd3TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/form/div[1]/div[3]/div[2]/input");

	public static By BillingPopSaveButton = By.xpath("//button[normalize-space()='Confirm']");
	public static By BillingPopCancelButton = By.id("lnkBillingAddressClose");
	public static By ShippingEditLink = By.xpath("//span[normalize-space()='New address']");
	public static By ShippingPopAdd1TextBox = By.xpath("//div[@id='Contact_0']/input");
	public static By ShippingPopAdd2TextBox = By.xpath("//div[@id='Contact_1']/input");
	public static By ShippingPopZipTextBox = By.xpath("//div[@id='Contact_5']/input");
	public static By ShippingPopStateDropdowL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[4]/div[2]/div/div/div/div[1]");
	public static By ShippingPopStateDropdow = By.xpath("//div[@id='Contact_3']/span/span/span");
	public static By ShippingPopCountryDropdowL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[5]/div[2]/div/div/div/div[1]/div");

	//Split Ship 
	public static By ShipToMyAddress_splitship = By.id("a_SingleShipping");
	public static By ShippingRefresh_splitship = By.xpath("//div[@id='divShippingMethods']/span/span/span");
	public static By ShippingSelect_splitship = By.xpath("//ul[@id='ddlShippingMethods_listbox']/li[2]");
	public static By SelectedShippingValue_splitship = By.xpath("//div[@id='divShippingMethods']/span/span/span");

	public static By ShippingPopZipTextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[6]/div[2]/input");
	public static By ShippingPopAdd1TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[1]/div[2]/input");
	public static By ShippingPopAdd2TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[2]/div[2]/input");
	public static By ShippingPopAdd3TextBoxL1 = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[3]/div[2]/input");

	public static By ShippingPopSaveButton = By.xpath("//button[normalize-space()='Confirm']");
	public static By ShippingPopCancelButton = By.id("lnkShippingAddressClose");

	
	public static By ordershipping = By.xpath("//span[normalize-space()='shipping_AT']");
	public static By ordershippingmethodselection = By.xpath("//*[@id='ddlShippingMethods_listbox']/li[3]");

	public static By ShippingDetailsToNext = By.xpath("//button[normalize-space()='Next']");
	public static By BillingDetailsToNext = By.xpath("//button[normalize-space()='Next']");

	public static By OCHandilingfee = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[5]/p[2]");
	public static By OPhandilingfee = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[5]/p[2]");

	public static By PromotionCodeTextBox = By.xpath("//input[@placeholder='Promotion code']");
	public static By PromotionApplyButton = By.xpath("//button[normalize-space()='Apply']");

	public static By OCSubTotal = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[2]/p[2]");
	public static By OCPromotionDiscount = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[3]/p[2]");
	public static By OCShippingPrice = By.id("lblShippingPrice");
	public static By OCTaxPercentage = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[6]/p[1]");
	public static By OCTaxPercentageB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/div[4]/p[1]");
	public static By OCTaxAmount = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[6]/p[2]");
	public static By OCTaxAmountB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/div[4]/p[2]");
	public static By OCTotal = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[8]/p[2]");
	public static By OCTotalB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/div[6]/p[2]");
	public static By OCTotalShippingPrice = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/p[2]");

	public static By OCBalanceAmountTextBox = By.cssSelector("#txtBalanceAmount");
	public static By OCRemainingBalance = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[2]/div/div[7]/p/span");

	public static By PaymentTypeDownArrow = By.xpath("//*[@aria-owns='ddlPaymentType_listbox']/span/span");
	public static By PaymentTypeDownArrowOrder = By.xpath("//*[@aria-owns='ddlPaymentType_listbox']/span/span");
	public static By PaymentTypeDownArrowOrderBasis = By.xpath("//div[@id='divPayment']/span/span/span[2]/span");
	public static By Layout2PaymentTypeDownArrow = By.xpath("//div[@id='divPayment']/span/span/span/span[2]/span");
	public static By PaymentTypelength = By.id("ddlPaymentType-list");
	public static By PaymentTypeLength1 = By.xpath("//div[@id='ddlPaymentType-list']/ul[@id='ddlPaymentType_listbox']/li[2]");
	public static By PaymentTypelength2 = By.xpath("ddlPaymentType-list");
	public static By PaymentTypelength3 = By.xpath("ddlPaymentType-list");
	public static By PaymentTypeDropDownData2 = By.xpath("(//ul[@id='ddlPaymentType_listbox'])[2]"); 

	public static By AppliedPayment = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/div[8]/p[2]");
	public static By AppliedPaymentB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/div[6]/p[2]");
	public static By AppliedPayment2 = By.xpath("//table[@id='GrayGrid']/tbody/tr[3]/td[4]");
	public static By AppliedPayment3 = By.xpath("//table[@id='GrayGrid']/tbody/tr[4]/td[4]");
	public static By AppliedPayment4 = By.xpath("//table[@id='GrayGrid']/tbody/tr[5]/td[4]");
	public static By AppliedPayment5 = By.xpath("//table[@id='GrayGrid']/tbody/tr[6]/td[4]");
	public static By RemainingBalance = By.id("lblRemaingBalance");
	public static By MultipaymentApplyButton = By.xpath("//button[normalize-space()='Confirm']");
	public static By BillingApplyButton = By.xpath("//button[normalize-space()='Confirm']");
	public static By coopApplyButton = By.xpath("//button[normalize-space()='Confirm']");
	public static By creditcardApplyButton = By.xpath("//button[normalize-space()='Confirm']");


	public static By CreditCardRadiOButton = By.xpath("//input[starts-with(@id, 'rdo_SaveCard_')]");

	public static By BillingPayment = By.xpath("//span[normalize-space()='Monthly billing']");
	public static By PaymentValue = By.xpath("//input[@class='sc-kgflAQ sc-iAkJRg bduioX EatUO amount-input' and @type='text']");
	public static By PONumber = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]/div[2]/form/div[1]/div[2]/input");
	public static By COOPPayment = By.xpath("//span[normalize-space()='Co-Op fund']");
	public static By CreditCardPayment = By.xpath("//p[normalize-space()='Credit card']");
	public static By CreditCardDown = By.xpath("//*[@aria-owns='ddlCreditCardTypes_listbox']/span/span[2]");
	public static By CreditCardDownC = By.xpath("//div[@id='divCardPayment']/section/span/span/span");
	public static By AuthAndChrgLaterRadioButton =  By.xpath("//input[starts-with(@id,'rdo_SaveCard_')]");

	public static By CreditCardTypeList = By.xpath("//*[@id='divCardPayment']/div[2]/div[1]/span[1]/span/span[1]");
	public static By CreditCardNumber = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]/div[2]/form/div[1]/div[1]/div[2]/input");
	public static By CreditCardNameOnCard = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]/div[2]/form/div[1]/div[4]/div[2]/input");
	public static By CreditCardLastNameOnCard = By.id("CCLastName");

	public static By CreditCardCVVNumber = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]/div[2]/form/div[1]/div[6]/div[2]/input");
	public static By CreditCardLastName = By.xpath("//*[@id='layout-component-wrapper']/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div[2]/div[2]/form/div[1]/div[5]/div[2]/input");
	public static By AgreementCheck = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div[2]/div[2]/div/label/div/input");
	public static By AgreementCheck1 = By.xpath("//div[@id='divAppEditMessage']/div/input");
	public static By SubmitOrder = By.xpath("//button[normalize-space()='Submit']");

	public static By ErrorMsg =  By.id("lblErrorMessage");
	
	public static By CostCenterDropDown = By.xpath("//*[@aria-owns='ddlCostCenter_listbox']/span");
	public static By CostCenterDropDownValue = By.xpath("//*[@id='ddlCostCenter_listbox']/li[2]");
	public static By CostCenterTag = By.xpath("//p[normalize-space()='Cost center']");

	public static By Gen5logout = By.xpath("//i[@title='Log out']");
	public static By MyAccountMenu = By.xpath("//button[@id='user-menu']//*[name()='svg']");
	public static By Alogout = By.xpath("//span[normalize-space()='Log out']");

	//order summary page verification
	public static By SubtotalOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[1]/p[2]");
	public static By SubtotalOsB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div[1]/div[1]/p[2]");
	public static By ShippingOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[2]/p[2]");
	public static By HandlingFeeOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[3]/p[2]");
	public static By TaxPercentageOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[4]/p[1]");
	public static By TaxPercentageOsB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div[1]/div[2]/p[1]");
	public static By TaxAmountOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[4]/p[2]");
	public static By TaxAmountOsB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div[1]/div[2]/p[2]");
	public static By TotalOs = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div/div[6]/p[2]");
	public static By TotalOsB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/div[1]/div[4]/p[2]");

	public static By vieworderdeatil = By.xpath("//*[@id='layout-container']/div[2]/div/div[2]/div/div/div[3]/table/tbody/tr/td[1]/div");	
	public static By APvieworderdeatil = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[2]/div[1]/div/div/table/tbody/tr[1]/td[2]");	


	//Gen5 view orders	
	public static By OrderInfoSubTotal = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[2]/p[2]");
	public static By OrderInfoAddOnPrice = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[3]/p[2]");
	public static By OrderInfoDiscount= By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[4]/p[2]");
	public static By OrderInfoPromotionDiscount = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[5]/p[2]");
	public static By OrderInfoShipping = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[5]/p[2]");
	public static By OrderInfoShippingHandling = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[6]/p[2]");
	public static By OrderInfoPostage = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[8]/p[2]");
	public static By OrderInfoTax = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[9]/p[2]");
	public static By OrderInfoTaxB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[8]/p[2]");
	public static By OrderInfoTotal = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[10]/p[2]");
	public static By OrderInfoTotalB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[5]/div[9]/p[2]");

	public static By ProductSubTotal = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[9]/p[2]");
	public static By ProductSubTotalB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[10]/p[2]");
	public static By ProductSubTotalB2 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[12]/p[2]");
	public static By ProductAddOnPrice = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[7]/p[2]");
	public static By ProductAddOnPriceB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[8]/p[2]");
	public static By ProductAddOnPriceB2 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[10]/p[2]");
	public static By ProductDiscount= By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[8]/p[2]");
	public static By ProductDiscountB1= By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[9]/p[2]");
	public static By ProductDiscountB2= By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[11]/p[2]");
	public static By ProductQuantity = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[2]/p[2]");
	public static By VoProductItemPrice = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[6]/p[2]");
	public static By VoProductItemPriceB1 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[7]/p[2]");
	public static By VoProductItemPriceB2 = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[9]/p[2]");
	public static By VoProductPostagePrice = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]/div[8]/p[2]");

	public static By voshipping = By.xpath("//*[@id='layout-container']/div[2]/div/div[1]/div[1]/div/div[3]/div/div[2]/div[3]/div[3]/p[2]");

	// User View orders verification
	public static By GetOrderNumber = By.xpath("//*[@id='layout-container']/div[2]/div/div/div/div/p[2]/a");
	public static By UserBackToHomeLayout1 = By.xpath("//a[@class='sc-jdAMXn geUmcJ']//img[@alt='No image']");

	//Approver
	public static By ApproverOrders = By.xpath("//div[@class='flex items-start']");
	public static By ApproverOverview = By.xpath("//a[normalize-space()='Overview']");
	public static By ApproverSearchOrder = By.xpath("//input[@placeholder='Search Orders...']");

	//App and PS  view orders	
	public static By APOrderInfoSubTotal = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[1]/dd");
	public static By APOrderInfoAddOnPrice = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[2]/dd");
	public static By APOrderInfoDiscount= By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[3]/dd");
	public static By APOrderInfoPromotionDiscount = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[7]/dd");
	public static By APOrderInfoShipping = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[4]/dd");
	public static By APOrderInfoShippingHandling = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[5]/dd");
	public static By APOrderInfoPostage = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[7]/dd");
	public static By APOrderInfoTax = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[8]/dd");
	public static By APOrderInfoTotal = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div[3]/dl/div[9]/dd");

	public static By APProductSubTotal = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/dl/div[8]/dd");
	public static By APProductSubTotal1 = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/dl/div[7]/dd");
	public static By APProductAddOnPrice = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/dl/div[7]/dd");
	public static By APProductQuantity = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/dl/div[2]/dd");
	public static By APVoProductItemPrice = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/dl/div[6]/dd");
	public static By APvoshipping = By.xpath("//*[@id='root']/div/div[2]/main/div/div/div[2]/div[3]/div[1]/div[4]/div/div/div/div[2]/dd");

	public static By UserReportsLinkLayout1 = By.xpath("//p[normalize-space()='Reports']");
	public static By UserViewOrderImageIconLayout1 = By.xpath("//*[@id='header']/header/div[1]/div/div[1]/ul/li/a/p");
	public static By OrderNumberIcon = By.xpath("//input[@placeholder='Search by order number']");

}