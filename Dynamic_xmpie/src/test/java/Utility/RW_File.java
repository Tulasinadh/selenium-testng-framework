package Utility;



import java.io.File;


import java.io.BufferedWriter;

import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import Common.*;
import Config.Config;
import Suite.OrderFlow;

public class RW_File {

 private static FileWriter fw;
 private static BufferedWriter bw;
 static DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm");
 static Date date = new Date();
 //private static String time = Config.wd+Config.ProjectName+"\\Results\\"+Commonclass.SheetName;
 private static String time = System.getProperty("user.dir")+"\\Results\\"+Commonclass.SheetName;
 public static String FolderPath = null;
 private static String TestFile =time;

public static void CreateFile() throws IOException {
  //Create File In D: Driver.  
 
	DateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd_h-mm-ss_a");
	 Date date = new Date();
	// System.out.println("Time Stamp : "+dateFormat.format(date));
	 
	 // Create folder for error images (and also create folder name)
	 if(Config.TakeScreenShot.equalsIgnoreCase("Yes"))
	 {
		 FolderPath = System.getProperty("user.dir")+"\\ScreenShots\\"+dateFormat.format(date);
		 new File(FolderPath).mkdir();	
	 }
	
	
  File FC = new File(TestFile); //Created object of java File class.
  FC.createNewFile();//Create file.
 // File Fc1 = new File(FolderPath).mkdirs();
 
  //Fc1.createNewFile();
  fw = new FileWriter(TestFile);
  //fw = new FileWriter(TestFile);
   bw = new BufferedWriter(fw);
  // bw.append("Test Data , Parameters , Page Name , Price Type , Actual Value , Expected Value , Status, OrderNumber, Comments");
   bw.append("Test Data , Parameters , Page Name , Price Type , Expected Value , Actual Value , OrderNumber, Status, Comments");
   bw.newLine();
  // System.out.println(time1);
 }

public static void WriteResult(String ActualValue, String ExpectedValue, String PageName, String PriceType, String Status) throws IOException {
    bw.append(OrderFlow.TestData1 + "," + OrderFlow.Parameters1 + "," + PageName + "," + PriceType + ",");
    //bw.append("'" + ActualValue + "," + "'" + ExpectedValue + "," + Status + "," + OrderFlow.OrderNumber);
    bw.append("'" + ExpectedValue + "," + "'" + ActualValue + "," + OrderFlow.OrderNumber + "," +Status );
    bw.newLine();
    bw.flush();
}

public static void Closefile() throws IOException, NullPointerException {
	 bw.close();
	 fw.close();
	 }
  }

