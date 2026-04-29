package Utility;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import Common.Commonclass;
import ExcelFiles.Xls_Reader;

public class Testutil extends Commonclass {
	public static boolean isSkip(Xls_Reader xlsFile, String testName) {
		
		for (int rowNum = 2; rowNum <= xlsFile.getRowCount("Suite_Classes"); rowNum++) {
			if (testName.equals(xlsFile.getCellData("Suite_Classes", "TDID",
					rowNum))) {
				if (xlsFile.getCellData("Suite_Classes", "Runmode", rowNum)
						.equals("Y")) {
					return false;
				} else
					return true;
			}

		}
		return true;
	}

	public static Object[][] getData(Xls_Reader xlsFile, String testName) {
		// if the sheet for data exists
		// return data
		
		int rows = xlsFile.getRowCount(testName) - 1;
		if (rows <= 0) {
			Object[][] testData = new Object[1][0];
			System.out.println("A");
			return testData;			
		}
		rows = xlsFile.getRowCount(testName);
		int cols = xlsFile.getColumnCount(testName);
		// Basetest.APP_LOGS.debug("Test Name---"+testName);
		// Basetest.APP_LOGS.debug("total rows---"+rows);
		// Basetest.APP_LOGS.debug("total cols---"+cols);

		Object data[][] = new Object[rows - 1][cols];

		for (int rowNum = 2; rowNum <= rows; rowNum++) {
			for (int colNum = 0; colNum < cols; colNum++) {
				data[rowNum - 2][colNum] = xlsFile.getCellData(testName,
						colNum, rowNum);
			}
		}
		return data;
	}

	// make zip of reports
	public static void zip(String filepath) {
		try {
			
			File inFolder = new File(filepath);
			File outFolder = new File("Reports.zip");
			ZipOutputStream out = new ZipOutputStream(new BufferedOutputStream(
					new FileOutputStream(outFolder)));
			BufferedInputStream in = null;
			byte[] data = new byte[1000];
			String files[] = inFolder.list();
			for (int i = 0; i < files.length; i++) {
				in = new BufferedInputStream(new FileInputStream(
						inFolder.getPath() + "/" + files[i]), 1000);
				out.putNextEntry(new ZipEntry(files[i]));
				int count;
				while ((count = in.read(data, 00, 1000)) != -1) {
					out.write(data, 00, count);
				}
				out.closeEntry();
			}
			out.flush();
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
