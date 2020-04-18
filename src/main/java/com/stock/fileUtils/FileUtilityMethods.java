package com.stock.fileUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.stock.Models.FileModels;

public class FileUtilityMethods {

	// Instantiating FileModels
	FileModels filemodels = new FileModels();

	// Getting the file to read the names of the company
	File file;
	FileInputStream inputStream;

	/*
	 * @Returns Company name
	 */
	public String GetCompanyNameAt(int row, int col) throws Exception {

		// Inside get company name
		file = new File(filemodels.getCompanyNamesSourcePath());
		//System.out.println("Reading Company names from file : " + filemodels.getCompanyNamesSourcePath());

		try {
			inputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
			e.printStackTrace();
			throw new Exception("File Not Found please check the name");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);

		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(filemodels.getDefaultSheetName());

		// Returns the cell value
		return sheet.getRow(row).getCell(col).getStringCellValue();
	}

	/*
	 * Get any row and column at
	 * 
	 * @params File name
	 */
	public String GetCellData(int row, int col, String FilePath) throws Exception {

		try {
			// System.out.println("File Opened Sucessfully");
			inputStream = new FileInputStream(FilePath);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
			throw new Exception("File Not Found please check the name");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(filemodels.getDataSheetname());

		// System.out.println(sheet.getRow(row).getCell(col).getCellType());
		// Returns the cell value
		try {
			if (sheet.getRow(row).getCell(col).getCellType() == 1) {
				return sheet.getRow(row).getCell(col).getStringCellValue();
			} else if (HSSFDateUtil.isCellDateFormatted(sheet.getRow(row).getCell(col))) {
				return String.valueOf(sheet.getRow(row).getCell(col).getDateCellValue());
			} else if (sheet.getRow(row).getCell(col).getCellType() == 0) {
				return String.valueOf(sheet.getRow(row).getCell(col).getNumericCellValue());
			} else if (sheet.getRow(row).getCell(col).getCellType() == 2) {
				return String.valueOf(evaluator.evaluateInCell(sheet.getRow(row).getCell(col)));
			} else {
				return "A new Cell type which you havent handled";
			}
		} catch (NullPointerException e) {
			return null;
		}
	}
	
	/*
	 * Get any row and column from any sheet from any workbook
	 * 
	 * @params File name
	 */
	public String GetCellDataFromSheet(int row, int col, String FilePath , String SheetName) throws Exception {

		try {
			// System.out.println("File Opened Sucessfully");
			inputStream = new FileInputStream(FilePath);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
			throw new Exception("File Not Found please check the name");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(SheetName);

		// System.out.println(sheet.getRow(row).getCell(col).getCellType());
		// Returns the cell value
		try {
			if (sheet.getRow(row).getCell(col).getCellType() == 1) {
				return sheet.getRow(row).getCell(col).getStringCellValue();
			} else if (HSSFDateUtil.isCellDateFormatted(sheet.getRow(row).getCell(col))) {
				return String.valueOf(sheet.getRow(row).getCell(col).getDateCellValue());
			} else if (sheet.getRow(row).getCell(col).getCellType() == 0) {
				return String.valueOf(sheet.getRow(row).getCell(col).getNumericCellValue());
			} else if (sheet.getRow(row).getCell(col).getCellType() == 2) {
				return String.valueOf(evaluator.evaluateInCell(sheet.getRow(row).getCell(col)));
			} else {
				return "A new Cell type which you havent handled";
			}
		} catch (NullPointerException e) {
			return null;
		}
	}
	
	

	/*
	 * @Returns the number of data present in the source sheet
	 */
	public int GetRowCount(String FilePath, String SheetName) throws Exception {

		int rowCount = 0;

		try {
			file = new File(filemodels.getCompanyNamesSourcePath());
			inputStream = new FileInputStream(FilePath);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
			throw new Exception("File Not Found please check the name");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);

		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(SheetName);

		rowCount = sheet.getLastRowNum() + 1;

		return rowCount;
	}

	/*
	 * Gets the Column count
	 */
	public int GetColCount(String FilePath, String sheetName, int row) throws Exception {

		int colCount = 0;

		try {
			file = new File(filemodels.getCompanyNamesSourcePath());
			inputStream = new FileInputStream(FilePath);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
			throw new Exception("File Not Found please check the name");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);

		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(sheetName);

		colCount = sheet.getRow(row).getPhysicalNumberOfCells();

		return colCount;
	}
	
	/*
	 * Returns the current System time
	 */
	public String GetDate() {
		 
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");  
		LocalDateTime now = LocalDateTime.now();   
		return dtf.format(now); 
	}
	
	/*
	 * Prints the data from the dump data from the sheet
	 * 
	 * @params File path of the Dump file
	 */
	public void GetDumpData(String DumpFile, String CompanyName ,String ExactCompanyName) throws Exception {
		System.out.println("The Dump Filepath = " + DumpFile);

//		Boolean CompanyNameFalg = false, VersionFlag = false, MetaData = false, ProfitandLossFlag = false,
//				QuartersFlag = false, BalanceSheetFlag = false, CashFlowFlag = false;

		// Retrieve Company Name since its common for all can use the row and column
		// details
		System.out.println("-----------------------------");
		System.out.println("PARSED DATA FROM EXCEL DUMPS");
		System.out.println("-----------------------------");

		System.out.println("Company Name from passed sheet : " + GetCellData(0, 1, DumpFile));
		System.out.println("Company Name from EQUITY Sheet : " + ExactCompanyName);
		// Asserting is the company names match
		if (GetCellData(0, 1, DumpFile).toLowerCase().startsWith(CompanyName.toLowerCase())) {
			// System.out.println("Can Proceed parsing");
			System.out.println("Version : " + GetCellData(2, 1, DumpFile));

			// Getting Meta Data
			// -------------------
			System.out.println(GetCellData(4, 0, DumpFile));
			HashMap<String, String> meta = new HashMap<String, String>();
			int metaVar = 5;
			for (int i = 0; i < 4; i++) {
				meta.put(GetCellData(metaVar, 0, DumpFile), GetCellData(metaVar, 1, DumpFile));
				metaVar++;
			}
			// Printing Meta Data
			System.out.println(meta);

			// Getting Profit and loss
			// -------------------------
			System.out.println(GetCellData(14, 0, DumpFile));
			// HashMap<String, String> ProfitLoss = new HashMap<String, String>();

			int ProfitLossCounter = 16;
			int ProfitLossCounterIN = 16;
			int ProfitLossDate = 15;
			for (int PL = 0; PL < 15; PL++) {
				System.out.println("----------------------------");
				System.out.println(GetCellData(ProfitLossCounter, 0, DumpFile));
				System.out.println("----------------------------");
				for (int colCount = 1; colCount < GetColCount(DumpFile, filemodels.getDataSheetname(),
						ProfitLossDate); colCount++) {
//					ProfitLoss.put(GetCellData(ProfitLossDate, colCount, DumpFile),
//							GetCellData(ProfitLossCounterIN, colCount, DumpFile));
					System.out.println("Date :" + GetCellData(ProfitLossDate, colCount, DumpFile) + "  ||  " + "Data :"
							+ GetCellData(ProfitLossCounterIN, colCount, DumpFile));
				}
				ProfitLossCounter++;
				ProfitLossCounterIN++;
			}
			System.out.println("**************************************************************************");

			// Getting Quarters
			// --------------------
			System.out.println(GetCellData(39, 0, DumpFile));
		//	HashMap<String, String> Quarters = new HashMap<String, String>();

			int QuarterCounter = 41;
			int QuarterCounterIN = 41;
			int QuarterDate = 40;
			for (int Q = 0; Q < 9; Q++) {
				System.out.println("----------------------------");
				System.out.println(GetCellData(QuarterCounter, 0, DumpFile));
				System.out.println("----------------------------");
				for (int colCountQ = 1; colCountQ < GetColCount(DumpFile, filemodels.getDataSheetname(),
						QuarterDate); colCountQ++) {
					System.out.println("Date :" + GetCellData(QuarterDate, colCountQ, DumpFile) + "  ||  " + "Data :"
							+ GetCellData(QuarterCounterIN, colCountQ, DumpFile));
				}
				QuarterCounter++;
				QuarterCounterIN++;
			}

			System.out.println("**************************************************************************");

			// Balance Sheet
			// ----------------
			System.out.println(GetCellData(54, 0, DumpFile));
			// HashMap<String, String> Quarters = new HashMap<String, String>();

			int BalanceCounter = 56;
			int BalanceCounterIN = 56;
			int BalanceDate = 55;
			for (int B = 0; B < 16; B++) {
				System.out.println("----------------------------");
				System.out.println(GetCellData(BalanceCounter, 0, DumpFile));
				System.out.println("----------------------------");
				if (!GetCellData(BalanceCounter, 0, DumpFile).equalsIgnoreCase("Total")) {
					for (int colCountB = 1; colCountB < GetColCount(DumpFile, filemodels.getDataSheetname(),
							BalanceDate); colCountB++) {
						System.out.println("Date :" + GetCellData(BalanceDate, colCountB, DumpFile) + "  ||  "
								+ "Data :" + GetCellData(BalanceCounterIN, colCountB, DumpFile));
					}
					BalanceCounter++;
					BalanceCounterIN++;
				}else {
					System.out.println("Skipping Total Field");
					BalanceCounter++;
					BalanceCounterIN++;
				}

			}

			System.out.println("**************************************************************************");
			
			
			// CASH FLOW
			//---------------
			System.out.println(GetCellData(79, 0, DumpFile));
		//	HashMap<String, String> Quarters = new HashMap<String, String>();

			int CashCounter = 81;
			int CashCounterIN = 81;
			int CashDate = 80;
			for (int c = 0; c < 4; c++) {
				System.out.println("----------------------------");
				System.out.println(GetCellData(CashCounter, 0, DumpFile));
				System.out.println("----------------------------");
				for (int colCountC = 1; colCountC < GetColCount(DumpFile, filemodels.getDataSheetname(),
						CashDate); colCountC++) {
					System.out.println("Date :" + GetCellData(CashDate, colCountC, DumpFile) + "  ||  " + "Data :"
							+ GetCellData(CashCounterIN, colCountC, DumpFile));
				}
				CashCounter++;
				CashCounterIN++;
			}

			System.out.println("**************************************************************************");

			
			
			// Price
			//------------
			System.out.println("----------------------------");
			System.out.println("Price");
			System.out.println("----------------------------");
			for (int price = 1; price <GetColCount(DumpFile, filemodels.getDataSheetname(),
					89) ; price++) {
				System.out.println("Date : "+GetCellData(80, price, DumpFile)+" || Data :"+GetCellData(89, price, DumpFile));
			}
			System.out.println("**************************************************************************");
			
			//Derived
			//-----------
			System.out.println("----------------------------");
			System.out.println(GetCellData(91, 0, DumpFile));
			System.out.println("----------------------------");
			for (int Derived = 1; Derived <GetColCount(DumpFile, filemodels.getDataSheetname(),
					92) ; Derived++) {
				System.out.println( "Date :"+ GetCellData(80, Derived, DumpFile)+" || Data : "+GetCellData(92, Derived, DumpFile));
			}
			System.out.println("**************************************************************************");
			System.out.println("++++++++++++++++++++++++++Parsing Completed++++++++++++++++++++++++++++++++");
			
			System.out.println("***************************************************************************");
			

		} else {
			System.out.println("Company name doesnt Match Aborting pasing for : " + CompanyName);
		}

	}

}
