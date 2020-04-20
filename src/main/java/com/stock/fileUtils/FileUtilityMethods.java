package com.stock.fileUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.stock.Models.FileModels;
import com.stock.database.DBoperationsUtils;

public class FileUtilityMethods {

	// Instantiating FileModels
	FileModels filemodels = new FileModels();
	DBoperationsUtils DB = new DBoperationsUtils();

	// Getting the file to read the names of the company
	File file;
	FileInputStream inputStream;

	/*
	 * @Returns Company name
	 */
	public String GetCompanyNameAt(int row, int col) throws Exception {

		// Inside get company name
		file = new File(filemodels.getCompanyNamesSourcePath());
		// System.out.println("Reading Company names from file : " +
		// filemodels.getCompanyNamesSourcePath());

		try {
			inputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
		}

		// Getting the workbook
		@SuppressWarnings("resource")
		Workbook workbook = new XSSFWorkbook(inputStream);

		// Getting hold of the sheet
		Sheet sheet = workbook.getSheet(filemodels.getDefaultSheetName());
		try {
			// Returns the cell value
			return sheet.getRow(row).getCell(col).getStringCellValue();
		} catch (NullPointerException e) {
			System.out.println("Null Cell Value");
			return "Null";
		}
	}

	/*
	 * returns the crore value of the passed amount
	 * 
	 * @params File name
	 */
	public String GetMoneyValue(String Money) {

		// long money = (long) DecimalFormat.parse(Money);
		try {
			BigDecimal money = new BigDecimal(Money);
			BigDecimal modifier = new BigDecimal(10000000.00);
			money = money.multiply(modifier);
			String MoneyString = String.valueOf(money);
			return MoneyString;
		} catch (NullPointerException e) {

			return null;
		}
	}

	/*
	 * returns the Large number value of the passed amount
	 * 
	 * @params File name
	 */
	public String GetLargeValue(String Money) {

		// long money = (long) DecimalFormat.parse(Money);
		try {
			BigDecimal money = new BigDecimal(Money);
			String MoneyString = String.valueOf(money);
			return MoneyString;
		} catch (NullPointerException e) {

			return null;
		}
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
		}
		SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

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
				Date date = sheet.getRow(row).getCell(col).getDateCellValue();
				return String.valueOf(format.format(date));
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
	public String GetCellDataFromSheet(int row, int col, String FilePath, String SheetName) throws Exception {

		try {
			// System.out.println("File Opened Sucessfully");
			inputStream = new FileInputStream(FilePath);
		} catch (FileNotFoundException e) {
			System.out.println("File not Found Exiting");
		}
		SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

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
				Date date = sheet.getRow(row).getCell(col).getDateCellValue();
				return String.valueOf(format.format(date));
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
	 * Adding or subtracting dates
	 */
	public String ManipulateDate(String Date, String Operation, String Parameter, int parameterValue) {

		Calendar cal = Calendar.getInstance();
		SimpleDateFormat parser = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		// DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");

		try {
			Date date = parser.parse(Date);
			System.out.println(date);
			cal.setTime(date);
			if (Operation.equalsIgnoreCase("Add")) {

				switch (Parameter) {
				case "YEAR":
					cal.add(Calendar.YEAR, +parameterValue);
					return String.valueOf(cal.get(Calendar.YEAR) + "/0" + cal.get(Calendar.MONTH) + "/"
							+ cal.get(Calendar.DATE) + " " + cal.get(Calendar.HOUR) + "0:0" + cal.get(Calendar.MINUTE)
							+ ":0" + cal.get(Calendar.SECOND));
				case "MONTH":
					cal.add(Calendar.MONTH, +parameterValue);
					return String.valueOf(cal.get(Calendar.YEAR) + "/0" + cal.get(Calendar.MONTH) + "/"
							+ cal.get(Calendar.DATE) + " " + cal.get(Calendar.HOUR) + "0:0" + cal.get(Calendar.MINUTE)
							+ ":0" + cal.get(Calendar.SECOND));
				default:
					return "Adding date Failure";
				}

			} else if (Operation.equalsIgnoreCase("Minus")) {
				switch (Parameter) {
				case "YEAR":
					cal.add(Calendar.YEAR, -parameterValue);
					return String.valueOf(cal.get(Calendar.YEAR) + "/0" + cal.get(Calendar.MONTH) + "/"
							+ cal.get(Calendar.DATE) + " " + cal.get(Calendar.HOUR) + "0:" + cal.get(Calendar.MINUTE)
							+ "0:0" + cal.get(Calendar.SECOND));
				case "MONTH":
					cal.add(Calendar.MONTH, -parameterValue);
					return String.valueOf(cal.get(Calendar.YEAR) + "/0" + cal.get(Calendar.MONTH) + "/"
							+ cal.get(Calendar.DATE) + " " + cal.get(Calendar.HOUR) + "0:" + cal.get(Calendar.MINUTE)
							+ "0:0" + cal.get(Calendar.SECOND));
				default:
					return "Adding Date Failure";

				}
			} else {
				return "Check Date Manipulation parameters";
			}

		} catch (Exception e) {
			return "Adding to current date method failed";

		}

	}

	/*
	 * Get Values of date
	 */
	public String GetDateParam(String Param, String date) {
		try {
			if (Param.equalsIgnoreCase("date")) {
				SimpleDateFormat parser = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
				Date newdate = parser.parse(date);
				SimpleDateFormat formatter = new SimpleDateFormat("dd");
				String formattedDate = formatter.format(newdate);
				return formattedDate;

			} else if (Param.equalsIgnoreCase("year")) {
				SimpleDateFormat parser = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
				Date newdate = parser.parse(date);
				SimpleDateFormat formatter = new SimpleDateFormat("yyyy");
				String formattedDate = formatter.format(newdate);
				return formattedDate;
			} else if (Param.equalsIgnoreCase("month")) {
				SimpleDateFormat parser = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
				Date newdate = parser.parse(date);
				SimpleDateFormat formatter = new SimpleDateFormat("MM");
				String formattedDate = formatter.format(newdate);
				return formattedDate;
			} else {
				return "Undefined Value";
			}
		} catch (Exception e) {
			System.out.println("Date Is not in a perfect Date Format");
			return null;
		}
	}

	/*
	 * Return String conditioned for comparison
	 * 
	 */
	public String PrepareForCompare(String name) {
		try {
			// System.out.println("Replacing Limited");
			String regex = "\\s*\\bLimited\\b\\s*";
			name = name.replaceAll(regex, "");
		} catch (Exception e) {
			System.out.println("Exception when replacing Limited in name " + name);
		}

		try {
			// System.out.println("Replacing limited");
			String regex = "\\s*\\blimited\\b\\s*";
			name = name.replaceAll(regex, "");
		} catch (Exception e) {
			System.out.println("Exception when replacing limited in name " + name);
		}
		try {
			// System.out.println("Replacing LTD");
			String regex = "\\s*\\bLTD\\b\\s*";
			name = name.replaceAll(regex, "");
		} catch (Exception e) {
			System.out.println("Exception when replacing LTD in name " + name);
		}
		try {
			// System.out.println("Replacing LIMITED");
			String regex = "\\s*\\bLIMITED\\b\\s*";
			name = name.replaceAll(regex, "");
		} catch (Exception e) {
			System.out.println("Exception when replacing Limited in name " + name);
		}
		try {
			// System.out.println("Replacing And");
			// String regex = "\\s*\\bAnd\\b\\s*";
			name = name.replaceAll("And", "&");
		} catch (Exception e) {
			System.out.println("Exception when replacing Limited in name " + name);
		}
		return name.trim();
	}

	/*
	 * Prints the data from the dump data from the sheet
	 * 
	 * @params File path of the Dump file
	 */
	public void GetDumpData(String DumpFile, String CompanyName, String ExactCompanyName, String CompanyID)
			throws Exception {
		System.out.println("The Dump Filepath = " + DumpFile);

//		Boolean CompanyNameFalg = false, VersionFlag = false, MetaData = false, ProfitandLossFlag = false,
//				QuartersFlag = false, BalanceSheetFlag = false, CashFlowFlag = false;

		// Retrieve Company Name since its common for all can use the row and column
		// details

		boolean ProceedFalg = false;

		System.out.println("-----------------------------");
		System.out.println("PARSED DATA FROM EXCEL DUMPS");
		System.out.println("-----------------------------");

		System.out.println("Dump File Path : " + DumpFile);
		System.out.println("Company Name : " + CompanyName);
		System.out.println("Valid Company Name : " + ExactCompanyName);
		System.out.println("Company ID : " + CompanyID);
		String CompanyNameFromDB = DB.RetriveField("T_COMPANY_LIST", "\"COMPANY_NAME\"", "\"COMPANY_SYMBOL\"",
				"'" + CompanyID + "'");
		System.out.println("Company name retrieved from DB : " + CompanyNameFromDB);
		String CompanyNameFromDump = GetCellData(0, 1, DumpFile);
		System.out.println("Company Name from Dump File : " + CompanyNameFromDump);

		// Main Assertion Code To check if the company names match
		// assign Proceed Falg to true to proceed forward

		CompanyNameFromDB = PrepareForCompare(CompanyNameFromDB).toLowerCase();
		CompanyNameFromDump = PrepareForCompare(CompanyNameFromDump).toLowerCase();

		System.out.println(CompanyNameFromDB);
		System.out.println(CompanyNameFromDump);

		if (CompanyNameFromDump.contains(CompanyNameFromDB)) {
			ProceedFalg = true;
			System.out.println("Proceed Flag : " + ProceedFalg);
		} else {
			System.out.println("Proceed Flag : " + ProceedFalg);
		}

//		Data Which needs to be retireved for each cycle
//		-----------------------------------------
//		T_YEARLY_DETAIL
//		------------------------------------------
//		PK_YEARLY_DETAIL = primary key serial
//		FK_COMPANY_LIST = Serial from T_COMPANY_LIST
//		FK_OMPANY_SYMBOL = From DB FROM TABLE  T_COMPANY_LIST
//		REPORT YEAR = From File Which Years report
//		Sales = From File
//		Raw_MATerial cost = From File
//		CHANGE_IN_INVENTORY 	money 	
//		POWER_AND_FUEL 	money 	
//		OTHER_MANUFACTURING_EXPENSES 	money 	
//		EMPLOYEE_COST 	money 	
//		SELLING_AND_ADMIN 	money 	
//		OTHER_EXPENSES 	money 	
//		OTHER_INCOME 	money 	
//		DEPRECIATION 	money 	
//		INTEREST 	money 	
//		PROFIT_BEFORE_TAX 	money 	
//		TAX 	money 	
//		NET_PROFIT 	money 	
//		DIVIDEND_AMOUNT 	money 	

		// Asserting is the company names match
		if (ProceedFalg) {
			// System.out.println("Can Proceed parsing");
			System.out.println("Company Name : " + GetCellData(0, 1, DumpFile));

			// ------------------------------------------------------------
			// Logic to update generate update query for T_YEARLY_DETAIL
			// ------------------------------------------------------------
			System.out.println("-----------------------------------------------------------");
			System.out.println("Logic to update generate update query for T_YEARLY_DETAIL");
			System.out.println("------------------------------------------------------------");
			// select "PK_COMPANY_LIST" from T_COMPANY_LIST where "COMPANY_SYMBOL" =
			// '20MICRONS';
			String FK_COMPANY_LIST = DB.RetriveField("T_COMPANY_LIST", "\"PK_COMPANY_LIST\"", "\"COMPANY_SYMBOL\"",
					"'" + CompanyID + "'");
			System.out.println("FK_COMPANY_LIST : " + FK_COMPANY_LIST);

			// Company Symbol or ID
			System.out.println("FK_COMPANY_SYMBOL : " + CompanyID);

			// For runs till the end of column of date
			for (int i = 1; i < GetColCount(DumpFile, filemodels.getDataSheetname(), 15); i++) {

				// Report Year - Iterate column
				String REPORT_YEAR_date = GetCellDataFromSheet(15, i, DumpFile, filemodels.getDataSheetname());

				if (REPORT_YEAR_date != null) {

					System.out.println("----------------------------------------------------");
					System.out.println("REport Year : "+GetDateParam("YEAR", REPORT_YEAR_date));
					System.out.println("-----------------------------------------------------");
					// Sales
					String Sales = GetMoneyValue(GetCellDataFromSheet(16, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("Sales for year : " + Sales);

					// Raw Material cost
					String RAWMATERIAL_COST = GetMoneyValue(
							GetCellDataFromSheet(17, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("RAWMATERIAL_COST for year : " + RAWMATERIAL_COST);

					// Change in inventory
					String CHANGE_IN_INVENTORY = GetMoneyValue(
							GetCellDataFromSheet(18, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("CHANGE_IN_INVENTORY for year : " + CHANGE_IN_INVENTORY);

					// POWER_FUEL
					String POWER_FUEL = GetMoneyValue(
							GetCellDataFromSheet(19, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("POWER_FUEL for year : " + POWER_FUEL);

					// OTHER_MFG
					String OTHER_MFG = GetMoneyValue(
							GetCellDataFromSheet(20, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("OTHER_MFG for year : " + OTHER_MFG);

					// EMPLOYEE_COST
					String EMPLOYEE_COST = GetMoneyValue(
							GetCellDataFromSheet(21, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("EMPLOYEE_COST for year : " + EMPLOYEE_COST);

					// SELLING_ADMIN
					String SELLING_ADMIN = GetMoneyValue(
							GetCellDataFromSheet(22, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("SELLING_ADMIN for year : " + SELLING_ADMIN);

					// OTHER_EXPENSES
					String OTHER_EXPENSES = GetMoneyValue(
							GetCellDataFromSheet(23, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("OTHER_EXPENSES for year : " + OTHER_EXPENSES);

					// OTHER_INCOME
					String OTHER_INCOME = GetMoneyValue(
							GetCellDataFromSheet(24, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("OTHER_INCOME for year : " + OTHER_INCOME);

					// DEPRECATION
					String DEPRECATION = GetMoneyValue(
							GetCellDataFromSheet(25, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("DEPRECATION for year : " + DEPRECATION);

					// INTEREST
					String INTEREST = GetMoneyValue(
							GetCellDataFromSheet(26, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("INTEREST for year : " + INTEREST);

					// PROFIT_BEFORE_TAX
					String PROFIT_BEFORE_TAX = GetMoneyValue(
							GetCellDataFromSheet(27, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("PROFIT_BEFORE_TAX for year : " + PROFIT_BEFORE_TAX);

					// TAX
					String TAX = GetMoneyValue(GetCellDataFromSheet(28, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("TAX for year : " + TAX);

					// NET_PROFIT
					String NET_PROFIT = GetMoneyValue(
							GetCellDataFromSheet(29, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("NET_PROFIT for year : " + NET_PROFIT);

					// DIVIDEND_AMOUNT
					String DIVIDEND_AMOUNT = GetMoneyValue(
							GetCellDataFromSheet(30, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("DIVIDEND_AMOUNT for year : " + DIVIDEND_AMOUNT);

//    		EQUITY_SHARE_CAPITAL 	money 	
//    		RESERVES 	money 	
//    		BORROWINGS 	money 	
//    		OTHER_LIABILITIES 	money 	
//    		BALANCE_TOTAL 	money 	
//    		NET_BLOCK 	money 	
//    		CAPITAL_WORK_INPROGRESS 	money 	
//    		INVESTMENTS 	money 	
//    		OTHER_ASSETS 	money 	
//    		BALANCE_TOTAL2 	money 	
//    		RECEIVABLES 	money 	
//    		INVENTORY 	money 	
//    		CASH_AND_BANK 	money 	
//    		NO_OF_EQUITY_SHARES 	bigserial(25) 	
//    		NO_OF_NEW_EQUITY_SHARES 	bigserial(25) 	
//    		SPLIT_VALUE 	decimal(5) 	
//    		FACE_VALUE 	money 	

					// Balance sheet Details
					System.out.println("-----------------------");
					System.out.println("BALANCE SHEET TABLE");
					System.out.println("-----------------------");

					String EQUITY_SHARE_CAPITAL = GetMoneyValue(
							GetCellDataFromSheet(56, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("EQUITY_SHARE_CAPITAL for year : " + EQUITY_SHARE_CAPITAL);

					String RESERVES = GetMoneyValue(
							GetCellDataFromSheet(57, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("RESERVES for year : " + RESERVES);

					String BORROWINGS = GetMoneyValue(
							GetCellDataFromSheet(58, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("BORROWINGS for year : " + BORROWINGS);

					String OTHER_LIABILITIES = GetMoneyValue(
							GetCellDataFromSheet(59, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("OTHER_LIABILITIES for year : " + OTHER_LIABILITIES);

					String TOTAL = GetMoneyValue(GetCellDataFromSheet(60, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("TOTAL for year : " + TOTAL);

					String NET_BLOCK = GetMoneyValue(
							GetCellDataFromSheet(61, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("NET_BLOCK for year : " + NET_BLOCK);

					String CAPITAL_WORK = GetMoneyValue(
							GetCellDataFromSheet(62, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("CAPITAL_WORK for year : " + CAPITAL_WORK);

					String INVESTMENTS = GetMoneyValue(
							GetCellDataFromSheet(63, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("INVESTMENTS for year : " + INVESTMENTS);

					String OTHER_ASSETS = GetMoneyValue(
							GetCellDataFromSheet(64, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("OTHER_ASSETS for year : " + OTHER_ASSETS);

					String TOTAL_SECOND = GetMoneyValue(
							GetCellDataFromSheet(65, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("TOTAL_SECOND for year : " + TOTAL_SECOND);

					String RECIEVABLES = GetMoneyValue(
							GetCellDataFromSheet(66, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("RECIEVABLES for year : " + RECIEVABLES);

					String INVENTORY = GetMoneyValue(
							GetCellDataFromSheet(67, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("INVENTORY for year : " + INVENTORY);

					String CASH_BANK = GetMoneyValue(
							GetCellDataFromSheet(68, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("CASH_BANK for year : " + CASH_BANK);

					String NO_OF_EQUITY_SHARES = GetLargeValue(
							GetCellDataFromSheet(69, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("NO_OF_EQUITY_SHARES for year : " + NO_OF_EQUITY_SHARES);

					String NEW_BONUS_SHARES = GetMoneyValue(
							GetCellDataFromSheet(70, i, DumpFile, filemodels.getDataSheetname()));
					System.out.println("NEW_BONUS_SHARES for year : " + NEW_BONUS_SHARES);

					String FACE_VALUE = GetCellDataFromSheet(71, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("FACE_VALUE for year : " + FACE_VALUE);

					// Cash Value Table
					System.out.println("-----------------------");
					System.out.println("CASH VALUE TABLE");
					System.out.println("-----------------------");

//		    		CASH_FROM_OPERATING_ACTIVITY 	money 	
//		    		CASH_FROM_INVESTING_ACTIVITY 	money 	
//		    		CASH_FROM_FINANCING_ACTIVITY 	money 	
//		    		NET_CASH_FLOW 	money 	
//		    		PRICE 	money 	
//		    		REPORT_START_DATE 	date 	
//		    		REPORT_END_DATE 	date 	
//		    		INSERT_TIME 	timestamp 	
//		    		UPDATE_TIME

					String OPERATING_ACTIVITY = GetCellDataFromSheet(81, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("OPERATING_ACTIVITY for year : " + OPERATING_ACTIVITY);

					String INVESTING_ACTIVITY = GetCellDataFromSheet(82, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("INVESTING_ACTIVITY for year : " + INVESTING_ACTIVITY);

					String FINANCING_ACTIVITY = GetCellDataFromSheet(83, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("FINANCING_ACTIVITY for year : " + FINANCING_ACTIVITY);

					String NET_CASH_FLOW = GetCellDataFromSheet(84, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("NET_CASH_FLOW for year : " + NET_CASH_FLOW);

					String PRICE = GetCellDataFromSheet(89, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("PRICE for year : " + PRICE);

					String START_DATE = ManipulateDate(
							GetCellDataFromSheet(15, i, DumpFile, filemodels.getDataSheetname()), "MINUS", "YEAR", 1);
					System.out.println("START_DATE for year : " + START_DATE);
					String END_DATE = GetCellDataFromSheet(15, i, DumpFile, filemodels.getDataSheetname());
					System.out.println("END_DATE for year : " + END_DATE);

					String INSERT_TIME = GetDate();
					System.out.println("INSERT_TIME for year : " + INSERT_TIME);
					String UPDATE_TIME = GetDate();
					System.out.println("UPDATE_TIME for year : " + UPDATE_TIME);

				} else {
					System.out.println("##################################");
					System.out.println("Report Date Null T_YEARLY_DETAIL");
					System.out.println("##################################");
				}

			}
			
			// ------------------------------------------------------------
			// Logic to update generate update query for T_QUARTERLY_DETAIL
			// ------------------------------------------------------------

//			PK_QUATER_DETAILS 	bigserial(25) 	
//			FK_COMPANY_LIST 	bigint 	
//			FK_COMPANY_SYMBOL 	varchar(50) 	
//			REPORT_YEAR 	int(4) 	
//			REPORT_MONTH 	int(4) 	
//			SALES 	money 	
//			EXPENSES 	money 	
//			OTHER_INCOME 	money 	
//			DEPRECIATION 	money 	
//			INTEREST 	money 	
//			PROFIT_BEFORE_TAX 	money 	
//			TAX 	money 	
//			NET_PROFIT 	money 	
//			OPERATING_PROFIT 	money 	
//			REPORT_START_DATE 	date 	
//			REPORT_END_DATE 	date 	
//			INSERT_TIME 	timestamp 	
//			UPDATE_TIME 	timestamp 			
			
			System.out.println("-----------------------------------------------------------");
			System.out.println("Logic to update generate update query for T_QUARTER_DETAIL");
			System.out.println("------------------------------------------------------------");
			
			// System.out.println("Can Proceed parsing");
			System.out.println("Company Name : " + GetCellData(0, 1, DumpFile));
			// select "PK_COMPANY_LIST" from T_COMPANY_LIST where "COMPANY_SYMBOL" =
			// '20MICRONS';
			String FK_COMPANY_LIST_QUARTER = DB.RetriveField("T_COMPANY_LIST", "\"PK_COMPANY_LIST\"", "\"COMPANY_SYMBOL\"",
					"'" + CompanyID + "'");
			System.out.println("FK_COMPANY_LIST_QUARTER : " + FK_COMPANY_LIST_QUARTER);

			// Company Symbol or ID
			System.out.println("FK_COMPANY_SYMBOL : " + CompanyID);

			for (int j = 1; j < GetColCount(DumpFile, filemodels.getDataSheetname(),40);j++) {
				
				// Report Year - Iterate column
				String REPORT_YEAR_date_QUARTER = GetCellDataFromSheet(40, j, DumpFile, filemodels.getDataSheetname());


				if (REPORT_YEAR_date_QUARTER != null) {
					System.out.println("----------------------------------------------------");
					System.out.println("REport Year : "+GetDateParam("YEAR", REPORT_YEAR_date_QUARTER));
					System.out.println("-----------------------------------------------------");
				
					String REPORT_YEAR_QUARTER = GetDateParam("YEAR", REPORT_YEAR_date_QUARTER);
					
					// Company List
				System.out.println("FK_COMPANY_LIST : " + FK_COMPANY_LIST);

				// Company Symbol
				System.out.println("FK_COMPANY_SYMBOL : " + CompanyID);

				// Report year
				System.out.println(REPORT_YEAR_QUARTER);

				// Report Month
				String REPORT_MONTH = GetDateParam("MONTH", REPORT_YEAR_date_QUARTER);
				System.out.println("REPORT_MONTH : " + REPORT_MONTH);

				// SALES
				String QUARTER_SALES = GetMoneyValue(
						GetCellDataFromSheet(41, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_SALES : "+QUARTER_SALES);
				
				// QUARTER_EXPENSES
				String QUARTER_EXPENSES = GetMoneyValue(
						GetCellDataFromSheet(42, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_EXPENSES : "+QUARTER_EXPENSES);
				
				// QUARTER_OTHER_INCOME
				String QUARTER_OTHER_INCOME = GetMoneyValue(
						GetCellDataFromSheet(43, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_OTHER_INCOME : "+QUARTER_OTHER_INCOME);
				
				// DEPRICATION
				String QUARTER_DEPRICATION = GetMoneyValue(
						GetCellDataFromSheet(44, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_DEPRICATION : "+QUARTER_DEPRICATION);
				
				// INTEREST
				String QUARTER_INTEREST = GetMoneyValue(
						GetCellDataFromSheet(45, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_INTEREST : "+QUARTER_INTEREST);
				
				// PROFIT BEFORE TAX
				String QUARTER_P_BEFORE_TAX = GetMoneyValue(
						GetCellDataFromSheet(46, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_P_BEFORE_TAX : "+QUARTER_P_BEFORE_TAX);
				
				//  TAX
				String QUARTER_TAX = GetMoneyValue(
						GetCellDataFromSheet(47, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_TAX : "+QUARTER_TAX);
				
				// QUARTER_NET_PROFIT
				String QUARTER_NET_PROFIT = GetMoneyValue(
						GetCellDataFromSheet(48, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_NET_PROFIT : "+QUARTER_NET_PROFIT);
				
				// OPERATING_PROFIT
				String QUARTER_OPERATING_PROFIT = GetMoneyValue(
						GetCellDataFromSheet(49, j, DumpFile, filemodels.getDataSheetname()));
				System.out.println("QUARTER_OPERATING_PROFIT : "+QUARTER_OPERATING_PROFIT);
				}else {
					System.out.println("##################################");
					System.out.println("Report Date Null T_QUARTER_DETAIL");
					System.out.println("##################################");
				}
			}

			System.out.println("***************************************************************************");
			System.out.println("++++++++++++++++++++++++++Parsing Completed++++++++++++++++++++++++++++++++");
			System.out.println("***************************************************************************");

		} else {
			System.out.println("**********************************************************************");
			System.out.println("Company name doesnt Match Aborting parsing for : " + CompanyName);
			System.out.println("Update Company Name in the company source file");
			System.out.println("**********************************************************************");
		}

	}

}
