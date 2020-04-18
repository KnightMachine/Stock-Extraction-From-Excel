package com.stock.StockDataParser;

import java.lang.reflect.Array;
import java.sql.ResultSet;
import java.util.Arrays;

import com.stock.Models.FileModels;
import com.stock.database.DBoperationsUtils;
import com.stock.fileUtils.FileUtilityMethods;

public class DataParser {

	public static void main(String[] args) throws Exception {

		FileUtilityMethods Fileutils = new FileUtilityMethods();
		DBoperationsUtils DBoperation = new DBoperationsUtils();
		FileModels Filemodels = new FileModels();

		// Main Data Parser class
		System.out.println("-------------------------------");
		System.out.println("INSIDE MAIN DATA PARSER CLASS");
		System.out.println("-------------------------------");

		
		/*
		 * View Heroku Table 
		 */
		
//		System.out.println("Heroku table");
//		System.out.println(DBoperation.ExecuteQuery("select * from T_COMPANY_LIST; "));
//		ResultSet rs = DBoperation.ExecuteQuery("select * from T_COMPANY_LIST; ");
//		while (rs.next()) {
//			System.out.println(rs.getDouble(1)+"  ");
//			System.out.print(rs.getString(2)+"  ");
//			System.out.print(rs.getString(3)+"   ");
//		}
		
		
		
		/*
		 * ++++++++++++++++++++++++++++++++++++++++++++++++++++++ Logic to Udpdate
		 * Yearly Details and Quarterly Details
		 * +++++++++++++++++++++++++++++++++++++++++++++++++++++++
		 */

//		boolean Flag = false;
//		// Number of data present in the sheet
//		System.out.println("Number of records in Sheet : "
//				+ Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(), Filemodels.getDefaultSheetName()));
//		int numOfRecord = Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(),
//				Filemodels.getDefaultSheetName());
//		// Runs a loop for the row count to extract all data from the sheet
//
//		for (int loop = 0; loop < numOfRecord; loop++) {
//
//			// The dump File path
//			String DumpFilePath = Filemodels.getStockDumpsSourcePath() + Fileutils.GetCompanyNameAt(loop, 0).trim()
//					+ ".xlsx";
//			System.out.println("Searching company : " + Fileutils.GetCompanyNameAt(loop, 0));
//			
//			
//			// Fetches the first letter of the company to ease the process of searching
//			String SearchingCompany = Fileutils.GetCompanyNameAt(loop, 0);
//			String[] Split = SearchingCompany.split("");
//			String CompanyNameLetter = (String)Array.get(Split, 0);
//			System.out.println(CompanyNameLetter.toUpperCase());
//			
//			try {
//
//				for (int j = 1; j < Fileutils.GetRowCount(Filemodels.getMainEquitySheet(),
//						CompanyNameLetter.toUpperCase()); j++) {
//					String NameFromMainNameSheet = Fileutils.GetCellDataFromSheet(j, 1, Filemodels.getMainEquitySheet(),
//							CompanyNameLetter.toUpperCase());
//					System.out.println(NameFromMainNameSheet.toLowerCase());
//					System.out.println(Fileutils.GetCompanyNameAt(loop, 0).toLowerCase());
//					if (NameFromMainNameSheet.toLowerCase()
//							.contains(Fileutils.GetCompanyNameAt(loop, 0).toLowerCase().trim())) {
//
//						// Can proceed company name exists
//						Flag = true;
//						Filemodels.setCurrentValidCompanyName(Fileutils.GetCellDataFromSheet(j, 1,
//								Filemodels.getMainEquitySheet(),CompanyNameLetter.toUpperCase() ));
//						System.out.println(" Company from Anu sheet : " + Fileutils.GetCompanyNameAt(1, 0));
//						System.out.println(" Company from Kishore sheet : " + Filemodels.getCurrentValidCompanyName());
//						System.out.println("Proceed Flag : " + Flag);
//						break;
//					}
//
//				}
//
//				if (Flag) {
//					System.out.println("Inside parser method");
//					System.out.println(
//							"Current working company from Equity sheet : " + Filemodels.getCurrentValidCompanyName());
//
//					// Sending the file path and company name to the parsing method
//					Fileutils.GetDumpData(DumpFilePath, Fileutils.GetCompanyNameAt(loop, 0), Filemodels.getCurrentValidCompanyName().trim());
//				} else {
//					System.out.println("Company Name Not found in the Equity Sheet");
//				}
//			} catch (Exception E) {
//				E.printStackTrace();
//				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
//				System.out.println("Check File name in the Company Name source File");
//				System.out.println("Skipping Company : " + Fileutils.GetCompanyNameAt(loop, 0));
//				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
//
//			}
//	}

		/*
		 * +++++++++++++++++++++++++++++++++++++++
		 * Logic to Update Customer Details Table
		 * +++++++++++++++++++++++++++++++++++++++
		 * 
		 * Main technical Flaw is the File Names in the StockDumpSource Folder should be updated
		 * in the SourceCompanyNames Xlx file with the same name as the file to map the excel.
		 * Note : The File Names Should be same and case sensitive 
		 */
		System.out.println("-------------------------------");
		System.out.println("UPDATE COMPANY NAMES TABLE");
		System.out.println("-------------------------------");
		// Insert Query to insert into T_COMPANY_LIST
		// insert into T_COMPANY_LIST
		// ("COMPANY_SYMBOL","COMPANY_NAME","LISTING_DATE","MARKET_LOT","ISIN_NUMER",
		// "FACE_VALUE","MARKET_CAP","PREVIOUS_DAY_VALUE","NO_OF_SHARES","INSERT_TIME","UPDATE_TIME")
		// values('cmy symbol','name','2017-11-06
		// 12:11:58',123,'123456789',458,567,789,478,'2017/11/06 12:11:58',
		// '2017/11/06 12:11:58');

			for (int i = 1; i < Fileutils.GetRowCount(Filemodels.getMainEquitySheet(),
					Filemodels.getDefaultSheetName()); i++) {
				try {
				System.out.println("Updating Company Names Table for Company : " + Fileutils.GetCellDataFromSheet(i, 1,
						Filemodels.getMainEquitySheet(), Filemodels.getDefaultSheetName()));
				String DumpFilePath = Filemodels.getStockDumpsSourcePath() + Fileutils.GetCompanyNameAt(i-1, 0)
						.trim() + ".xlsx";
				System.out.println("DumpFile Path : " + DumpFilePath);
				
				
				// Company Names Table Fields
				String 
				
				        COMPANY_SYMBOL = Fileutils.GetCellDataFromSheet(i, 0, Filemodels.getMainEquitySheet(),
						        Filemodels.getDefaultSheetName()),
						COMPANY_NAME = Fileutils.GetCellDataFromSheet(i, 1, Filemodels.getMainEquitySheet(),
								Filemodels.getDefaultSheetName()),
						LISTING_DATE = Fileutils.GetCellDataFromSheet(i, 3, Filemodels.getMainEquitySheet(),
								Filemodels.getDefaultSheetName()),
						MARKET_LOT = Fileutils.GetCellDataFromSheet(i, 5, Filemodels.getMainEquitySheet(),
								Filemodels.getDefaultSheetName()),
						ISIN_NUMBER = Fileutils.GetCellDataFromSheet(i, 6, Filemodels.getMainEquitySheet(),
								Filemodels.getDefaultSheetName()),
						FACE_VALUE = Fileutils.GetCellDataFromSheet(i, 7, Filemodels.getMainEquitySheet(),
								Filemodels.getDefaultSheetName()),
						MARKET_CAP = Fileutils.GetCellDataFromSheet(8, 1, DumpFilePath, Filemodels.getDataSheetname()),
						PREVIOUS_DAY_VALUE = "0",
						NO_OF_SHARES = Fileutils.GetCellDataFromSheet(5, 1, DumpFilePath,
								Filemodels.getDataSheetname()),
						INSERT_TIME = Fileutils.GetDate(), UPDATE_TIME = Fileutils.GetDate();

				System.out.println(COMPANY_SYMBOL + " " + COMPANY_NAME + " " + LISTING_DATE + " " + MARKET_LOT + " "
						+ ISIN_NUMBER + " " + FACE_VALUE + " " + MARKET_CAP + " " + PREVIOUS_DAY_VALUE + " "
						+ NO_OF_SHARES + " " + INSERT_TIME + " " + UPDATE_TIME);
				
				// Use this update when using local DB
//				String UpdateQuery = "Insert into T_COMPANY_LIST (\"COMPANY_SYMBOL\",\"COMPANY_NAME\",\"LISTING_DATE\",\"MARKET_LOT\",\"ISIN_NUMER\",\"FACE_VALUE\","
//						+ "\"MARKET_CAP\",\"PREVIOUS_DAY_VALUE\",\"NO_OF_SHARES\",\"INSERT_TIME\",\"UPDATE_TIME\")values('"+COMPANY_SYMBOL+"','"+COMPANY_NAME+"','"+LISTING_DATE+"',"+MARKET_LOT+",'"+ISIN_NUMBER+"',"+FACE_VALUE+","+MARKET_CAP+","+PREVIOUS_DAY_VALUE+","+NO_OF_SHARES+",'"+INSERT_TIME+"','"+UPDATE_TIME+"');";
			
				
				// Use this update when using heroku
				String UpdateQuery = "Insert into T_COMPANY_LIST (COMPANY_SYMBOL,COMPANY_NAME,LISTING_DATE,MARKET_LOT,ISIN_NUMER,FACE_VALUE,"
						+ "MARKET_CAP,PREVIOUS_DAY_VALUE,NO_OF_SHARES,INSERT_TIME,UPDATE_TIME)values('"+COMPANY_SYMBOL+"','"+COMPANY_NAME+"','"+LISTING_DATE+"',"+MARKET_LOT+",'"+ISIN_NUMBER+"',"+FACE_VALUE+","+MARKET_CAP+","+PREVIOUS_DAY_VALUE+","+NO_OF_SHARES+",'"+INSERT_TIME+"','"+UPDATE_TIME+"');";
			
				
				System.out.println("------------------------------------------------------------");
				System.out.println("Update Queery For Company : "+COMPANY_NAME);
				System.out.println(UpdateQuery);
		        
		        // Create T_Company_List Table in Heroku 	
		        // String UpdateQuery = "CREATE TABLE T_COMPANY_LIST ( PK_COMPANY_LIST serial NOT NULL,COMPANY_SYMBOL varchar(50) NOT NULL UNIQUE,COMPANY_NAME varchar(300) NOT NULL,LISTING_DATE DATE NOT NULL,MARKET_LOT bigint NOT NULL,ISIN_NUMER varchar(15) NOT NULL,FACE_VALUE money NOT NULL,MARKET_CAP money NOT NULL,PREVIOUS_DAY_VALUE money NOT NULL,NO_OF_SHARES money NOT NULL,INSERT_TIME DATE NOT NULL,UPDATE_TIME DATE NOT NULL,CONSTRAINT T_COMPANY_LIST_pk PRIMARY KEY (PK_COMPANY_LIST)) WITH (OIDS=FALSE);";
				System.out.println(DBoperation.ExecuteUpdateQuery(UpdateQuery));
				System.out.println("------------------------------------------------------------");
				
				} catch (Exception e) {
			    System.out.println("----------------------------------------------");
				System.out.println("Exception Updating Company Names Table");
				System.out.println("-----------------------------------------------");
			}
		}

	}
}
