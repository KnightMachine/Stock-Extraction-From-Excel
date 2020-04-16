package com.stock.StockDataParser;

import com.stock.Models.FileModels;
import com.stock.fileUtils.FileUtilityMethods;

public class DataParser {

		
		public static void main( String[] args ) throws Exception {

			FileUtilityMethods Fileutils = new FileUtilityMethods();
			FileModels Filemodels = new FileModels();

			// Main Data Parser class
			System.out.println("-------------------------------");
			System.out.println("INSIDE MAIN DATA PARSER CLASS");
			System.out.println("-------------------------------");
			
			// Number of data present in the sheet 
			System.out.println("Number of records in Sheet : "+Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(), Filemodels.getDefaultSheetName()));
			int numOfRecord = Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(), Filemodels.getDefaultSheetName());
			// Runs a loop for the row count to extract all data from the sheet 
			
			for(int loop = 0; loop<numOfRecord ; loop++) {
			
			// The dump File path
			String DumpFilePath = Filemodels.getStockDumpsSourcePath()+Fileutils.GetCompanyNameAt(loop, 0).trim()+".xlsx";
			
			try {
			//Sending the file path and company name to the parsing method
			Fileutils.GetDumpData(DumpFilePath,Fileutils.GetCompanyNameAt(loop, 0) );
			}catch(Exception E) {
				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
				System.out.println("Check File name in the Company Name source File");
				System.out.println("Skipping Company : "+Fileutils.GetCompanyNameAt(loop, 0));
				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
			}
			}
			
	}
		
	}
	

