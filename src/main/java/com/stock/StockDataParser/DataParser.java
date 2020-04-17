package com.stock.StockDataParser;

import java.lang.reflect.Array;
import java.util.Arrays;

import com.stock.Models.FileModels;
import com.stock.fileUtils.FileUtilityMethods;

public class DataParser {

	public static void main(String[] args) throws Exception {

		FileUtilityMethods Fileutils = new FileUtilityMethods();
		FileModels Filemodels = new FileModels();

		boolean Flag = false;

		// Main Data Parser class
		System.out.println("-------------------------------");
		System.out.println("INSIDE MAIN DATA PARSER CLASS");
		System.out.println("-------------------------------");

		// Number of data present in the sheet
		System.out.println("Number of records in Sheet : "
				+ Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(), Filemodels.getDefaultSheetName()));
		int numOfRecord = Fileutils.GetRowCount(Filemodels.getCompanyNamesSourcePath(),
				Filemodels.getDefaultSheetName());
		// Runs a loop for the row count to extract all data from the sheet

		for (int loop = 0; loop < numOfRecord; loop++) {

			// The dump File path
			String DumpFilePath = Filemodels.getStockDumpsSourcePath() + Fileutils.GetCompanyNameAt(loop, 0).trim()
					+ ".xlsx";
			System.out.println("Searching company : " + Fileutils.GetCompanyNameAt(loop, 0));
			
			
			// Fetches the first letter of the company to ease the process of searching
			String SearchingCompany = Fileutils.GetCompanyNameAt(loop, 0);
			String[] Split = SearchingCompany.split("");
			String CompanyNameLetter = (String)Array.get(Split, 0);
			System.out.println(CompanyNameLetter.toUpperCase());
			try {

				for (int j = 1; j < Fileutils.GetRowCount(Filemodels.getMainEquitySheet(),
						CompanyNameLetter.toUpperCase()); j++) {
					String NameFromMainNameSheet = Fileutils.GetCellDataFromSheet(j, 1, Filemodels.getMainEquitySheet(),
							CompanyNameLetter.toUpperCase());
					System.out.println(NameFromMainNameSheet.toLowerCase());
					System.out.println(Fileutils.GetCompanyNameAt(loop, 0).toLowerCase());
					if (NameFromMainNameSheet.toLowerCase()
							.contains(Fileutils.GetCompanyNameAt(loop, 0).toLowerCase().trim())) {

						// Can proceed company name exists
						Flag = true;
						Filemodels.setCurrentValidCompanyName(Fileutils.GetCellDataFromSheet(j, 1,
								Filemodels.getMainEquitySheet(),CompanyNameLetter.toUpperCase() ));
						System.out.println(" Company from Anu sheet : " + Fileutils.GetCompanyNameAt(1, 0));
						System.out.println(" Company from Kishore sheet : " + Filemodels.getCurrentValidCompanyName());
						System.out.println("Proceed Flag : " + Flag);
						break;
					}

				}

				if (Flag) {
					System.out.println("Inside parser method");
					System.out.println(
							"Current working company from Equity sheet : " + Filemodels.getCurrentValidCompanyName());

					// Sending the file path and company name to the parsing method
					Fileutils.GetDumpData(DumpFilePath, Fileutils.GetCompanyNameAt(loop, 0), Filemodels.getCurrentValidCompanyName().trim());
				} else {
					System.out.println("Company Name Not found in the Equity Sheet");
				}
			} catch (Exception E) {
				E.printStackTrace();
				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
				System.out.println("Check File name in the Company Name source File");
				System.out.println("Skipping Company : " + Fileutils.GetCompanyNameAt(loop, 0));
				System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");

			}

		}

	}
}
