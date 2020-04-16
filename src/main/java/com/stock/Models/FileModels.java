package com.stock.Models;

public class FileModels {
	
	public String getDataSheetname() {
		return DataSheetname;
	}

	public void setDataSheetname(String dataSheetname) {
		DataSheetname = dataSheetname;
	}

	public String getDefaultSheetName() {
		return DefaultSheetName;
	}

	public void setDefaultSheetName(String defaultSheetName) {
		DefaultSheetName = defaultSheetName;
	}

	public String getCompanyNamesSourcePath() {
		return CompanyNamesSourcePath;
	}

	public void setCompanyNamesSourcePath(String companyNamesSourcePath) {
		CompanyNamesSourcePath = companyNamesSourcePath;
	}

	public String getStockDumpsSourcePath() {
		return StockDumpsSourcePath;
	}

	public void setStockDumpsSourcePath(String stockDumpsSourcePath) {
		StockDumpsSourcePath = stockDumpsSourcePath;
	}

	String DefaultSheetName = "Sheet1";
	String CompanyNamesSourcePath = "SourceCompanyNames/SourceCompanyNames.xlsx";
	
	// Append company name with .xlsx extension
	String StockDumpsSourcePath = "StockDumpsSource/";
	
	// Sheet name in the Dump sheet
	String DataSheetname = "Data Sheet";
	

}
