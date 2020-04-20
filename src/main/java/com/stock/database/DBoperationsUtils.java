package com.stock.database;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

public class DBoperationsUtils {
	
	// Local DB
	private String URL = "jdbc:postgresql://localhost:5432/smartstockuserbase";
	private String Username = "smartstockadmin";
	private String Password = "stockadmin";
	
	//  Heroku DB
//	private String URL = "jdbc:postgresql://ec2-54-210-128-153.compute-1.amazonaws.com:5432/d5r0jlj0ums8g2";
//	private String Username = "vwqnlmodrvyxdn";
//	private String Password = "56b6704abcb019ae537e8832813ba99fe759261cd1c407b8dc653ed15df704ad";
	
	
	Connection connection = null;
	
	/*
	 *  Execute Insert Query
	 */

	public String ExecuteUpdateQuery(String Query) {
		
		try {
			connection = DriverManager.getConnection(URL, Username, Password);
			Statement statement = connection.createStatement();
			statement.executeUpdate(Query);
			return "Query Execution Sucessful";
		}catch (Exception e) {
			e.printStackTrace();
			return "Unable to Execute Query";
		}
	}
	
	/*
	 * Execute Query
	 */
	public ResultSet ExecuteQuery(String Query) {
		try {
			connection = DriverManager.getConnection(URL, Username, Password);
			Statement statement = connection.createStatement();
			
			return statement.executeQuery(Query);
		}catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	/*
	 * Returns String value of the required data from DB
	 */
	public String RetriveField(String TableName, String Field, String Constraint , String ConstraintValue) {
		
		String Query = "Select "+Field+" FROM "+TableName+" WHERE "+Constraint+" = "+ConstraintValue+";";
		System.out.println("Data Retieval Query : "+Query);
		try {
			connection = DriverManager.getConnection(URL, Username, Password);
			Statement statement = connection.createStatement();
			ResultSet rs = statement.executeQuery(Query); 
			rs.next();
			return rs.getString(1);
		}catch (Exception e) {
			return "Exception While retieving data";
		}

	}
}
