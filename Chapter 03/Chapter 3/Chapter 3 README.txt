README for Chapter 3 Sample Project


Project: AspNetChap3
---------------------

Purpose: Illustrates various data access methods using ADO.NET and ASP.NET. Illustrates the various methods of the Command, DataReader, 
	 DataAdapter, DataSet, and Typed DataSet objects.

	 NOTE, you must run the SQL Script: DataAccess.sql, in Query Analyzer, before running the AspNetChap3 project.
	 Please refer to the sample project installation instructions for more details.

					
Page					Purpose
------------------------------		------------------------------
Menu.aspx				Default page with links to each of the sample pages.

ExecNonQuery.aspx			Demonstrates how to use the ExecuteNonQuery method to call a stored procedure to update the inventory of a product.

ExecScalar.aspx				Demonstrates how to use the ExecuteNonQuery method to retrieve scalar data.

ExecSPOutParam.aspx			Demonstrates how to use Output Parameters collection to retrieve data.

ExecSPReturnDR.aspx			Demonstrates how to use a DataReader to retrieve different sets of data.
					Retrieves a scalar value, a row of data, and multiple rows of data.

ExecSPReturnDS.aspx			Demonstrates how to use a DataSet and a typed DataSet to retrieve differnet sets of data.
					Retrieves a row of data, and multiple rows of data.
					Output techniques include binding to a DataGrid, reading the elements programmatically, and using a DataView.

ExecSPReturnXML.aspx			Demonstrates how to use the XmlReader to retrieve multiple rows of data.

ExecSPReturnXSD.aspx			Demonstrates how the DataSet object interacts with XML and XSD.

UpdateDSWithDataGrid1.aspx		Demonstrates how to use the DataSet and DataAdapter objects to update a data 
					source. Uses the SQLCommandBuilder to autogenerate the update command.

UpdateDSWithDataGrid2.aspx		Demonstrates how to use the DataSet and DataAdapter objects to update a data 
					source. Uses a parameterized stored procedure for the update command.

UpdateWithParams.aspx			Demonstrates how to use the Parameters collection and the Command object to update a row of data.

ShowSQLExceptions.aspx			Demonstrates how to handle a SQL Exception.


Component				Purpose
------------------------------		------------------------------
dsCustOrderHist				XSD schema for the dsCustOrderHist typed DataSet.

DataAccess				Component that exposes generic and helper data access wrapper methods that are used throughout the sample project.
- ExecScalar				Uses the ExecuteScalar method to execute a stored procedure and returns a scalar value.
- ExecSPOutputParams			Uses the ExecuteNonQuery method to execute a stored procedure.
- ExecSPReturnDR			Uses the ExecuteReader method to execute a stored procedure and return a DataReader.
- ExecSPReturnDS			Uses the DataAdapter object to execute a stored procedure, fill and return a DataSet.
- ExecSPReturnDS1			Uses the DataAdapter object to execute a stored procedure, fill and return a typed DataSet.
- ExecSPReturnXML			Uses the ExecuteXmlReader method to execute a stored procedure and return a XmlReader.
- GetProductInventory			Returns the product name, unit price, and units in stock using the Parameters collection.
- SetupParameters			Private subroutine that populates the Command objects's Parameters collection based on the passed in string array.
- UpdateCustomerProfile			Updates a customer's profile using the the DataAdapter Update method and an assocaited DataSet.
- UpdateData				Private subroutine that updates the DataSet with the passed in string array.




