
-- This stored procedure is required by the Web Service:
-- WebService6A.wsNorthwind
-- Web Method: GetCustomerList()
CREATE PROCEDURE CustomerList
AS
SELECT CustomerID, CompanyName FROM Customers