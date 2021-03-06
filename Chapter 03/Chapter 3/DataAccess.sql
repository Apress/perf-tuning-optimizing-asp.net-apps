if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustOrderHistXML]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CustOrderHistXML]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteLastProduct]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteLastProduct]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetCustomerCount]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetCustomerCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetCustomerDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetCustomerDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetCustomerDetailsXML]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetCustomerDetailsXML]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetEmployeeList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetEmployeeList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetProductCount]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetProductCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetProductCountRetParam]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetProductCountRetParam]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetProductInventory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetProductInventory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetProductInventoryRow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetProductInventoryRow]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetProductInventoryXML]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetProductInventoryXML]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[InsertCustomerDetailsXML]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[InsertCustomerDetailsXML]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpdateCustomerDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UpdateCustomerDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpdateInventory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UpdateInventory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpdateOrderDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UpdateOrderDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustomerList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CustomerList]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE CustOrderHistXML @CustomerID nchar(5)
AS
SELECT ProductName, Total=SUM(Quantity)
FROM Products P, [Order Details] OD, Orders O, Customers C
WHERE C.CustomerID = @CustomerID
AND C.CustomerID = O.CustomerID AND O.OrderID = OD.OrderID AND OD.ProductID = P.ProductID
GROUP BY ProductName FOR XML RAW


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE DeleteLastProduct 
AS

DELETE FROM Products 
WHERE ProductId = (SELECT MAX(ProductID) FROM Products)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE  PROCEDURE GetCustomerCount
AS

SELECT COUNT(*) FROM Customers





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE  PROCEDURE GetCustomerDetails
@CustomerID nchar(5)
AS

IF (@CustomerID = '')
BEGIN
	SELECT * FROM Customers
END
ELSE
BEGIN
	SELECT * FROM Customers where CustomerID = @CustomerID
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE  PROCEDURE GetCustomerDetailsXML
@CustomerID nchar(5)
AS

IF (@CustomerID = '')
BEGIN
	SELECT * FROM Customers FOR XML AUTO
END
ELSE
BEGIN
	SELECT * FROM Customers where CustomerID = @CustomerID FOR XML AUTO
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO













CREATE       PROCEDURE GetEmployeeList
@EmployeeID as int,
@UserType as char(1)

AS

IF (@UserType = 'A')
BEGIN
	SELECT EmployeeID, FirstName, LastName, Title 
	FROM Employees 
END
ELSE
BEGIN
	SELECT EmployeeID, FirstName, LastName, Title 
	FROM Employees 
	WHERE ReportsTo = @EmployeeID OR EmployeeID = @EmployeeID
END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE GetProductCount 
AS

SELECT MAX(ProductID) FROM Products




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE   PROCEDURE GetProductCountRetParam
@ProductCount Int OUTPUT
AS

SELECT @ProductCount = MAX(ProductID) FROM Products







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    PROCEDURE GetProductInventory 
@ProductID int,
@ProductName nvarchar(40) OUTPUT,
@UnitPrice money OUTPUT,
@UnitsInStock smallint OUTPUT
AS

SELECT @ProductName = ProductName, @UnitPrice = UnitPrice, @UnitsInStock = UnitsInStock 
FROM Products WHERE ProductID = @ProductID





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE    PROCEDURE GetProductInventoryRow
@ProductID int
AS

SELECT ProductName, UnitPrice, UnitsInStock 
FROM Products WHERE ProductID = @ProductID






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE    PROCEDURE GetProductInventoryXML
@ProductID int
AS

SELECT ProductName, UnitPrice, UnitsInStock 
FROM Products WHERE ProductID = @ProductID
FOR XML AUTO







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE  PROCEDURE InsertCustomerDetailsXML
@CustomersXMLDoc text

AS

DECLARE @hDoc int
EXEC sp_xml_preparedocument @hDoc OUTPUT, @CustomersXMLDoc

INSERT Customers 
SELECT * 
FROM OPENXML(@hDoc, N'/ROOT/Customers') 
     WITH Customers

EXEC sp_xml_removedocument @hDoc






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE PROCEDURE UpdateCustomerDetails
@CustomerID nchar(5),
@CompanyName nvarchar(40),
@ContactName nvarchar(30),
@ContactTitle nvarchar(30),
@Address nvarchar(60),
@City nvarchar(15),
@Region nvarchar(15),
@PCode nvarchar(10),
@Country nvarchar(15),
@Phone nvarchar(24),
@Fax nvarchar(24)

AS

UPDATE [Customers]
SET [CustomerID]=@CustomerID, [CompanyName]=@CompanyName, 
	[ContactName]=@ContactName, [ContactTitle]=@ContactTitle, 
	[Address]=@Address, [City]=@City, 
	[Region]=@Region, [PostalCode]=@PCode, 
	[Country]=@Country, [Phone]=@Phone, [Fax]=@Fax
WHERE CustomerID = @CustomerID





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE    PROCEDURE UpdateInventory 
@ProductID int,
@InvChange int
AS

UPDATE Products 
SET UnitsInStock = UnitsInStock - @InvChange 
WHERE ProductID = @ProductID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









CREATE       PROCEDURE UpdateOrderDetails 
@OrderID int,
@ProductID int,
@Quantity smallint
AS
UPDATE [Order Details] SET Quantity = @Quantity 
WHERE OrderID = @OrderID and ProductID = @ProductID

SELECT P.ProductID, ProductName,
    UnitPrice=ROUND(Od.UnitPrice, 2),
    Quantity,
    ExtendedPrice=ROUND(CONVERT(money, Quantity * Od.UnitPrice), 2),
    ShippedDate
FROM Products P, [Order Details] Od, Orders O
WHERE Od.ProductID = P.ProductID and Od.OrderID = @OrderID and O.OrderID = @OrderID










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE CustomerList
AS
SELECT CustomerID, CompanyName FROM Customers

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



