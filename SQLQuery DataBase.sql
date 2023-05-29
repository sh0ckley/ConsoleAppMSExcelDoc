CREATE DATABASE dbStore
GO
USE dbStore
GO
CREATE TABLE Product 
(
	ID_Product		int				PRIMARY KEY	NOT NULL,
	Product_Name	nvarchar(max)				NULL,
	Price			int							NULL,
	Unit			nvarchar(max)				NULL,
)
GO
CREATE TABLE Clients 
(
	ID_Client	int				PRIMARY KEY	NOT NULL,
	Name_Org	nvarchar(max)				NULL,
	Address_Org	nvarchar(max)NULL,
	Surname		nvarchar(max)				NULL,
	Firstname	nvarchar(max)				NULL,
	Patronymic	nvarchar(max)				NULL
)
GO
CREATE TABLE Applications
(
	ID_Application		int				PRIMARY KEY									NOT NULL,
	ID_Product			int				FOREIGN KEY REFERENCES Product(ID_Product)	NOT NULL,
	ID_Client			int				FOREIGN KEY REFERENCES Clients(ID_Client)	NOT NULL,
	Application_Number	int															NULL,
	Required_Quantity	int															NULL,
	Date_Placement		nvarchar(10)												NULL		
)
GO
CREATE TABLE Storekeeper 
(
	ID_Storekeeper	int				PRIMARY KEY	NOT NULL,
	Surname			nvarchar(max)				NULL,
	Firstname		nvarchar(max)				NULL,
	Patronymic		nvarchar(max)				NULL
)
GO
CREATE TABLE Storage
(
	ID_Storage			int				PRIMARY KEY											NOT NULL,
	ID_Storekeeper		int				FOREIGN KEY REFERENCES Storekeeper (ID_Storekeeper)	NOT NULL,
	ID_Product			int				FOREIGN KEY REFERENCES Product(ID_Product)			NOT NULL,
	Product_Quantity	int																	NULL,
	Date_Supply			nvarchar(max)														NULL
)
GO
CREATE TABLE StorageStaff
(
	ID_Staff		int			PRIMARY KEY											NOT NULL,
	ID_Storekeeper	int			FOREIGN KEY REFERENCES Storekeeper (ID_Storekeeper) NOT NULL,
	Surname		nvarchar(max)														NULL,
	Firstname	nvarchar(max)														NULL,
	Patronymic	nvarchar(max)														NULL,
	Age			int																	NULL,
	Post		nvarchar(max)														NULL
)
GO

SELECT Product_Name, Price, Unit FROM Product
JOIN Applications ON Product.ID_Product = Applications.ID_Product
WHERE Price = '0'

UPDATE Product SET 
Price = Price * 1.05
WHERE Price = 300

select * from Product where price = 315

ALTER TABLE Applications
ADD Name nvarchar(max) NOT NULL 
DEFAULT 'Заявка №[Номер заявки] на приобретение [Наименование товара].'




