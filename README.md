# SQL Server + Excel Integration Project  
**Database: AdventureWorks DW 2014**

## Objective

Develop an integration between SQL Server and Excel to analyze online sales performance indicators using data from the AdventureWorks DW 2014 database. The focus is on performance metrics for the year 2013.

---

## Database Installation

To install the AdventureWorks DW 2014 database, follow Microsoft’s official instructions:  
🔗 [AdventureWorks - Installation and Configuration](https://docs.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms)

---

## Defined Metrics

1. **Total Internet Sales by Product Category**  
2. **Total Internet Revenue by Order Month**  
3. **Internet Revenue and Cost by Country**  
4. **Total Internet Sales by Customer Gender**

> **Analysis Year:** 2013

---

## Tables Used

| Table                  | Description                          |
|------------------------|--------------------------------------|
| `FactInternetSales`    | Internet sales facts                 |
| `DimCustomer`          | Customer demographic data            |
| `DimSalesTerritory`    | Geographic sales information         |
| `DimProductCategory`   | Product category                     |
| `DimProductSubcategory`| Product subcategory                  |
| `DimProduct`           | Product details                      |

---

## View Structure: `VENDAS_INTERNET`

#### CREATE VIEW VENDAS_INTERNET
```sql
CREATE VIEW VENDAS_INTERNET AS
SELECT
    fis.SalesOrderNumber AS [ORDER_NO],
    fis.OrderDate AS [ORDER_DATE],
    dpc.EnglishProductCategoryName AS [PRODUCT_CATEGORY],
    dc.FirstName + ' ' + dc.LastName AS [CUSTOMER_NAME],
    dc.Gender AS [GENDER],
    dst.SalesTerritoryCountry AS [COUNTRY],
    fis.OrderQuantity AS [QUANTITY_SOLD],
    fis.TotalProductCost AS [SALE_COST],
    fis.SalesAmount AS [SALE_REVENUE]
FROM FactInternetSales fis
INNER JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
INNER JOIN DimProductSubcategory dps ON dp.ProductSubcategoryKey = dps.ProductSubcategoryKey
INNER JOIN DimProductCategory dpc ON dps.ProductCategoryKey = dpc.ProductCategoryKey
INNER JOIN DimCustomer dc ON fis.CustomerKey = dc.CustomerKey
INNER JOIN DimSalesTerritory dst ON fis.SalesTerritoryKey = dst.SalesTerritoryKey
WHERE YEAR(fis.OrderDate) = 2013
````

#### Adding Profit Margin to the View
````sql
ALTER VIEW VENDAS_INTERNET AS
SELECT
    fis.SalesOrderNumber AS [ORDER_NO],
    fis.OrderDate AS [ORDER_DATE],
    dpc.EnglishProductCategoryName AS [PRODUCT_CATEGORY],
    dc.FirstName + ' ' + dc.LastName AS [CUSTOMER_NAME],
    dc.Gender AS [GENDER],
    dst.SalesTerritoryCountry AS [COUNTRY],
    fis.OrderQuantity AS [QUANTITY_SOLD],
    fis.TotalProductCost AS [SALE_COST],
    fis.SalesAmount AS [SALE_REVENUE],
    fis.SalesAmount - fis.TotalProductCost AS [PROFIT],
    CASE 
        WHEN fis.SalesAmount = 0 THEN 0
        ELSE ROUND((fis.SalesAmount - fis.TotalProductCost) / fis.SalesAmount, 2)
    END AS [PROFIT_MARGIN]
FROM FactInternetSales fis
INNER JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
INNER JOIN DimProductSubcategory dps ON dp.ProductSubcategoryKey = dps.ProductSubcategoryKey
INNER JOIN DimProductCategory dpc ON dps.ProductCategoryKey = dpc.ProductCategoryKey
INNER JOIN DimCustomer dc ON fis.CustomerKey = dc.CustomerKey
INNER JOIN DimSalesTerritory dst ON fis.SalesTerritoryKey = dst.SalesTerritoryKey
WHERE YEAR(fis.OrderDate) = 2013
````
### Analytical Queries
#### 1. Sales by Product Category
````sql
SELECT PRODUCT_CATEGORY, SUM(QUANTITY_SOLD) AS TOTAL_SALES
FROM VENDAS_INTERNET
GROUP BY PRODUCT_CATEGORY
````
#### 2. Revenue by Month
````sql
SELECT MONTH(ORDER_DATE) AS MONTH, SUM(SALE_REVENUE) AS TOTAL_REVENUE
FROM VENDAS_INTERNET
GROUP BY MONTH(ORDER_DATE)
ORDER BY MONTH
````
#### 3. Revenue and Cost by Country
````sql
SELECT COUNTRY, SUM(SALE_REVENUE) AS TOTAL_REVENUE, SUM(SALE_COST) AS TOTAL_COST
FROM VENDAS_INTERNET
GROUP BY COUNTRY
````
#### 4. Sales by Gender
````sql
SELECT GENDER, SUM(QUANTITY_SOLD) AS TOTAL_SALES
FROM VENDAS_INTERNET
GROUP BY GENDER
````
## Advanced Analysis
#### 1. Average Ticket per Customer
````sql
SELECT 
    CUSTOMER_NAME,
    COUNT(ORDER_NO) AS TOTAL_ORDERS,
    SUM(SALE_REVENUE) AS TOTAL_REVENUE,
    AVG(SALE_REVENUE) AS AVERAGE_TICKET
FROM VENDAS_INTERNET
GROUP BY CUSTOMER_NAME
ORDER BY AVERAGE_TICKET DESC
````
#### 2. Top-Selling Products
````sql
SELECT 
    dp.EnglishProductName AS PRODUCT,
    SUM(fis.OrderQuantity) AS TOTAL_QUANTITY
FROM FactInternetSales fis
JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
WHERE YEAR(fis.OrderDate) = 2013
GROUP BY dp.EnglishProductName
ORDER BY TOTAL_QUANTITY DESC
````
#### 3. Revenue vs. Cost by Category
````sql
SELECT 
    PRODUCT_CATEGORY,
    SUM(SALE_REVENUE) AS REVENUE,
    SUM(SALE_COST) AS COST,
    SUM(SALE_REVENUE - SALE_COST) AS PROFIT
FROM VENDAS_INTERNET
GROUP BY PRODUCT_CATEGORY
ORDER BY PROFIT DESC
````
`#### 4. Sales Distribution by Age Group
````sql
SELECT 
    CASE 
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 18 AND 25 THEN '18-25'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 26 AND 35 THEN '26-35'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 36 AND 50 THEN '36-50'
        ELSE '51+'
    END AS AGE_GROUP,
    COUNT(*) AS TOTAL_CUSTOMERS,
    SUM(fis.SalesAmount) AS TOTAL_REVENUE
FROM FactInternetSales fis
JOIN DimCustomer dc ON fis.CustomerKey = dc.CustomerKey
WHERE YEAR(fis.OrderDate) = 2013
GROUP BY 
    CASE 
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 18 AND 25 THEN '18-25'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 26 AND 35 THEN '26-35'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 36 AND 50 THEN '36-50'
        ELSE '51+'
    END
````
#### 5. Average Time Between Orders per Customer
````sql
WITH CustomerOrders AS (
    SELECT 
        CustomerKey,
        OrderDate,
        LAG(OrderDate) OVER (PARTITION BY CustomerKey ORDER BY OrderDate) AS PreviousOrder
    FROM FactInternetSales
    WHERE YEAR(OrderDate) = 2013
),
DaysBetween AS (
    SELECT 
        CustomerKey,
        DATEDIFF(DAY, PreviousOrder, OrderDate) AS DaysBetweenOrders
    FROM CustomerOrders
    WHERE PreviousOrder IS NOT NULL
)
SELECT 
    CustomerKey,
    AVG(DaysBetweenOrders) AS AVG_DAYS_BETWEEN_ORDERS
FROM DaysBetween
GROUP BY CustomerKey
````
#### 6. Top 5 Customers by Revenue
Table generated in Excel from the VENDAS_INTERNET view, connected via Power Query. Used to identify customers with the highest value per order.

Fields used:

- CUSTOMER_NAME

- TOTAL_ORDERS (Count of ORDER_NO)

- TOTAL_REVENUE (Sum of SALE_REVENUE)

- AVERAGE_TICKET (Total Revenue ÷ Total Orders)

Key Performance Indicators (KPI Cards)


#### 1. Total Revenue
````sql
SELECT SUM(SALE_REVENUE) AS TOTAL_REVENUE FROM VENDAS_INTERNET
````
#### 2. Total Cost
````sql
SELECT SUM(SALE_COST) AS TOTAL_COST FROM VENDAS_INTERNET
````
#### 3. Total Profit
````sql
SELECT SUM(SALE_REVENUE - SALE_COST) AS TOTAL_PROFIT FROM VENDAS_INTERNET
````
#### 4. Total Orders
````sql
SELECT COUNT(DISTINCT ORDER_NO) AS TOTAL_ORDERS FROM VENDAS_INTERNET
````
`#### 5. Average Ticket
````sql
SELECT 
    SUM(SALE_REVENUE) / COUNT(DISTINCT ORDER_NO) AS AVERAGE_TICKET
FROM VENDAS_INTERNET
````
### Dashboard Visualizations
Below are two examples of visualizations generated in Excel from the SQL Server integration:

#### 1. Report Page
![Dashboard](https://github.com/user-attachments/assets/b71a43b6-1f5c-4a2e-bf79-c9e7348f2446)




---

####  2. Report Page
![Dashboard](https://github.com/user-attachments/assets/2801479b-cd45-42d5-a003-7e582689fc25)


