# 📊 Projeto de Integração SQL Server + Excel  
**Base de Dados: AdventureWorks DW 2014**

## 🚀 Objetivo

Desenvolver uma integração entre SQL Server e Excel para análise de indicadores de vendas online, utilizando dados do banco AdventureWorks DW 2014. O foco será em métricas de desempenho para o ano de 2013.

---

## 📥 Instalação do Banco de Dados

Para instalar o banco AdventureWorks DW 2014, siga as instruções oficiais da Microsoft:  
🔗 [AdventureWorks - Instalação e Configuração](https://docs.microsoft.com/pt-br/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms)

---

## 📌 Indicadores Definidos

1. **Total de Vendas Internet por Categoria do Produto**  
2. **Receita Total Internet por Mês do Pedido**  
3. **Receita e Custo Total Internet por País**  
4. **Total de Vendas Internet por Sexo do Cliente**

> ⚠️ **Ano de análise:** 2013

---

## 🗃️ Tabelas Utilizadas

| Tabela               | Descrição                          |
|----------------------|-------------------------------------|
| `FactInternetSales`  | Fatos de vendas pela internet       |
| `DimCustomer`        | Dados demográficos dos clientes     |
| `DimSalesTerritory`  | Informações geográficas de vendas   |
| `DimProductCategory` | Categoria dos produtos              |
| `DimProductSubcategory` | Subcategoria dos produtos        |
| `DimProduct`         | Detalhes dos produtos               |

---

## 🧱 Estrutura da View `VENDAS_INTERNET`

```
CREATE VIEW VENDAS_INTERNET AS
SELECT
    fis.SalesOrderNumber AS [Nº_PEDIDO],
    fis.OrderDate AS [DATA_PEDIDO],
    dpc.EnglishProductCategoryName AS [CATEGORIA_PRODUTO],
    dc.FirstName + ' ' + dc.LastName AS [NOME_CLIENTE],
    dc.Gender AS [SEXO],
    dst.SalesTerritoryCountry AS [PAÍS],
    fis.OrderQuantity AS [QTD_VENDIDA],
    fis.TotalProductCost AS [CUSTO_VENDA],
    fis.SalesAmount AS [RECEITA_VENDA]
FROM FactInternetSales fis
INNER JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
INNER JOIN DimProductSubcategory dps ON dp.ProductSubcategoryKey = dps.ProductSubcategoryKey
INNER JOIN DimProductCategory dpc ON dps.ProductCategoryKey = dpc.ProductCategoryKey
INNER JOIN DimCustomer dc ON fis.CustomerKey = dc.CustomerKey
INNER JOIN DimSalesTerritory dst ON fis.SalesTerritoryKey = dst.SalesTerritoryKey
WHERE YEAR(fis.OrderDate) = 2013

```
## Incluindo Margem de Vendas na View
```
ALTER VIEW VENDAS_INTERNET AS
SELECT
    fis.SalesOrderNumber AS [Nº_PEDIDO],
    fis.OrderDate AS [DATA_PEDIDO],
    dpc.EnglishProductCategoryName AS [CATEGORIA_PRODUTO],
    dc.FirstName + ' ' + dc.LastName AS [NOME_CLIENTE],
    dc.Gender AS [GENERO],
    dst.SalesTerritoryCountry AS [PAÍS],
    fis.OrderQuantity AS [QTD_VENDIDA],
    fis.TotalProductCost AS [CUSTO_VENDA],
    fis.SalesAmount AS [RECEITA_VENDA],
    fis.SalesAmount - fis.TotalProductCost AS [MARGEM],
    CASE 
        WHEN fis.SalesAmount = 0 THEN 0
        ELSE ROUND((fis.SalesAmount - fis.TotalProductCost) / fis.SalesAmount, 2)
    END AS [MARGEM_%]
FROM FactInternetSales fis
INNER JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
INNER JOIN DimProductSubcategory dps ON dp.ProductSubcategoryKey = dps.ProductSubcategoryKey
INNER JOIN DimProductCategory dpc ON dps.ProductCategoryKey = dpc.ProductCategoryKey
INNER JOIN DimCustomer dc ON fis.CustomerKey = dc.CustomerKey
INNER JOIN DimSalesTerritory dst ON fis.SalesTerritoryKey = dst.SalesTerritoryKey
WHERE YEAR(fis.OrderDate) = 2013


---
```
## 📈 Consultas Analíticas

### 1. Vendas por Categoria de Produto
```
SELECT CATEGORIA_PRODUTO, SUM(QTD_VENDIDA) AS TOTAL_VENDAS
FROM VENDAS_INTERNET
GROUP BY CATEGORIA_PRODUTO
```
### 2. Receita por Mês
```
SELECT MONTH(DATA_PEDIDO) AS MES, SUM(RECEITA_VENDA) AS RECEITA_TOTAL
FROM VENDAS_INTERNET
GROUP BY MONTH(DATA_PEDIDO)
ORDER BY MES
```
### 3. Receita e Custo por País
```
SELECT PAÍS, SUM(RECEITA_VENDA) AS RECEITA_TOTAL, SUM(CUSTO_VENDA) AS CUSTO_TOTAL
FROM VENDAS_INTERNET
GROUP BY PAÍS

```
### 4. Vendas por Sexo
```
SELECT SEXO, SUM(QTD_VENDIDA) AS TOTAL_VENDAS
FROM VENDAS_INTERNET
GROUP BY SEXO

```
## 🔍 Análises Avançadas 

### 1. Ticket Médio por Cliente
```
SELECT 
    NOME_CLIENTE,
    COUNT(Nº_PEDIDO) AS TOTAL_PEDIDOS,
    SUM(RECEITA_VENDA) AS RECEITA_TOTAL,
    AVG(RECEITA_VENDA) AS TICKET_MEDIO
FROM VENDAS_INTERNET
GROUP BY NOME_CLIENTE
ORDER BY TICKET_MEDIO DESC

```
### 2. Produtos Mais Vendidos
```
SELECT 
    dp.EnglishProductName AS PRODUTO,
    SUM(fis.OrderQuantity) AS QTD_TOTAL
FROM FactInternetSales fis
JOIN DimProduct dp ON fis.ProductKey = dp.ProductKey
WHERE YEAR(fis.OrderDate) = 2013
GROUP BY dp.EnglishProductName
ORDER BY QTD_TOTAL DESC

```
### 3. Comparativo Receita vs. Custo por Categoria
```
SELECT 
    CATEGORIA_PRODUTO,
    SUM(RECEITA_VENDA) AS RECEITA,
    SUM(CUSTO_VENDA) AS CUSTO,
    SUM(RECEITA_VENDA - CUSTO_VENDA) AS MARGEM
FROM VENDAS_INTERNET
GROUP BY CATEGORIA_PRODUTO
ORDER BY MARGEM DESC

```
### 4. Distribuição de Vendas por Faixa Etária
```
SELECT 
    CASE 
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 18 AND 25 THEN '18-25'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 26 AND 35 THEN '26-35'
        WHEN YEAR(GETDATE()) - YEAR(dc.BirthDate) BETWEEN 36 AND 50 THEN '36-50'
        ELSE '51+'
    END AS FAIXA_ETARIA,
    COUNT(*) AS TOTAL_CLIENTES,
    SUM(fis.SalesAmount) AS RECEITA_TOTAL
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


```
### 5. Tempo Médio entre Pedidos por Cliente
```

WITH PedidoCliente AS (
    SELECT 
        CustomerKey,
        OrderDate,
        LAG(OrderDate) OVER (PARTITION BY CustomerKey ORDER BY OrderDate) AS PedidoAnterior
    FROM FactInternetSales
    WHERE YEAR(OrderDate) = 2013
),
DiferencaDias AS (
    SELECT 
        CustomerKey,
        DATEDIFF(DAY, PedidoAnterior, OrderDate) AS DiasEntrePedidos
    FROM PedidoCliente
    WHERE PedidoAnterior IS NOT NULL
)
SELECT 
    CustomerKey,
    AVG(DiasEntrePedidos) AS MEDIA_DIAS_ENTRE_PEDIDOS
FROM DiferencaDias
GROUP BY CustomerKey
```
🏅 Tabela Dinâmica: Ranking de Clientes
Tabela gerada no Excel a partir da view VENDAS_INTERNET, conectada via Power Query. Utilizada para identificar os clientes com maior valor agregado por pedido.

📊 Campos utilizados:
NOME_CLIENTE
TOTAL_PEDIDOS (Contagem de Nº_PEDIDO)
RECEITA_TOTAL (Soma de RECEITA_VENDA)
TICKET_MEDIO (Receita Total ÷ Total de Pedidos)
```
```
## 🧾 Indicadores-Chave (KPI Cards)
### 🎯 1. Receita Total
```
SELECT SUM(RECEITA_VENDA) AS RECEITA_TOTAL FROM VENDAS_INTERNET

```
### 🧮 2. Custo Total
```
SELECT SUM(CUSTO_VENDA) AS CUSTO_TOTAL FROM VENDAS_INTERNET

```
### 📊 3. Margem Total
```
SELECT SUM(RECEITA_VENDA - CUSTO_VENDA) AS MARGEM_TOTAL FROM VENDAS_INTERNET

```
### 📦 4. Total de Pedidos
```
SELECT COUNT(DISTINCT Nº_PEDIDO) AS TOTAL_PEDIDOS FROM VENDAS_INTERNET

```
### 🧾 5. Ticket Médio
```
SELECT 
    SUM(RECEITA_VENDA) / COUNT(DISTINCT Nº_PEDIDO) AS TICKET_MEDIO
FROM VENDAS_INTERNET


```
## 🔄 Atualização de Dados e Integração com Excel
Exemplo de atualização de dados via transação SQL:
```
BEGIN TRANSACTION T1
    UPDATE FactInternetSales
    SET OrderQuantity = 20
    WHERE ProductKey = 361 -- Categoria Bike
````
### A view VENDAS_INTERNET pode ser conectada ao Excel via Power Query ou conexão ODBC para criação de dashboards dinâmicos.

## 🧠 Autor(a)
Débora Rebula Klein - Projeto acadêmico e exploratório para fins de aprendizado em BI e integração de dados.

### 📎 Licença
Este projeto é de uso livre para fins educacionais. AdventureWorks é uma base de dados pública fornecida pela Microsoft.
```
