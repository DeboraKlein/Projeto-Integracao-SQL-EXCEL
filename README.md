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

```sql
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
