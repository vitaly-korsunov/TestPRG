# DbToExcelExporter

Small C# console app that runs a SQL Server stored procedure and saves the first result set as an Excel `.xlsx` file.

## Run

```powershell
dotnet run -- --connection "Server=.;Database=MyDb;Trusted_Connection=True;TrustServerCertificate=True" --sp dbo.ExportCustomers --out customers.xlsx
```

With stored procedure parameters:

```powershell
dotnet run -- -c "Server=.;Database=MyDb;Trusted_Connection=True;TrustServerCertificate=True" -s dbo.ExportOrders -o orders.xlsx -p StartDate=2026-01-01 -p EndDate=2026-06-29
```

## Options

- `--connection`, `-c`: SQL Server connection string.
- `--sp`, `-s`: Stored procedure name.
- `--out`, `-o`: Output Excel file path.
- `--sheet`: Excel sheet name. Default is `Data`.
- `--timeout`: SQL command timeout in seconds. Default is `120`.
- `--param`, `-p`: Stored procedure parameter as `Name=Value`. Can be repeated.

## Example Stored Procedure

```sql
CREATE OR ALTER PROCEDURE dbo.ExportCustomers
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        CustomerId,
        CustomerName,
        Email,
        CreatedAt
    FROM dbo.Customers
    ORDER BY CustomerName;
END;
```

