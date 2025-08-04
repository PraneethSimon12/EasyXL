import pyodbc
import pandas as pd

# Replace these
server = 'DESKTOP-1234\\SQLEXPRESS'  # Double backslashes
database = 'OnlineRetailDb'

# Connection string with Trusted_Connection=yes for Windows auth
conn = pyodbc.connect(
    f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'
)

query = "SELECT TOP 10 * FROM dbo.Customers"  # Example query
df = pd.read_sql(query, conn)

print(df)
conn.close()
