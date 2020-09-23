<div align="center">

## Convert your ODBC connections to OLE DB


</div>

### Description

Looking for faster performance? If you have an older database driven ASP app, it probably uses an ODBC DSN in its connection string to reach the database. It probably looks alot like this: "DSN=myDSNName;".

If you see this you should immediately upgrade the code to ADO/OLEDB--Microsoft's new standard. Not only will this help you keep the code current, but it will run faster and take up less memory.

Below are the connection strings for OLEDB/ADO for both Access and SQL Server--the two most common databases for IIS platforms. If you are using another database, you'll need to consult your db's ADO Provider's documentation for the proper connection string.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-convert-your-odbc-connections-to-ole-db__4-35/archive/master.zip)





### Source Code

```
'Access Example
"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\MyDatabase.mdb;"
'SQL Server example
"Provider=SQLOLEDB; Data Source=MySQLServer; Initial Catalog=pubs; User Id=MyUserId; Password=MyPassword;"
```

