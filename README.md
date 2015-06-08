# powershell-tools
powershell script for many applications
## Dependencies
* [ACE.OLEDB provider](http://www.microsoft.com/en-us/download/details.aspx?id=13255)

## Functions:

### Get-ExcelData (by [Martin Schvartzman](http://blogs.technet.com/b/pstips/archive/2014/06/02/get-excel-data-without-excel.aspx))

Then, you can use it to get the entire default worksheet (Sheet1): 

```posh
Get-ExcelData -Path C:\myFiles\Users.xlsx
```

Or to get a specific worksheet:

```posh
Get-ExcelData -Path C:\myFiles\Users.xlsx -WorksheetName 'Sheet2'
```

Or by specifying a query:

```posh
Get-ExcelData -Path C:\myFiles\Users.xlsx -Query 'SELECT TOP 3 * FROM Sheet3'
```

Or an even more complex query:

```posh
Get-ExcelData -Path C:\myFiles\Users.xlsx -Query "SELECT GivenName, Surname, City, State FROM Sheet1 WHERE State in ('CA','WA')"
```