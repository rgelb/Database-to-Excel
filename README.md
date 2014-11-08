Database to Excel
======================

This is a quick and dirty command line program to run 1 or more queries against a SQL Server database and then save each result to its own sheet within Excel.

Example command
```sh
DatabaseToExcel.exe /s:MySqlBox /d:MyLocalDb /u:foo /p:bar /q:query.sql /o:out.xlsx
```

This is the query that we run:
```sh
  SELECT  TOP 10 ProposalID
  INTO    #pidsToGo
  FROM    dbo.Proposal
  WHERE   CurrentWorkflowState IN (10, 11, 20, 60, 70)
  
  -- Query 1
  SELECT  p.ProposalID,
          p.ProposalTitle,
          p.BRAsOfDate
  FROM    dbo.Proposal p  
  WHERE   p.ProposalID IN (SELECT ProposalID
                           FROM   #pidsToGo)
                           
  -- Query 2
  SELECT  pl.ProposalLineId,
          pl.UIIndex,
          pl.ProposalID,
          pl.SpotLength
  FROM    dbo.ProposalLine pl
  WHERE   pl.ProposalID IN (SELECT    ProposalID
                            FROM      #pidsToGo)                             
```
  
The application then produces an Excel file with 2 worksheets.  

![alt tag](https://raw.github.com/rgelb/Database-to-Excel/master/Extras/Images/ExcelSimple.png)

But we can make it prettier with named sheets.  To that end we can provide a named sheets file.  Example:

```sh
Proposals
Proposal Lines
```

Then run the command again with the sheetFile parameter:

```sh
DatabaseToExcel.exe /s:MySqlBox /d:MyLocalDb /u:foo /p:bar /q:query.sql /o:out.xlsx /sheetFile:sheetNames.txt 
```

And then, the app generates an Excel file with named sheets:

![alt tag](https://raw.github.com/rgelb/Database-to-Excel/master/Extras/Images/ExcelMoBetta.png)

Other command line switches:
```sh
/i - Login via Windows Authentication
/e - Launch file in Excel after creation
```



