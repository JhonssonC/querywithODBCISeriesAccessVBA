# querywithODBCISeriesAccessVBA
Interface between Excel(VBA) and DSN connection ISeries Access ODBC Driver for obtaining data through Query.


Execution Test:

![Imgur](https://i.imgur.com/qdLrgvZ.gif)

* Prerequisites (only for windows 7 - 10):

Office 2012-2019 (32/64 bit)
Personal Communications iSeries Access for Windows

Instructions:

* Open or create a macro-enabled Excel file.
* Create a table of contents to a specific sheet called VAR.

Important note: Specify Query Select in B1, the reference code for each select is '[[CODE]]', and also add username and password for the connection in cells E1 and E2 respectively.

![Imgur2](https://i.imgur.com/JPWxF55.png)

SQL 

```SQL
...
WHERE
C.CLCODE = '[[CODE]]'
...
```

* In another Sheet (it can be Sheet1) build the following table in an empty sheet paying special attention to the columns specified in the VAR sheet (Column A) in the previous step the columns must match the headers, not textually but they must be the data that was specified in the VAR sheet.

![Imgur3](https://i.imgur.com/VWyjiod.png)

* Check that the system data access (ODBC) has the following settings. (see code)

![Imgur4](https://i.imgur.com/iZ5JITV.png)

VBA 

```VBA
...
USERNAME = Sheets("VAR").Range("E1")
PASSWORD = Sheets("VAR").Range("E2")
conn.ConnectionString = "dsn=SICOPUB-MAN;User Id=" & USERNAME & ";Password=" & PASSWORD & " ;"
...
```

Note: The data to be searched are the codes, these are taken as a reference to locate the rest of the data based on the Query specified in the VAR Sheet applied to each code.

Enter codes to search for, select the codes in the table and execute the macro.

![Imgur](https://i.imgur.com/qdLrgvZ.gif)

Note: The selection can be one or several elements and it also supports elements only from a specified filter (previously the table data must be filtered in excel and the macro will only be executed on the selection without considering hidden rows).
