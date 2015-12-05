Xls2Sqlite
==========
Reads an Excel file and creates an SQLite file containing the data from the Excel file.

* One table is created for each Excel sheet.
* Empty rows are ignored
* The first row of the Excel file must contain the column names.
* Column names beginning with # are ignored.


Limitations:
-----------
This is a very simple implementation, with the following limitations:

* No primary keys or foreign keys are created.
* No autoincrement or unique indexes are created.
* No NOT NULL constraints (or any constraints for that matter) are created.
* All fields in the Excel file are treated as strings.
* All columns in the SQLite file are created as TEXT.
* If the SQLite file exists already, it is deleted before creating a new one.

Usage:
-----
Build the program:

```
mvn clean package
```

Run the program:

```
java -jar target/xls2sqlite-0.0.1.jar /path/to/file.xls /path/to/file.db
```

