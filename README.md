# ExcelInterface
Simple Excel interface allows you to read a WorkSheet and write a new Excel file.

## How to Read a WorkSheet

Using GetExcelRowEnumerator it will return a IEnumerable with the entire contents of the sheet
```cs
public static IEnumerable<List<dynamic>> GetExcelRowEnumerator(int sheetNumber, string path, int RowStart = 1)
```
Example:
```cs
var ExcelSheetOne = Interface.GetExcelRowEnumerator(1, @"c:\ExcelFile.xlsx");
```
It reads the number one card of "Excel File.xlsx" and return a list.

## How to write an Excel spreadsheet

Using SetExcelRow passing as argument a list of a class, a new Excel will open with the printed list
```cs
public static void SetExcelRow<T>(List<T> data)
```
Example:
```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWs;


namespace Test
{
    class Program {
        static void Main(string[] args) {
            List<Person> Users = new List<Person>();

            for(int i = 0; i < 10; i++) {
                Users.Add(new Person() {
                    Id = i,
                    Name = "Claudio"
                });
            }
            
            Interface.SetExcelRow<Person>(Users);
        }

        public class Person {
            public int Id { get; set; }
            public string Name { get; set; }
        }
    }
}
```
