<div align="center">

## Display Excel Spreadsheet Cell Data


</div>

### Description

This code snippet shows how to take an existing Excel 2000 document and display the cell data from different worksheets. This may work with other versions of Excel
 
### More Info
 
The following COM references are required:

Microsoft Excel 9.0 Object Library

Microsoft Office 9.0 Object Library


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lloyd Christmas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lloyd-christmas.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB\.NET
**Category**       |[Documents/ Frames](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/documents-frames__10-27.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lloyd-christmas-display-excel-spreadsheet-cell-data__10-1478/archive/master.zip)





### Source Code

```
'Open an Excel document, get a value from 3 cells, all on different worksheets
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet1 As Excel._Worksheet
    Dim xlSheet2 As Excel._Worksheet
    Dim xlSheet3 As Excel._Worksheet
    Dim mytext1 As String
    Dim mytext2 As String
    Dim mytext3 As String
    xlApp = CreateObject("Excel.Application")
    xlBook = xlApp.Workbooks.Open("c:\test.xls")
    xlSheet1 = xlBook.Worksheets(1)
    xlSheet2 = xlBook.Worksheets(2)
    xlSheet3 = xlBook.Worksheets(3)
    mytext1 = xlSheet1.Range("A1").Value
    mytext2 = xlSheet2.Range("A9").Value
    mytext3 = xlSheet3.Range("C1").Value
    MessageBox.Show(mytext1 + " " + mytext2 + " " + mytext3)
    xlBook.Close()
```

