# VisualBasic
This an an example of how to use Visual Basic in Excel to take specfic strings and mtach them to a number based on their category.
This example will be using crime data and the code removes columns that are not needed, adds a column for the wighted numbers to populate based on a VLOOKUP reference table.

# Below shows the crimes and their weighed numbers 
![Total Employees](https://github.com/kbvss/VisualBasic/blob/main/CategoryAndWeight.png?raw=true)

# This is the data I started out with
![Total Employees](https://github.com/kbvss/VisualBasic/blob/main/Data.PNG?raw=true)


# Here is the breakdown of the code


    Sub WeightedNumber()
  
    Dim currentColumn As Integer
    Dim columnHeading As String

    For currentColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

    'Check whether to keep the column or delete
        Select Case columnHeading
        'Titles of the columns to keep
            Case "Report Number", "Date Reported", "Crime", "location"
                'Do nothing
            Case Else
                'Delete the column
                ActiveSheet.Columns(currentColumn).Delete
            End Select
        Next


    'Insert a column for weight
     Columns("D:D").Select
        Selection.Insert Shift:=xlToLeft
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Weight"


    'Create a blank sheet after the data sheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(Type:=xlWorksheet, After:=Application.ActiveSheet)

    'Selecting the new sheet
    Sheets("Sheet1").Select

    'Adding text of the crimes and their coorosponding values
    'This sheet will be referenced on the main sheet using VLOOKUP in the weigh column

    'Crimes reference column
    Range("A1").Value = "Crimes:" & Row4Num
    Range("A2").Value = "Murder" & Row4Num
    Range("A3").Value = "Self Defense" & Row4Num
    Range("A4").Value = "Assault" & Row4Num
    Range("A5").Value = "Robbery Person" & Row4Num
    Range("A6").Value = "Business Robbery" & Row4Num
    Range("A7").Value = "Trespassing" & Row4Num
    Range("A8").Value = "Theft" & Row4Num
    Range("A9").Value = "Vehicle Theft" & Row4Num
    Range("A10").Value = "Shoplifting" & Row4Num
    Range("A11").Value = "Fraud" & Row4Num


            'Crimes weight column values
            Range("B1").Value = "Weight:" & Row4Num
            Range("B2").Value = "4" & Row4Num
            Range("B3").Value = "4" & Row4Num
            Range("B4").Value = "3" & Row4Num
            Range("B5").Value = "3" & Row4Num
            Range("B6").Value = "2" & Row4Num
            Range("B7").Value = "2" & Row4Num
            Range("B8").Value = "1" & Row4Num
            Range("B9").Value = "1" & Row4Num
            Range("B10").Value = "1" & Row4Num
            Range("B11").Value = "1" & Row4Num


    'Inserting VLookup function to extract the values from the reference table
    Sheets("Weighted example Data").Select
    With Range("D2:D" & Range("C" & Rows.Count).End(xlUp).Row) 'Autofill
     .Formula = "=VLookup(C2,Sheet1!$A$1:$B$11,2,False)"
     .Value = .Value
     End With


    'Get rid of the decimal places
    Columns("D:D").Select
        Selection.NumberFormat = "0.00"
        Selection.NumberFormat = "0.0"
        Selection.NumberFormat = "0"


    End Sub




