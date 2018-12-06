Attribute VB_Name = "Module1"
Sub homework_2_test()
'Create a script that will loop through each year of stock data
' and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
' Your result should look as follows (note: all solution images are for 2015 data).
'[easy_solution](Images/easy_solution.png)

' Create a loop that goes through each row totaling the volume and placing the ticker in a seperate chart

'Determine the Last Row
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim total_volume As Double
Dim ticker_header As String
'Dim Sheet1, sheet2, sheet3, Sheet4, sheet5, sheet6, sheet7 As Worksheet
Dim Year_volume As String
Dim yearly_change As Double
Dim percent_change As Integer


Dim ws As Worksheet, wsCollection As Sheets
Set wsCollection = Sheets 'Get entire collection of Worksheets
Set ws = Sheets(1) 'Get first Worksheet in ActiveWorkbook
'Set ws = Sheets("Sheet(1)") 'Get Worksheet named "Sheet1" in ActiveWorkbook

total_volume = 0

'Set Sheet1 = Worksheets("A")
'Set sheet2 = Worksheets("B")
'Set sheet3 = Worksheets("C")
'Set Sheet4 = Worksheets("D")
'Set sheet5 = Worksheets("E")
'Set sheet6 = Worksheets("F")
'Set sheet7 = Worksheets("P")

'Dim res As String, ws As Worksheet
 ' res = "Selected sheets:" & vbNewLine
  'Sheets(1).Select
  
'Sheets(1).Select
'Call Sheets(2).Select(False)
'Call Sheets(3).Select(False)
'Call Sheets(4).Select(False)
'Call Sheets(5).Select(False)
'Call Sheets(6).Select(False)
'Call Sheets(7).Select(False)

For Each ws In Worksheets
'Determine the Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
' Grabbed the WorksheetName
    'WorksheetName = ws.Name
        'MsgBox WorksheetName
        
'Set Headers for chart
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'   Keep track of the location for each ticker in the summary table
     Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2
  
' Loop through each row
     For i = 2 To lastrow
    
'Cells(i, 3) = Year_volume
'Left(Str(Cells(i, 3)), 4) = Year_volume
'MsgBox (Year_volume)

' If a column contains the same thing as current cell add together volumes
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Add to total volume & print ticker
        ticker_header = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        
' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker_header
      
 ' Print the volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      total_volume = 0
' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Volume Total
      total_volume = total_volume + ws.Cells(i, 7).Value

    End If
' Once we've completed all rows, print the value in the total column
 ' MsgBox (Year("B2"))
    Next i
    
Next ws

End Sub
