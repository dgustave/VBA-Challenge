Attribute VB_Name = "Module11"
Sub Stock_Diagnosis()

'Define all Stock compenents and setup
Dim ws As Worksheet 'ws as a worksheet object variable.
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim vol As Long
Dim Summary_Table_Row As Long
Dim Final_Row As Long
Dim i As Long
Dim Summary_Table_Headers As Boolean
    Summary_Table_Headers = False       'Set Header flag
    
'Set variables for stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
Dim max_ticker As String
Dim min_ticker As String
Dim max_ticker_volume As String
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Double

'Set values for "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
max_ticker = " "
min_ticker = " "
max_ticker_volume = " "
max_percent = 0
min_percent = 0
max_volume = 0

' overflow error
On Error Resume Next


'Loop for all worksheets in the active workbock not specific
For Each ws In Worksheets

    'setup rows and values for loop
    Summary_Table_Row = 2
    Final_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    year_open = 0
    year_close = 0
    vol = 0
    
    If Summary_Table_Headers Then
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set Additional Titles for new Summary Table on the right for current worksheet
    ws.Cells(1, 15).Value = "Greatest % Increase"
    ws.Cells(1, 16).Value = "Greatest % Decrease"
    ws.Cells(1, 17).Value = "Greatest Total Volume"
    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"
    
    Else
        'This is the first, resulting worksheet, reset flag for the rest of worksheets
        Summary_Table_Headers = True
    End If
    
    'Set initial value of Open Price for the first Ticker of CurrentWs,
    ' The rest ticker's open price will be initialized within the for loop below
    year_open = ws.Cells(2, 3).Value

    For i = 2 To Final_Row
        
        'If cell values don't match perform calculations and assign values
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
            'Ticker name data will be inserted
            ticker = Cells(i, 1).Value
                
            'calculations
            
            vol = vol + ws.Cells(i, 7).Value
            year_close = Cells(i, 6).Value
            yearly_change = year_close - year_open
                
                ' Not divisble by zero error
                If year_open <> 0 Then
                
                    percent_change = ((yearly_change) / year_open)
                Else
                    percent_change = 0
    
                End If
                
            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset Delta_rice and Delta_Percent holders, as we will be working with new Ticker
            yearly_change = 0
            ' Hard part,do this in the beginning of the for loop Delta_Percent = 0
            Close_Price = 0
            ' Capture next Ticker's Open_Price
            year_open = ws.Cells(i + 1, 3).Value
                
              If (yearly_change > 0) Then
                    'Fill column with GREEN color - good
                    ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
                ElseIf (yearly_change <= 0) Then
                    'Fill column with RED color - bad
                    ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
                End If

            'reset volume for new ticker
            vol = 0
              
        Else
            ' Increase the Total Ticker Volume
            vol = vol + ws.Cells(i, 7).Value
        End If
   
    Next i
            
    'Column K percent format
    ws.Columns("K").NumberFormat = "0.00%"
         

    
   
Next ws

End Sub

