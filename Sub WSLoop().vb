Sub WSLoop()

'   Set WS as a worksheet object variable

    Dim WS As Worksheet

    Dim Requires_Summary_Table_Header As Boolean
    Requires_Summary_Table_Header = False

    Dim COMMAND_SPREADSHEET As Boolean
    COMMAND_SPREADSHEET = True

'   Loop Through all worksheets

For Each WS In Worksheets

'   Set initial variable for holding ticker name

    Dim Ticker_Name As String
    Ticker_Name = " "

'   Set initial variable for holding total per ticker

    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0

'   Set new variables for Moderate Solution

    Dim Open_Price As Double
    Open_Price = 0

    Dim Close_Price As Double
    Close_Price = 0

    Dim Price_Change As Double
    Price_Change = 0

    Dim Percent_Change As Double
    Percent_Change = 0

'   Set new variables for Hard Solution

    Dim MIN_TICKER_NAME As String
    MIN_TICKER_NAME = " "

    Dim MAX_TICKER_NAME As String
    MAX_TICKER_NAME = " "

    Dim MIN_PERCENT As Double
    MIN_PERCENT = 0

    Dim MAX_PERCENT As Double
    MAX_PERCENT = 0

    Dim MAX_VOLUME_TICKER As String
    MAX_VOLUME_TICKER = " "

    Dim MAX_VOLUME As Double
    MAX_VOLUME = 0

'   Track location for Ticker Name in the Summary Table for the worksheet

    Dim Summary_Table_Row As Integer

    Summary_Table_Row = 2

'   Set row count for worksheet

    Dim Lastrow As Long
    Dim i As Long

    Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    If Requires_Summary_Table_Header Then

    '   Set titles for Summary Table
    
        WS.Range("I1").Value = "Ticker"
    
        WS.Range("J1").Value = "Yearly Change"
    
        WS.Range("K1").Value = "Percent Change"
    
        WS.Range("L1").Value = "Total Stock Volume"
    
    '   Set titles for Hard Solution
    
        WS.Range("O2").Value = "Greatest % Increase"
    
        WS.Range("O3").Value = "Greatest % Decrease"
    
        WS.Range("O4").Value = "Greatest Total Volume"
    
        WS.Range("P1").Value = "Ticker"
    
        WS.Range("Q1").Value = "Value"
    
    Else
    
        Requires_Summary_Table_Header = True
    
    End If

'   Set initial value of Open Price for first Ticker of WS

    Open_Price = WS.Cells(2, 3).Value

'   Loop from beginning of WS to last row

    For i = 2 To Lastrow
    
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1) Then
    
            Ticker_Name = WS.Cells(i, 1).Value
        
            Close_Price = WS.Cells(i, 6).Value
        
        '   Calculate change in price and percent
        
            Price_Change = Close_Price - Open_Price
        
        '   Checks division by 0
        
            If Open_Price <> 0 Then
        
                Percent_Change = (Price_Change / Open_Price) * 100
            
            Else
        
                MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Adjust <open> field and save the worksheet.")
       
            End If
        
        '   Add to the Ticker Name Total Volume
        
            Total_Ticker_Volume = Total_Ticker_Volume + WS.Cells(i, 7).Value
        
        '   Print Ticker Name in Summary Table in Column I
        
            WS.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        '   Print Price Change in Summary Table in Column J
        
            WS.Range("J" & Summary_Table_Row).Value = Price_Change
        
        '   Conditional formatting to fill color
        
            If (Price_Change > 0) Then
            
            '   Fill column with Green - positive
            
                WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf (Price_Change <= 0) Then
            
            '   Fill column with Red - negative
            
                WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
        '   Print Percent Change in Summary Table in Column K
        
            WS.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
        
        '   Print Total Ticker Volume in Summary Table in Column L
        
            WS.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
        
        '   Add 1 to the Summary Table row count
        
            Summary_Table_Row = Summary_Table_Row + 1
        
        '   Reset Price_Change and Percent_Change holders
        
            Price_Change = 0
        
            Close_Price = 0
        
        '   Next Ticker's Open_Price
        
            Open_Price = WS.Cells(i + 1, 3).Value
        
        ' Summary Table for Hard Solution
        
            If (Percent_Change > MAX_PERCENT) Then
            
                MAX_PERCENT = Percent_Change
            
                MAX_TICKER_NAME = Ticker_Name
            
            ElseIf (Percent_Change < MIN_PERCENT) Then
        
                MIN_PERCENT = Percent_Change
            
                MIN_TICKER_NAME = Ticker_Name
            
            End If
        
            If (Total_Ticker_Volume > MAX_VOLUME) Then
        
                MAX_VOLUME = Total_Ticker_Volume
            
                MAX_VOLUME_TICKER = Ticker_Name
            
            End If
        
        '   Reset counters
        
            Percent_Change = 0
        
            Total_Ticker_Volume = 0
        
    '   If cell in next row is the same Ticker Name, add to Total Ticker Volume
    
        Else
        
            Total_Ticker_Volume = Total_Ticker_Volume + WS.Cells(i, 7).Value
        
        End If
            
        Next i

    '   Check that it is not the first worksheet
    
            If Not COMMAND_SPREADSHEET Then
    
                WS.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
        
                WS.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
        
                WS.Range("P2").Value = MAX_TICKER_NAME
        
                WS.Range("P3").Value = MIN_TICKER_NAME
        
                WS.Range("Q4").Value = MAX_VOLUME
        
                WS.Range("P4").Value = MAX_VOLUME_TICKER
        
            Else
    
                COMMAND_SPREADSHEET = False
        
            End If
    
    Next WS
    
End Sub
