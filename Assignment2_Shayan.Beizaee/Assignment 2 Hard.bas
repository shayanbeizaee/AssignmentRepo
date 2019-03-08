Attribute VB_Name = "Module1"
Sub Alphabetic()

For Each ws In Worksheets
    ws.Select
    
Dim count As Integer
Dim coun As Double
Dim num As Integer
Dim sum As Double
Dim opens As Double
Dim clos As Double
Dim diff As Double
Dim percen As Double
Dim greatinc As Double
Dim greatdec As Double
Dim greattotal As Double
Dim tickerinc As String
Dim tickerdec As String
Dim tickertotal As String

count = 2
num = 2
coun = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    lRow = Cells(Rows.count, 1).End(xlUp).Row
     
     For i = 1 To lRow
        If Cells(i + 1, 1) <> Cells(i, 1) Then
        Cells(count, 9).Value = Cells(i + 1, 1)
        count = count + 1
        End If
    Next i

    sum = Cells(2, 7)
    For i = 2 To lRow
        If Cells(i, 1) = Cells(i + 1, 1) Then
        sum = sum + Cells(i + 1, 7).Value
        Else
        Cells(num, 12) = sum
        num = num + 1
        sum = Cells(i + 1, 7).Value
        End If
    Next i
    
    For i = 2 To lRow
        If Cells(i, 1) <> Cells(i - 1, 1) Then
        opens = Cells(i, 3).Value
        ElseIf Cells(i, 1) <> Cells(i + 1, 1) Then
        clos = Cells(i, 6).Value
               
        diff = clos - opens
        Cells(coun, 10).Value = diff
        
        
            If opens <> 0 Then
            percen = diff / opens
            Cells(coun, 11).Value = percen
            ElseIf opens = 0 Then
            Cells(coun, 11).Value = 0
            End If
        Cells(coun, 11).NumberFormat = "0.00%"
        coun = coun + 1
        opens = 0
        clos = 0
        End If
    Next i
        
        
    For i = 2 To lRow
        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    lRow2 = Cells(Rows.count, 11).End(xlUp).Row
    greattotal = Cells(2, 12).Value
    greatinc = Cells(2, 11).Value
    greatdec = 0
    
    For i = 2 To lRow2
    
        If Cells(i + 1, 11) > greatinc Then
        greatinc = Cells(i + 1, 11).Value
        tickerinc = Cells(i + 1, 9).Value
        End If
        
        If Cells(i + 1, 11) < greatdec Then
        greatdec = Cells(i + 1, 11).Value
        tickerdec = Cells(i + 1, 9).Value
        End If
        
        If Cells(i + 1, 12) > greattotal Then
        greattotal = Cells(i + 1, 12).Value
        tickertotal = Cells(i + 1, 9).Value
        End If
       
    Next i
       
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 16).Value = tickerinc
    Cells(3, 16).Value = tickerdec
    Cells(4, 16).Value = tickertotal
    Cells(2, 17).Value = greatinc
    Cells(3, 17).Value = greatdec
    Cells(4, 17).Value = greattotal
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
   
    
    
Next ws



End Sub


